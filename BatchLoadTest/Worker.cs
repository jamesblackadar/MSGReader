using MsgReader;
using RtfPipe.Tokens;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace BatchLoadTest
{
    internal class Worker
    {
        string pathToCSV = "";
        string pathToLog = "";
        string pathToOutPutCSV = "";
        public void Load(string folder)
        {
            if(!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            pathToCSV = Path.Combine(folder, "emails.tsv");
            pathToLog = Path.Combine(folder, "emaillog." + DateTime.Now.ToString("ddHHmmss") + ".txt");
            pathToOutPutCSV = Path.Combine(folder, "emails." + DateTime.Now.ToString("ddHHmmss") + ".tsv");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Log("starting");
           
            if (!File.Exists(pathToCSV))
            {
                throw new FileNotFoundException($"{pathToCSV}' tab delimited file didn't exist. It should have minimum two columns where the first col is status of email(error message etc), and second col is the path to the msg file");
            }
            var lines = File.ReadAllLines(pathToCSV);
            Log("loaded " + lines.Count());
          
            
            foreach (var line in lines)
            {  
                var elements = line.Split("\t");
                if(elements.Length >1) {
                    Test(elements[1], elements[0]);
                }
                else
                {
                    Log("error: " + line);
                }
            }
            Log("done");
        }
        private void Test(string fileName,string oldError)
        {
            FileInfo f = new FileInfo(fileName);
            string state = "";
            string message = "";
            try
            {
                var msgReader = new Reader();                
                if(f.Exists)
                {
                   if(f.Length> 36000000)
                    {
                        state = "Larger than desired size threshold. ";message = f.Length.ToString() + "b";
                        return;
                    }
                    else
                    {
                        using (var msg = new MsgReader.Outlook.Storage.Message(fileName))
                        {
                            var from = msg.Sender;
                            var subject = msg.Subject;
                            var body = msg.BodyText;
                            Log(string.Join("", string.Join("", subject?.Take(100) ?? "null")));
                            state = "OK";
                            message = string.Join("", subject?.Take(10)??"blanksubject") + "body:" + string.Join("", body?.ToString().Where(c => char.IsLetterOrDigit(c)).Take(100)??"blankbody");
                        }
                    }
                }
                else
                {
                    state = "NOT FOUND";
                    message = "didn't exist";
                }
              
            }
            catch(Exception ex)
            {
                Log(ex.ToString());  
                if(f.Exists )
                {
                    if (f.Directory.Name != "badmaillocalcopy")
                    {
                        var dest = Path.Combine("badmaillocalcopy", f.Name);
                        if (!Directory.Exists("badmaillocalcopy"))Directory.CreateDirectory("badmaillocalcopy");
                        if(!File.Exists(dest)) f.CopyTo(dest);
                        Test(dest, "trylocal-"+ oldError);
                        state = "COPYToLocalDrive";
                    }
                    else
                    {
                        new Process
                        {
                            StartInfo = new ProcessStartInfo(fileName)
                            {
                                UseShellExecute = true
                            }
                        }.Start();
                        Console.WriteLine("If it opened OK in Outlook press 'y' and {enter}");
                        var result = Console.ReadLine();
                        if (result?.ToUpperInvariant() == "Y")
                        {
                            state = "ERR-butopenedOKinoutlook";
                            message= ex.Message;
                        }
                        else
                        {
                            state = "ERR-CORRUPT";
                            message = "didn't open: " + result;
                        }
                    }
                }                
            }finally
            {
                CSV(oldError, fileName, state ,message);
            }
        }
        private void CSV(string state,string message,string oldError,string newError )
        {
            File.AppendAllText(pathToOutPutCSV, string.Join("\t", state,message.Replace("\"",""), string.Join("", oldError.Replace("\"", "").Take(50)), string.Join("", newError.Replace("\"", "").Take(200)) + "\r\n"));
        }
        private void Log(string message)
        {
            Console.WriteLine(message);
            File.AppendAllText(pathToLog, message);
        }
    }
}
