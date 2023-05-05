// See https://aka.ms/new-console-template for more information
using BatchLoadTest;


try
{
    Worker worker = new Worker();
    worker.Load("emails");
}catch(Exception ex)
{
    Console.WriteLine(ex.ToString());
}
Console.WriteLine("done...press any key to quit");
Console.ReadLine();