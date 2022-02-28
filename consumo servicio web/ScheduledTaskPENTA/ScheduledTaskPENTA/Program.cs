using System;
using System.IO;
using System.Net;
using System.Timers;

namespace ScheduledTaskPENTA
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Timers.Timer tmrDelay;
            Console.WriteLine("entro");

            WebRequest req = WebRequest.Create("http://localhost:8888/api/MegaPrint");
            WebResponse res =  req.GetResponse() ;
            string reader = new StreamReader(res.GetResponseStream()).ReadToEnd();
            Console.WriteLine(res);
            Console.WriteLine("salgo");
           // tmrDelay = new System.Timers.Timer(5000);
           // tmrDelay.Elapsed += new System.Timers.ElapsedEventHandler(tmrDelay_Elapsed);
        }

        private async static void tmrDelay_Elapsed(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine("entro");

            WebRequest req = WebRequest.Create("http://localhost:8888/api/MegaPrint");
            WebResponse res = await req.GetResponseAsync().ConfigureAwait(false);
            string reader = new StreamReader(res.GetResponseStream()).ReadToEnd();
            Console.WriteLine(res);
            Console.WriteLine("salgo");

        }
    }
}
