using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ScheduledTaskMultimedicaXML
{
    public partial class WindowsApiCallService : ServiceBase
    {
        System.Timers.Timer tmrDelay;

        public WindowsApiCallService()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("MyLogSrc"))
            {
                System.Diagnostics.EventLog.CreateEventSource("MyLogSrc", "MyLog");
            }
            tmrDelay = new System.Timers.Timer(60000);
            tmrDelay.Elapsed += new System.Timers.ElapsedEventHandler(tmrDelay_Elapsed);
        }
        async void tmrDelay_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                WriteLogFile(DateTime.Now.ToString() + ": ENTRA A PROCESO");

                if (ApplicationContext.SetApplication())
                {
                    ApplicationContext.SBOCompany.StartTransaction();
                    DateTime inicio = DateTime.Now;
                    WriteLogFile(inicio.ToString() + ": EMPIEZA TRANSACCIÓN");
                    ApplicationContext.GetDocs();
                    ApplicationContext.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    DateTime fin = DateTime.Now;
                    WriteLogFile(fin.ToString() + ": FINALIZÓ TRANSACCIÓN, DURACION DEL PROCESO: "+ (fin-inicio).ToString());
                }
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                string ret = ex.Message +" "+ ex.InnerException+" "+ex.Source +"\n"+st.ToString();
                if(ApplicationContext.SBOCompany != null && ApplicationContext.SBOCompany.InTransaction)
                {
                    ret += ", SAP: " + ApplicationContext.SBOError;
                    ApplicationContext.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                DateTime fin = DateTime.Now;
                WriteLogFile(fin.ToString() + "ERROR: " +ret);
                ApplicationContext.SBOCompany = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationContext.SBOCompany);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationContext.SBOError);
                GC.Collect();
            }
            finally
            {
                WriteLogFile(DateTime.Now.ToString() + ": SALE DE PROCESO");
                ApplicationContext.SBOCompany = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationContext.SBOCompany);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ApplicationContext.SBOError);
                GC.Collect();
            }
        }
        protected override void OnStart(string[] args)
        {
            WriteLogFile("Service is started");
            tmrDelay.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteLogFile("Service is stopped");
        }
        public void WriteLogFile(string message)
        {
            StreamWriter sw = null;
            sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
            sw.WriteLine($"{DateTime.Now.ToString()} : {message}");
            sw.Flush();
            sw.Close();
        }
    }
}
