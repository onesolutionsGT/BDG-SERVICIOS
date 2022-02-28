using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace FEL_BACKEND
{
    class ApplicationContext
    {
        public static string path { get; set; }
        public static string ConnectionString { get; set; }
        public static SAPbouiCOM.Application SBOApplication { get; set; }
        public static SAPbobsCOM.Company SBOCompany { get; set; }
        public static string SBOError { get { return "Error (" + SBOCompany.GetLastErrorCode() + "): " + SBOCompany.GetLastErrorDescription(); } }
        public Form MainForm { get; set; }

        public FormCreationParams creationPackage;
        public static void SetApplication()
        {
            path = System.Windows.Forms.Application.StartupPath;
            //Se obtiene string de conexion de Cliente SAP B1
            if (Environment.GetCommandLineArgs().Count() == 1) { throw new Exception("No se agregaron los parametros de conexión...", new Exception("No se encontro string de conexión SAP B1")); }
            ConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();

            //Se realiza conexion 
            SAPbouiCOM.SboGuiApi client = new SAPbouiCOM.SboGuiApi();
            client.Connect(ConnectionString);
            SBOApplication = client.GetApplication(-1);

            //Se carga <<Company>> de aplicacion   
            SBOCompany = new SAPbobsCOM.Company();
            string cookies = SBOCompany.GetContextCookie();
            string connectionContext = SBOApplication.Company.GetConnectionContext(SBOCompany.GetContextCookie());
            SBOCompany.SetSboLoginContext(connectionContext);

            //Conexion con sociedad
            if (SBOCompany.Connect() != 0) { throw new Exception(SBOError); }


        }

        public static void LoadFromXml(string fileName, string formName, FormCreationParams creationPackage)
        {
            if (ActivateFormIsOpen(formName)) { return; }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(fileName);
            creationPackage.XmlData = xmlDocument.InnerXml;
            SBOApplication.LoadBatchActions(xmlDocument.InnerXml);
        }


        //Activa Form de SAP BO Client con atributo <<formId>>

        public static bool ActivateFormIsOpenWithTypeEx(string TypeEx)
        {

            if (SBOApplication == null) { throw new Exception("Interfaz de la aplicación nula"); }
            if (SBOApplication.Forms.Count == 0) { throw new Exception("No se encontraron formularios"); }

            for (int x = 0; x < SBOApplication.Forms.Count; x++)
            {

                string typeEx = SBOApplication.Forms.Item(x).TypeEx;
                string type = SBOApplication.Forms.Item(x).Type.ToString();

                if (typeEx == TypeEx)
                {
                    SBOApplication.Forms.Item(x).Select();
                    return true;
                }
            }

            return false;
        }
        public static bool ActivateFormIsOpen(string formID)
        {

            if (SBOApplication == null) { throw new Exception("Interfaz de la aplicación nula"); }
            if (SBOApplication.Forms.Count == 0) { throw new Exception("No se encontraron formularios"); }

            for (int x = 0; x < SBOApplication.Forms.Count; x++)
            {
                string uniqueId = SBOApplication.Forms.Item(x).UniqueID;

                if (uniqueId == formID)
                {
                    SBOApplication.Forms.Item(x).Select();
                    return true;
                }
            }

            return false;
        }
        public static void PrintGreen(string mensaje)
        {
            SBOApplication
               .StatusBar
               .SetText(mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }
        public static void PrintRed(string mensaje)
        {
            SBOApplication
               .StatusBar
               .SetText(mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
    }
}
