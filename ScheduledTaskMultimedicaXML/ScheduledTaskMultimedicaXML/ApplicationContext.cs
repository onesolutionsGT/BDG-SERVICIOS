using SAPbobsCOM;
using ScheduledTaskMultimedicaXML.Modelos;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduledTaskMultimedicaXML
{
    class ApplicationContext
    {
        public static string ConnectionString { get; set; }
        public static SAPbobsCOM.Company SBOCompany { get; set; }
        public static string SBOError { get { return "Error (" + SBOCompany.GetLastErrorCode() + "): " + SBOCompany.GetLastErrorDescription(); } }
        private static string Server { get; set; }
        private static string CompanyDB { get; set; }
        private static string UserName { get; set; }
        private static string Password { get; set; }
        private static string UseTrusted { get; set; }
        private static string DbUserName { get; set; }
        private static string DbPassword { get; set; }


        public static bool SetApplication()
        {
            Server = System.Configuration.ConfigurationManager.AppSettings["Server"];
            CompanyDB = System.Configuration.ConfigurationManager.AppSettings["Database"];
            UserName = System.Configuration.ConfigurationManager.AppSettings["UserName"];
            Password = System.Configuration.ConfigurationManager.AppSettings["Password"];
            UseTrusted = System.Configuration.ConfigurationManager.AppSettings["useTrusted"];
            DbUserName = System.Configuration.ConfigurationManager.AppSettings["DbUserName"];
            DbPassword = System.Configuration.ConfigurationManager.AppSettings["DbPassword"];

            if (string.IsNullOrEmpty(Server)) throw new Exception("Se necesita el nombre o dirección del servidor SAP");
            if (string.IsNullOrEmpty(CompanyDB)) throw new Exception("Se necesita del nombre de la compania a conectar");
            if (string.IsNullOrEmpty(UserName)) throw new Exception("Se necesita del nombre del usuario para acceder la compania");
            if (string.IsNullOrEmpty(Password)) throw new Exception("Se necesita de la contraseña para acceder la compania");
            if (string.IsNullOrEmpty(DbUserName)) throw new Exception("Se necesita del usuario de DB para acceder la compania");
            if (string.IsNullOrEmpty(DbPassword)) throw new Exception("Se necesita de la clave de DB para acceder la compania");

            bool trusted = false;
            Boolean.TryParse(UseTrusted, out trusted);

            SBOCompany = new SAPbobsCOM.Company();
            Company B1Company = new Company();
            B1Company.Server = Server;
            B1Company.UserName = UserName;
            B1Company.Password = Password;
            B1Company.CompanyDB = CompanyDB;
            B1Company.DbUserName = DbUserName;
            B1Company.DbPassword = DbPassword;
            B1Company.DbServerType = BoDataServerTypes.dst_HANADB;
            B1Company.language = SAPbobsCOM.BoSuppLangs.ln_Spanish;
            B1Company.UseTrusted = trusted;
            int ret = B1Company.Connect();
            if (ret != 0)
            {
                string errMsg = B1Company.GetLastErrorDescription();
                int ErrNo = B1Company.GetLastErrorCode();
                return false;
            }

            SBOCompany = B1Company;
            return true;
        }

        public static Recordset GetDoc(string table, string docNum, string serie, string series_arr)
        {
            string[] docs = System.Configuration.ConfigurationManager.AppSettings[series_arr].Split(';');
            Recordset Arr_RecordSet = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            int contador = 0;
            for (int i = 0; i < docs.Length; i++)
            {
                if (docs[i] != "")
                {
                    string[] serie_info = docs[i].Split(':');
                    string serie_num = serie_info[1];
                    if (serie_num == serie)
                    {
                        Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string query = "Select \"DocEntry\", \"U_Proveedor\",\"DocDate\", \"U_ESTADO_FACE\", \"Series\" from " + table + " WHERE \"U_Proveedor\" = 'AUTO' and(\"U_ESTADO_FACE\" is null OR \"U_ESTADO_FACE\" = 'R' ) and \"Series\" =" + serie_num + " and \"DocNum\" = " + docNum;
                        oRecordset.DoQuery(query);
                        Arr_RecordSet = oRecordset;
                        contador++;
                    }
                }
            }
            return Arr_RecordSet;
        }


        public static void GetDocs()
        {
            string[] docs_fact = System.Configuration.ConfigurationManager.AppSettings["series_fact"].Split(';');
            string[] docs_ncre = System.Configuration.ConfigurationManager.AppSettings["series_ncre"].Split(';');
            string[] docs_ndeb = System.Configuration.ConfigurationManager.AppSettings["series_ndeb"].Split(';');
            string[] docs_fesp = System.Configuration.ConfigurationManager.AppSettings["series_fesp"].Split(';');
            for (int i = 0; i < 4; i++)
            {
                switch (i)
                {
                    case 0:
                        for (int j = 0; j < docs_fact.Length; j++)
                        {
                            if (docs_fact[j] != "")
                            {
                                string[] serie_info = docs_fact[j].Split(':');
                                string serie_num = serie_info[1];
                                Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string query = "Select \"DocEntry\", \"U_Proveedor\",\"DocDate\", \"U_ESTADO_FACE\", \"Series\" from OINV WHERE \"U_Proveedor\" = 'AUTO' and(\"U_ESTADO_FACE\" is null OR \"U_ESTADO_FACE\" = 'R' ) and \"Series\" =" + serie_num;
                                oRecordset.DoQuery(query);
                                RecorrerDocs(oRecordset, "series_fact", 0, serie_num);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);

                            }
                        }
                        break;
                    case 1:
                        for (int j = 0; j < docs_ncre.Length; j++)
                        {
                            if (docs_ncre[j] != "")
                            {
                                string[] serie_info = docs_ncre[j].Split(':');
                                string serie_num = serie_info[1];
                                Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string query = "Select \"DocEntry\", \"U_Proveedor\",\"DocDate\", \"U_ESTADO_FACE\", \"Series\" from ORIN WHERE \"U_Proveedor\" = 'AUTO' and(\"U_ESTADO_FACE\" is null OR \"U_ESTADO_FACE\" = 'R' ) and \"Series\" =" + serie_num;
                                oRecordset.DoQuery(query);
                                RecorrerDocs(oRecordset, "series_ncre", 1, serie_num);

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);

                            }
                        }
                        break;
                    case 2:
                        for (int j = 0; j < docs_ndeb.Length; j++)
                        {
                            if (docs_ndeb[j] != "")
                            {
                                string[] serie_info = docs_ndeb[j].Split(':');
                                string serie_num = serie_info[1];
                                Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string query = "Select \"DocEntry\", \"U_Proveedor\",\"DocDate\", \"U_ESTADO_FACE\", \"Series\" from OINV WHERE \"U_Proveedor\" = 'AUTO' and(\"U_ESTADO_FACE\" is null OR \"U_ESTADO_FACE\" = 'R' ) and \"Series\" =" + serie_num;
                                oRecordset.DoQuery(query);
                                RecorrerDocs(oRecordset, "series_ndeb", 2, serie_num);

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);

                            }
                        }
                        break;
                    case 3:
                        for (int j = 0; j < docs_fesp.Length; j++)
                        {
                            if (docs_fesp[j] != "")
                            {
                                string[] serie_info = docs_fesp[j].Split(':');
                                string serie_num = serie_info[1];
                                Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string query = "Select \"DocEntry\", \"U_Proveedor\",\"DocDate\", \"U_ESTADO_FACE\", \"Series\" from OPCH WHERE \"U_Proveedor\" = 'AUTO' and(\"U_ESTADO_FACE\" is null OR \"U_ESTADO_FACE\" = 'R' ) and \"Series\" =" + serie_num;
                                oRecordset.DoQuery(query);
                                RecorrerDocs(oRecordset, "series_fesp", 3, serie_num);

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);

                            }
                        }
                        break;
                }
            }
            GC.Collect();
        }

        public static void ActualizarCorrecto(int DocEntry)
        {
            Documents Factura = SBOCompany.GetBusinessObject(BoObjectTypes.oInvoices);
            if (Factura.GetByKey(DocEntry) == true)
            {
                Factura.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = "RECORRIDO";
                Factura.Update();
            }
            Factura = null;
            GC.Collect();
        }

        public static void RecorrerDocs(Recordset Arr, string docs_arr, int indicador, string serie)
        {
            string[] parametros = GetParams();
            string[] docs = System.Configuration.ConfigurationManager.AppSettings[docs_arr].Split(';');
            Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            if (Arr != null)
            {
                for (int i = 0; i < docs.Length; i++)
                {
                    if (docs[i] != "")
                    {
                        string[] serie_info = docs[i].Split(':');
                        string serie_num = serie_info[1];
                        if (serie_num == serie)
                        {
                            Arr.MoveFirst();
                            while (!Arr.EoF)
                            {
                                string call = "Call " + serie_info[0] + "(" + Arr.Fields.Item("DocEntry").Value + ", '" + serie_info[2] + "')";
                                oRecordset.DoQuery(call);
                                string xml = oRecordset.Fields.Item(0).Value;
                                //ActualizarCorrecto(Arr.Fields.Item("DocEntry").Value);
                                if (indicador == 0)
                                {
                                    Factura nueva = new Factura(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                    if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                    {
                                        nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                        nueva.CertificarMegaPrint();
                                    }
                                    nueva.SalvarInfoCert();
                                }
                                else if (indicador == 1)
                                {
                                    NotaCredito nueva = new NotaCredito(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                    if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                    {
                                        nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                        nueva.CertificarMegaPrint();
                                    }
                                    nueva.SalvarInfoCert();
                                }
                                else if (indicador == 2)
                                {
                                    NotaDebito nueva = new NotaDebito(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                    if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                    {
                                        nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                        nueva.CertificarMegaPrint();
                                    }
                                    nueva.SalvarInfoCert();
                                }
                                else if (indicador == 3)
                                {
                                    FacturaEspecial nueva = new FacturaEspecial(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                    if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                    {
                                        nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                        nueva.CertificarMegaPrint();
                                    }
                                    nueva.SalvarInfoCert();
                                }
                                Arr.MoveNext();
                            }
                        }
                    }
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            GC.Collect();
        }
        public static void RecorrerDoc(Recordset Arr, string docs_arr, int indicador, string serie)
        {
            string[] parametros = GetParams();
            string[] docs = System.Configuration.ConfigurationManager.AppSettings[docs_arr].Split(';');
            Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            if (Arr != null)
            {
                for (int i = 0; i < docs.Length; i++)
                {
                    if (docs[i] != "")
                    {
                        string[] serie_info = docs[i].Split(':');
                        string serie_num = serie_info[1];
                        if (serie_num == serie)
                        {
                            string call = "Call " + serie_info[0] + "(" + Arr.Fields.Item("DocEntry").Value + ", '" + serie_info[2] + "')";
                            oRecordset.DoQuery(call);
                            string xml = oRecordset.Fields.Item(0).Value;
                            //ActualizarCorrecto(Arr.Fields.Item("DocEntry").Value);
                            if (indicador == 0)
                            {
                                Factura nueva = new Factura(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                {
                                    nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                    nueva.CertificarMegaPrint();
                                }
                                nueva.SalvarInfoCert();
                            }
                            else if (indicador == 1)
                            {
                                NotaCredito nueva = new NotaCredito(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                {
                                    nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                    nueva.CertificarMegaPrint();
                                }
                                nueva.SalvarInfoCert();
                            }
                            else if (indicador == 2)
                            {
                                NotaDebito nueva = new NotaDebito(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                {
                                    nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                    nueva.CertificarMegaPrint();
                                }
                                nueva.SalvarInfoCert();
                            }
                            else if (indicador == 3)
                            {
                                FacturaEspecial nueva = new FacturaEspecial(Arr.Fields.Item("DocEntry").Value.ToString(), serie_info[2], xml, parametros, serie_info[1]);
                                if (System.Configuration.ConfigurationManager.AppSettings["certificador"].ToLower() == "megaprint")
                                {
                                    nueva.FirmarMegaPrint(nueva.AlmacenarXML());
                                    nueva.CertificarMegaPrint();
                                }
                                nueva.SalvarInfoCert();
                            }
                        }
                    }
                }
            }
            oRecordset = null;
            GC.Collect();
        }

        public static string GetParametro(string parametro)
        {
            Recordset oRecordset = SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select * from \"@FEL_PARAMETROS\" where \"U_PARAMETRO\"='" + parametro + "'");
            return oRecordset.Fields.Item("U_VALOR").Value;
        }

        public static string[] GetParams()
        {
            string[] parametros = new string[12];
            parametros[0] = GetParametro("PATHXML");
            parametros[1] = GetParametro("PATHPDF");
            parametros[2] = GetParametro("PATHXMLc");
            parametros[3] = GetParametro("PATHXMLcp");
            parametros[4] = GetParametro("PATHXMLaut");
            parametros[5] = GetParametro("PATHXMLres");
            parametros[6] = GetParametro("ApiKey");
            parametros[7] = GetParametro("UR_r");
            parametros[8] = GetParametro("UR_t");
            parametros[9] = GetParametro("UR_p");
            parametros[10] = GetParametro("PATHXMLerr");
            parametros[11] = GetParametro("NitEmi");
            return parametros;
        }
    }
}
