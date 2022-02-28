using FirmaXadesNet;
using FirmaXadesNet.Crypto;
using FirmaXadesNet.Signature;
using FirmaXadesNet.Signature.Parameters;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ScheduledTaskMultimedicaXML.Modelos
{
    class Documento
    {
        public string DocEntry;
        public string DocNum;
        public string Tipo;
        public string XML;
        public string[] Parametros;
        public string NombreArchivo;
        public string RetError;
        public string FirmaElectronica;
        public string Autorizacion;
        public string SerieSAT;
        public string NoDocSAT;
        public string FechaEnvio;
        public string FechaCert;
        public string Motivo;
        public string CodigoSerie;
        public string NombreSerie;
        public string XMLCert;
        public string token;
        public string PDF;
        public int estado_token;
        public int estado_certificacion;

        public Documento()
        {
            this.RetError = "";
            this.estado_certificacion = 0;
            this.estado_token = 0;
        }

        public void CertificarMegaPrint()
        {
            this.token = SolicitarTokenMegaPrint();
            string respuesta_certificador = "";
            if (this.estado_token == 1)
            {
                respuesta_certificador = EnviarXMLMegaPrint(this.token);
                AlmacenarRespuestaMegaPrint(respuesta_certificador);
            }
            else
            {

            }
        }

        private void AlmacenarRespuestaMegaPrint(string respuesta_certificador)
        {
            XmlDocument respuesta_xml = new XmlDocument();
            respuesta_xml.LoadXml(respuesta_certificador);
            XmlNodeList lista_respuesta = respuesta_xml.SelectNodes("RegistraDocumentoXMLResponse");
            for (int i = 0; i < lista_respuesta.Count; i++)
            {
                if (lista_respuesta.Item(i).SelectSingleNode("tipo_respuesta").InnerText == "0")
                {
                    string datos_certificacion = lista_respuesta.Item(i).SelectSingleNode("xml_dte").InnerText.Replace("dte:", "");
                    XmlDocument respuesta = new XmlDocument();
                    respuesta.LoadXml(datos_certificacion);
                    XmlNodeList nodos_res = respuesta.SelectNodes("GTDocumento/SAT/DTE/Certificacion");
                    for (int j = 0; j < nodos_res.Count; j++)
                    {
                        this.Autorizacion = nodos_res.Item(j).SelectSingleNode("NumeroAutorizacion").InnerText;
                        this.SerieSAT = nodos_res.Item(j).SelectSingleNode("NumeroAutorizacion").Attributes.Item(1).Value.ToString();
                        this.NoDocSAT = nodos_res.Item(j).SelectSingleNode("NumeroAutorizacion").Attributes.Item(0).Value.ToString();
                        this.FechaCert = nodos_res.Item(j).SelectSingleNode("FechaHoraCertificacion").InnerText;
                    }
                    XmlNodeList nodos_respuesta_2 = respuesta.SelectNodes("GTDocumento/SAT/DTE/DatosEmision");
                    for (int k = 0; k < nodos_respuesta_2.Count; k++)
                    {
                        this.FechaEnvio = nodos_respuesta_2.Item(k).SelectSingleNode("DatosGenerales").Attributes.Item(1).Value.ToString();
                    }
                    this.FirmaElectronica = lista_respuesta.Item(i).SelectSingleNode("uuid").InnerText;
                    SavePDF(GetPDF());
                    this.estado_certificacion = 1;
                }
                else
                {
                    XmlNodeList nodos = lista_respuesta.Item(i).SelectNodes("listado_errores/error");
                    for (int j = 0; j < nodos.Count; j++)
                    {
                        string cod_err = nodos.Item(j).SelectSingleNode("cod_error").InnerText;
                        string dec_error = nodos.Item(j).SelectSingleNode("desc_error").InnerText;
                        this.RetError += "(" + cod_err + ") " + dec_error + "\n";
                    }
                    this.estado_certificacion = 2;
                }
            }
        }
        public void SaveCertData()
        {
            Recordset oRecordset = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string call = "CALL FELONE_UTILS ('True'" +
                                               ", '" + this.DocEntry + "'" +
                                               ", '" + this.Tipo + "'" +
                                               ", '" + this.FirmaElectronica + "'" +
                                               ", '" + this.NoDocSAT + "'" +
                                               ", '" + this.SerieSAT + "'" +
                                               ", '" + this.PDF + "'" +
                                               ", '" + this.FechaCert + "'" +
                                               ", '" + this.FechaEnvio + "','','')";
            oRecordset.DoQuery(call);
        }
        private string SavePDF(XmlDocument PDFRes)
        {
            XmlNodeList Lista_res_PDF = PDFRes.SelectNodes("RetornaPDFResponse");
            for (int i = 0; i < Lista_res_PDF.Count; i++)
            {
                if (Lista_res_PDF.Item(i).SelectSingleNode("tipo_respuesta").InnerText == "0")
                {
                    string PDF_Base_64 = Lista_res_PDF.Item(i).SelectSingleNode("pdf").InnerText;
                    byte[] bytes_PDF = System.Convert.FromBase64String(PDF_Base_64);
                    System.IO.BinaryWriter writer = new System.IO.BinaryWriter(System.IO.File.Open(this.Parametros[1] + "\\\\" + this.Tipo + "_" + this.NombreSerie + "_" + this.DocNum + ".pdf", System.IO.FileMode.Create));
                    writer.Write(bytes_PDF);
                    writer.Close();
                    this.PDF = System.Configuration.ConfigurationManager.AppSettings["virtual_dir"] + this.Tipo + "_" + this.NombreSerie + "_" + this.DocNum + ".pdf";
                }
            }
            return "";
        }
        private XmlDocument GetPDF()
        {
            WebRequest req = WebRequest.Create(this.Parametros[9]);
            string body = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><RetornaPDFRequest><uuid>" + this.FirmaElectronica + "</uuid></RetornaPDFRequest>";
            /*XmlDocument doc = new XmlDocument();
            doc.LoadXml(body);*/
            byte[] data = Encoding.ASCII.GetBytes(body);
            req.Headers.Add("authorization", "Bearer " + this.token);
            req.Method = "POST";
            req.ContentType = "application/json";
            req.ContentLength = data.Length;
            Stream stream = req.GetRequestStream();
            stream.Write(data, 0, data.Length);
            WebResponse res = req.GetResponse();
            string reader = new StreamReader(res.GetResponseStream()).ReadToEnd();
            XmlDocument salida = new XmlDocument();
            salida.LoadXml(reader);
            return salida;
        }

        private string EnviarXMLMegaPrint(string token)
        {
            byte[] data = Encoding.ASCII.GetBytes(this.XMLCert);
            WebRequest req = WebRequest.Create(this.Parametros[7]);
            req.Headers.Add("authorization", "Bearer " + token);
            req.Method = "POST";
            req.ContentType = "application/json";
            req.ContentLength = data.Length;
            Stream stream = req.GetRequestStream();
            stream.Write(data, 0, data.Length);
            WebResponse res = req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            return responseString;
        }

        private string SolicitarTokenMegaPrint()
        {
            string token = "<SolicitaTokenRequest><usuario>" + this.Parametros[11] + "</usuario><apikey>" + this.Parametros[6] + "</apikey></SolicitaTokenRequest>";
            WebRequest req = WebRequest.Create(this.Parametros[8]);
            req.Method = "POST";
            req.ContentType = "application/json";
            req.ContentLength = token.Length;
            byte[] data = Encoding.ASCII.GetBytes(token);
            Stream stream = req.GetRequestStream();
            stream.Write(data, 0, token.Length);
            WebResponse res = req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            XmlDocument doc_token = new XmlDocument();
            doc_token.LoadXml(responseString);
            XmlNodeList nodo = doc_token.SelectNodes("SolicitaTokenResponse");
            string respuesta = "";
            for (int i = 0; i < nodo.Count; i++)
            {
                if (nodo.Item(i).SelectSingleNode("tipo_respuesta").InnerText == "0")
                {
                    respuesta += nodo.Item(i).SelectSingleNode("token").InnerText;
                    estado_token = 1;
                }
                else
                {
                    estado_token = 2;
                    XmlNodeList nodo_err = nodo.Item(i).SelectNodes("listado_errores");
                    for (int j = 0; j < nodo_err.Count; j++)
                    {
                        this.RetError += "ERROR: (" + nodo_err.Item(j).SelectSingleNode("error").SelectSingleNode("cod_error").InnerText + ") " + nodo_err.Item(j).SelectSingleNode("error").SelectSingleNode("desc_error").InnerText + "\n";
                    }
                }
            }
            return respuesta;
        }
        public void Anular()
        {

        }

        public string AlmacenarXML()
        {
            string filename = string.Format("{0}\\{1}_{2}_{3}.xml", Parametros[0], DocNum, Tipo, NombreSerie);
            File.WriteAllText(filename, XML);
            return filename;
        }

        public void FirmarMegaPrint(string filename)
        {
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SignatureParameters parametros = new SignatureParameters();
            parametros.SignaturePackaging = SignaturePackaging.INTERNALLY_DETACHED;
            parametros.InputMimeType = "text/xml";
            parametros.ElementIdToSign = "DatosEmision";
            parametros.SignatureMethod = SignatureMethod.RSAwithSHA256;
            parametros.DigestMethod = DigestMethod.SHA256;
            X509Certificate2 certificado = new X509Certificate2(Parametros[2], Parametros[3], X509KeyStorageFlags.Exportable);
            Signer firma = new Signer(certificado);
            parametros.Signer = firma;
            XadesService xadesService = new XadesService();
            FileStream fs = new FileStream(filename, FileMode.Open);
            SignatureDocument documento = xadesService.Sign(fs, parametros);
            fs.Close();
            XmlDocument doc = documento.Document;
            XmlNode NodoFirma = doc.GetElementsByTagName("ds:Signature").Item(0);
            NodoFirma.ParentNode.RemoveChild(NodoFirma);
            doc.DocumentElement.AppendChild(NodoFirma);
            string ComplementoHexaDecimal = int.Parse(this.DocEntry).ToString("000000000000");
            string XMLReg = "<RegistraDocumentoXMLRequest id =\"A00B00C0-A714-44CE-0000-000000000000\" ><xml_dte><![CDATA[" + doc.InnerXml.ToString() + "]]></xml_dte></RegistraDocumentoXMLRequest>";
            this.XMLCert = XMLReg;
            StreamWriter escritor;
            string path = this.Parametros[4] + "/Auth_" + NombreSerie + "_" + DocNum + ".xml";
            escritor = File.AppendText(path);
            escritor.Write(XMLReg.ToString());
            escritor.Flush();
            escritor.Close();
        }
    }
}
