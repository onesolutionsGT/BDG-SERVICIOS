using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduledTaskMultimedicaXML.Modelos
{
    class NotaDebito:Documento
    {
        Documents NDEB;
        bool Lleno;
        public NotaDebito(string DocEntry, string Tipo, string XML, string[] Parametros, string CodigoSerie)
        {
            this.DocEntry = DocEntry;
            this.Tipo = Tipo;
            this.XML = XML;
            this.Parametros = Parametros;
            this.CodigoSerie = CodigoSerie;
            Recordset oRecordset = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select \"SeriesName\" from NNM1 where \"Series\" = " + CodigoSerie);
            this.NombreSerie = oRecordset.Fields.Item(0).Value;
            NDEB = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.oInvoices);
            this.Lleno = NDEB.GetByKey(Int32.Parse(this.DocEntry));
            this.DocNum = this.NDEB.DocNum.ToString();
            oRecordset = null;
            GC.Collect();
        }

        public void SalvarInfoCert()
        {
            if (this.estado_certificacion == 1 && this.estado_token == 1)
            {
                NDEB.UserFields.Fields.Item("U_ESTADO_FACE").Value = "A";
                NDEB.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = this.FirmaElectronica;
                NDEB.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = this.NoDocSAT;
                NDEB.UserFields.Fields.Item("U_SERIE_FACE").Value = this.SerieSAT;
                NDEB.UserFields.Fields.Item("U_FACE_PDFFILE").Value = this.PDF;
                NDEB.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = this.FechaCert;
                NDEB.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = this.FechaEnvio;
                NDEB.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = "";
                NDEB.UserFields.Fields.Item("U_FECHA_NC").Value = this.FechaCert.Split('T')[0];

            }
            else
            {
                NDEB.UserFields.Fields.Item("U_ESTADO_FACE").Value = "R";
                NDEB.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = this.RetError;
                NDEB.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = "";
                NDEB.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = "";
                NDEB.UserFields.Fields.Item("U_SERIE_FACE").Value = "";
                NDEB.UserFields.Fields.Item("U_FACE_PDFFILE").Value = "";
                NDEB.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = "";
                NDEB.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = "";
                NDEB.UserFields.Fields.Item("U_FELMensaje").Value = "";
            }
            NDEB.UserFields.Fields.Item("U_FELMensaje").Value = this.XML;

            NDEB.Update();
        }
    }
}
