using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduledTaskMultimedicaXML.Modelos
{
    class NotaCredito:Documento
    {
        Documents ORIN;
        bool Lleno;
        public NotaCredito(string DocEntry, string Tipo, string XML, string[] Parametros, string CodigoSerie)
        {
            this.DocEntry = DocEntry;
            this.Tipo = Tipo;
            this.XML = XML;
            this.Parametros = Parametros;
            this.CodigoSerie = CodigoSerie;
            Recordset oRecordset = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select \"SeriesName\" from NNM1 where \"Series\" = " + CodigoSerie);
            this.NombreSerie = oRecordset.Fields.Item(0).Value;
            ORIN = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);
            this.Lleno = ORIN.GetByKey(Int32.Parse(this.DocEntry));
            this.DocNum = this.ORIN.DocNum.ToString();
            oRecordset = null;
            GC.Collect();
        }

        public void SalvarInfoCert()
        {
            if (this.estado_certificacion == 1 && this.estado_token == 1)
            {
                ORIN.UserFields.Fields.Item("U_ESTADO_FACE").Value = "A";
                ORIN.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = this.FirmaElectronica;
                ORIN.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = this.NoDocSAT;
                ORIN.UserFields.Fields.Item("U_SERIE_FACE").Value = this.SerieSAT;
                ORIN.UserFields.Fields.Item("U_FACE_PDFFILE").Value = this.PDF;
                ORIN.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = this.FechaCert;
                ORIN.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = this.FechaEnvio;
                ORIN.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = "";
                ORIN.UserFields.Fields.Item("U_FECHA_NC").Value = this.FechaCert.Split('T')[0];

            }
            else
            {
                ORIN.UserFields.Fields.Item("U_ESTADO_FACE").Value = "R";
                ORIN.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = this.RetError;
                ORIN.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = "";
                ORIN.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = "";
                ORIN.UserFields.Fields.Item("U_SERIE_FACE").Value = "";
                ORIN.UserFields.Fields.Item("U_FACE_PDFFILE").Value = "";
                ORIN.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = "";
                ORIN.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = "";
                ORIN.UserFields.Fields.Item("U_FELMensaje").Value = "";
            }
            ORIN.UserFields.Fields.Item("U_FELMensaje").Value = this.XML;

            ORIN.Update();
        }
    }
}
