using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FEL_AUTO.Models
{
    public class FacturaEspecial:Documento
    {
        Documents FESP;
        bool Lleno;
        public FacturaEspecial(string DocEntry, string Tipo, string XML, string[] Parametros, string CodigoSerie)
        {
            this.DocEntry = DocEntry;
            this.Tipo = Tipo;
            this.XML = XML;
            this.Parametros = Parametros;
            this.CodigoSerie = CodigoSerie;
            Recordset oRecordset = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select \"SeriesName\" from NNM1 where \"Series\" = " + CodigoSerie);
            this.NombreSerie = oRecordset.Fields.Item(0).Value;
            FESP = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            this.Lleno = FESP.GetByKey(Int32.Parse(this.DocEntry));
            this.DocNum = this.FESP.DocNum.ToString();
            oRecordset = null;
            GC.Collect();
        }

        public void SalvarInfoCert()
        {
            if (this.estado_certificacion == 1 && this.estado_token == 1)
            {
                FESP.UserFields.Fields.Item("U_ESTADO_FACE").Value = "A";
                FESP.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = this.FirmaElectronica;
                FESP.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = this.NoDocSAT;
                FESP.UserFields.Fields.Item("U_SERIE_FACE").Value = this.SerieSAT;
                FESP.UserFields.Fields.Item("U_FACE_PDFFILE").Value = this.PDF;
                FESP.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = this.FechaCert;
                FESP.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = this.FechaEnvio;
                FESP.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = "";
                FESP.UserFields.Fields.Item("U_FECHA_NC").Value = this.FechaCert.Split('T')[0];

            }
            else
            {
                FESP.UserFields.Fields.Item("U_ESTADO_FACE").Value = "R";
                FESP.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = this.RetError;
                FESP.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = "";
                FESP.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = "";
                FESP.UserFields.Fields.Item("U_SERIE_FACE").Value = "";
                FESP.UserFields.Fields.Item("U_FACE_PDFFILE").Value = "";
                FESP.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = "";
                FESP.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = "";
                FESP.UserFields.Fields.Item("U_FELMensaje").Value = "";
            }
            FESP.UserFields.Fields.Item("U_FELMensaje").Value = this.XML;

            FESP.Update();
        }
    }
}