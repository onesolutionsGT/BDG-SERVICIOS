using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FirmaXadesNet;
using System.Xml;
using FirmaXadesNet.Signature.Parameters;

namespace FEL_AUTO.Models
{
    public class Factura : Documento
    {
        Documents OINV;
        bool Lleno;
        public Factura(string DocEntry, string Tipo, string XML, string[] Parametros, string CodigoSerie)
        {
            this.DocEntry = DocEntry;
            this.Tipo = Tipo;
            this.XML = XML;
            this.Parametros = Parametros;
            this.CodigoSerie = CodigoSerie;
            Recordset oRecordset = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select \"SeriesName\" from NNM1 where \"Series\" = " + CodigoSerie);
            this.NombreSerie = oRecordset.Fields.Item(0).Value;
            OINV = ApplicationContext.SBOCompany.GetBusinessObject(BoObjectTypes.oInvoices);
            this.Lleno = OINV.GetByKey(Int32.Parse(this.DocEntry));
            this.DocNum = this.OINV.DocNum.ToString();
            oRecordset = null;
            GC.Collect();
        }

        public void SalvarInfoCert()
        {
            if (this.estado_certificacion == 1 && this.estado_token == 1)
            {
                OINV.UserFields.Fields.Item("U_ESTADO_FACE").Value = "A";
                OINV.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = this.FirmaElectronica;
                OINV.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = this.NoDocSAT;
                OINV.UserFields.Fields.Item("U_SERIE_FACE").Value = this.SerieSAT;
                OINV.UserFields.Fields.Item("U_FACE_PDFFILE").Value = this.PDF;
                OINV.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = this.FechaCert;
                OINV.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = this.FechaEnvio;
                OINV.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = "";
                OINV.UserFields.Fields.Item("U_NUMERO_DOCUMENTO_NC").Value = this.FirmaElectronica;
                OINV.UserFields.Fields.Item("U_FECHA_NC").Value = this.FechaCert.Split('T')[0];
            }
            else
            {
                OINV.UserFields.Fields.Item("U_ESTADO_FACE").Value = "R";
                OINV.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = this.RetError;
                OINV.UserFields.Fields.Item("U_FIRMA_ELETRONICA").Value = "";
                OINV.UserFields.Fields.Item("U_NUMERO_DOCUMENTO").Value = "";
                OINV.UserFields.Fields.Item("U_SERIE_FACE").Value = "";
                OINV.UserFields.Fields.Item("U_FACE_PDFFILE").Value = "";
                OINV.UserFields.Fields.Item("U_FECHA_CERT_FACE").Value = "";
                OINV.UserFields.Fields.Item("U_FECHA_ENVIO_FACE").Value = "";
                OINV.UserFields.Fields.Item("U_FELMensaje").Value = "";
            }
            OINV.UserFields.Fields.Item("U_FELMensaje").Value = this.XML;
            OINV.Update();
        }
    }
}