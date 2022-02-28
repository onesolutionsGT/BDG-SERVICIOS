using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace FEL_BACKEND
{
    class SystemForm
    {
        Form oForm;
        public SystemForm()
        {
            ApplicationContext.SetApplication();
            SetEvents();
            ApplicationContext
               .SBOApplication
               .StatusBar
               .SetText("Inicializando Add-on FACTURACIÓN ELECTRÓNICA ONE SOLUTIONS =)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            ApplicationContext
                .SBOApplication
                .AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApp_AppEvent);

        }

        private void SBOApp_AppEvent(BoAppEventTypes EventType)
        {
            ApplicationContext
                .SBOApplication
                .StatusBar
                .SetText("Finalizando Add-on FACTURACIÓN ELECTRÓNICA ONE SOLUTIONS =)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            System.Environment.Exit(0);
        }

        private void SetEvents()
        {
            ApplicationContext.SBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(FormHandler);
            ApplicationContext
                .SBOApplication
                .ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApp_ItemEvent);
        }

        private void SBOApp_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if ((pVal.FormType == 133 || pVal.FormType == 179 || pVal.FormType == 65303 || pVal.FormType == 141) && ((pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)))
            {
                oForm = ApplicationContext.SBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
            }
        }

        private void FormHandler(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (oForm != null)
            {
                if (BusinessObjectInfo.ActionSuccess == true && (BusinessObjectInfo.FormTypeEx == "133" || BusinessObjectInfo.FormTypeEx == "179" || BusinessObjectInfo.FormTypeEx == "65303" || BusinessObjectInfo.FormTypeEx == "60091" || BusinessObjectInfo.FormTypeEx == "141") && (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE))
                {
                    XmlDocument xml = new XmlDocument();
                    xml.LoadXml(BusinessObjectInfo.ObjectKey);
                    string Serie = oForm.Items.Item("88").Specific.Value.Trim();
                    string DocEntry = xml.GetElementsByTagName("DocEntry").Item(0).InnerText.Trim();
                    string DocNum = oForm.Items.Item("8").Specific.Value;
                    string endponit = "http://10.0.1.10:8888/api/MegaPrint/";
                    switch (BusinessObjectInfo.FormTypeEx)
                    {
                        case "133":
                            endponit += DocNum + "/OINV/0/" + Serie + "/series_fact";//FACTURA
                            break;
                        case "179":
                            endponit += DocNum + "/ORIN/1/" + Serie + "/series_ncre";//NOTA DE CREDITO
                            break;
                        case "65303":
                            endponit += DocNum + "/OINV/2/" + Serie + "/series_ndeb";//NOTA DE DÉBITO
                            break;
                        case "141":
                            endponit += DocNum + "/OPCH/3/" + Serie + "/series_fesp";//FACTURA PROVEEDORES
                            break;
                    }
                    try
                    {
                        WebRequest req = WebRequest.Create(endponit);
                        WebResponse res = req.GetResponse();
                        string reader = new StreamReader(res.GetResponseStream()).ReadToEnd();
                        ApplicationContext.PrintGreen(reader);
                        oForm.Refresh();
                    }
                    catch (Exception ex)
                    {
                        ApplicationContext.PrintRed(ex.Message);
                    }
                }
            }
        }
    }
}
