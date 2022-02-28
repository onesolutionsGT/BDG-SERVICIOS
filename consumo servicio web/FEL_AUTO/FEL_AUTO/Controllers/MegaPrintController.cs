using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Mvc;

namespace FEL_AUTO.Controllers
{
    public class MegaPrintController : ApiController
    {
        // GET: api/MegaPrint
        [System.Web.Http.Route("api/MegaPrint/{docNum}/{document}/{indicador:int}/{serie}/{series_arr}")]
        public ActionResult Get(string docNum, string document, int indicador, string serie, string series_arr)
        {
            try
            {
                ApplicationContext.SetApplication();
                ApplicationContext.SBOCompany.StartTransaction();
                ApplicationContext.RecorrerDoc(ApplicationContext.GetDoc(document, docNum, serie, series_arr), series_arr, indicador, serie);
                ApplicationContext.SBOCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                return new HttpStatusCodeResult(201, "Ok");
            }
            catch (Exception ex)
            {
                //string sap = ApplicationContext.SBOError;
                //ApplicationContext.SBOCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                return new HttpStatusCodeResult(400, "ERROR: " +ex.Message);
            }
        }
        public ActionResult Get()
        {
            try
            {
                ApplicationContext.SetApplication();
                ApplicationContext.SBOCompany.StartTransaction();
                ApplicationContext.GetDocs();
                ApplicationContext.SBOCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                return new HttpStatusCodeResult(201, "Ok");
            }
            catch (Exception ex)
            {
                string sap = ApplicationContext.SBOError;
                ApplicationContext.SBOCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                return new HttpStatusCodeResult(400, "ERROR: " + ex.Message);
            }
        }

        // GET: api/MegaPrint/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/MegaPrint
        public void Post([FromBody] string value)
        {
        }

        // PUT: api/MegaPrint/5
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE: api/MegaPrint/5
        public void Delete(int id)
        {
        }
    }
}
