using PROJ_OHIO.Clases;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace PROJ_OHIO.Controllers
{
    public class ResultadoVPController : Controller
    {
        // GET: ResultadoVP
        public ActionResult Index(int anio, int mes)
        {
            if (Session["UserID"] != null)
            {

                ViewData["busqueda_anio"] = anio;
                ViewData["busqueda_mes"] = mes;

                return View();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        [HttpGet]
        public ActionResult ListarResultado(int anio, int mes, string moneda)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("SP_VALOR_POLIZA_COMPARAR_RESULTADO", anio, mes, moneda).Tables[0];
            
            var lst = new List<Dictionary<string, object>>();
            foreach (DataRow row in nDT_Parametros.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in nDT_Parametros.Columns)
                {
                    dict[col.ColumnName] = (Convert.ToString(row[col]));
                }
                lst.Add(dict);
            }
            nObj = null;
            nDT_Parametros = null;

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 819200000;

            var jsonResult = Json(serializer.Serialize(lst), JsonRequestBehavior.AllowGet);
            jsonResult.MaxJsonLength = int.MaxValue;
            return jsonResult;

            //return Json(serializer.Serialize(lst), JsonRequestBehavior.AllowGet);

        }

        public ActionResult Exportarconsultageneral(int anio, int mes, string moneda)
        {
            ClassReporteExcel _excel = new ClassReporteExcel();
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();
            DataTable nDT_Obj;
            nDT_Obj = nObj.Obtener_Listado("sp_valor_poliza_comparar_resultado", anio, mes, moneda).Tables[0];

            string handle = Guid.NewGuid().ToString();
            byte[] bytesStream = _excel.ExportarBusquedaDocumentosGeneral(nDT_Obj);
            TempData[handle] = bytesStream.ToArray();

            var jsonObject = new JsonResult();
            jsonObject.JsonRequestBehavior = JsonRequestBehavior.AllowGet;
            jsonObject.MaxJsonLength = int.MaxValue;

            string mone = "";
            if (moneda == "S")
                mone = "Soles";
            else
                mone = "Dolares";
            jsonObject.Data = new { FileGuid = handle, FileName = "Resultado_Valor_Poliza  " + anio.ToString() + "-" + mes.ToString() + "  " + mone + "  " + DateTime.Now.ToString() + ".xlsx" };
            return jsonObject;
        }

        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-excel", fileName);
            }
            else
            {
                // Problem - Log the error, generate a blank file,
                //           redirect to another controller action - whatever fits with your application
                //return new EmptyResult();
                return Json(false);
            }
        }
    }
}