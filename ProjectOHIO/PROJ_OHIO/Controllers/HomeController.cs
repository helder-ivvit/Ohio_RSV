using System;
using System.Web.Mvc;
using System.Data;
using System.Web.Script.Serialization;

namespace ProjectOHIO.Controllers
{
    public class HomeController : Controller
    {
        private string GetUser()
        {
            var sesion = Convert.ToString(Session["User"]);
            return sesion;
        }
        private string GetOrigen()
        {
            var sesion = Convert.ToString(Session["Origen"]);
            return sesion;
        }
        public ActionResult Index()
        {
            if (Session["UserID"] != null)
            {

                //ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();
                //DataTable nDT_Obj, nDT_Req;

                //string origen = GetOrigen();
                //if (origen == "  " || origen == "")
                //{
                //    origen = "01";
                //}
                //string usuario = GetUser();
                //DateTime fecha = DateTime.Now;
                //nDT_Obj = nObj.Obtener_Listado("sp_web_dashboard_oit", origen, fecha).Tables[0];

                //foreach (DataRow row in nDT_Obj.Rows)
                //{
                //    ViewBag.registradas = Convert.ToInt32(row["registrada"]);
                //    ViewBag.proceso = Convert.ToInt32(row["proceso"]);
                //    ViewBag.cerradas = Convert.ToInt32(row["cerrada"]);                    
                //}

                //nDT_Req = nObj.Obtener_Listado("sp_web_dashboard_listar_requerimientos", origen, usuario).Tables[0];

                //return View(nDT_Req);
                return View();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}