using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;

namespace ProjectOHIO.Controllers
{
    public class MenuController : Controller
    {
        // GET: Menu
        public ActionResult Index()
        {
            if (Session["UserID"] != null)
            {
                return View();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        //public ActionResult _Sidebar()
        //{
        //    ModeloNegocio.MyNegocio nObj1 = new ModeloNegocio.MyNegocio();
        //    DataTable nDT_Parametros1;
        //    nDT_Parametros1 = nObj1.Obtener_Listado("sp_modulo_listar").Tables[0];

        //    return PartialView(nDT_Parametros1);
        //}


        public ActionResult Salir()
        {
            Session.Abandon();
            Response.Cookies.Add(new HttpCookie("ASP.NET_SessionId", ""));

            //return RedirectToAction("Index", "Login");
            return Json(new
            {
                data = 1,
                valor = "Se cerro la sesion"
            });
        }


    }
}