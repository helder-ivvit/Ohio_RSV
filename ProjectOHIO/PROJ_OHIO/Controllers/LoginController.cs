using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Web.Script.Serialization;

namespace ProjectOHIO.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public JsonResult ValidateUser(string usuario, string password)
        {

            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            //Int32 rpta = nObj.Ejecutar_Procedimiento_result_int("sp_usuario_login", usuario, password);

            //Session["UserID"] = rpta;

            //if (rpta == 1)
            //    return Json(new { Success = true }, JsonRequestBehavior.AllowGet);
            //else
            //    return Json(new { Success = false }, JsonRequestBehavior.AllowGet);


            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                DataTable nDT_Obj;
                nDT_Obj = nObj.Obtener_Listado("sp_web_usuario_login", usuario, password).Tables[0];

                foreach (DataRow row in nDT_Obj.Rows)
                {
                    result = Convert.ToBoolean(row["result"]);
                    msg = Convert.ToString(row["mensaje"]);
                    if (result)
                    {
                        Session["UserID"] = Convert.ToInt32(row["iduser"]);
                        Session["User"] = Convert.ToString(row["vchusuario"]);
                        Session["Perfil"] = Convert.ToInt32(row["perfil"]);
                        Session["Origen"] = Convert.ToString(row["origen"]);
                        //Session["NombreFuncionario"] = Convert.ToString(row["nombrefuncionario"]);
                        //Session["DependenciaID"] = Convert.ToInt32(row["iddependencia"]);
                        //Session["NombreDependencia"] = Convert.ToString(row["nombredependencia"]);
                        //string ip = Request.UserHostAddress;
                        //string hostName = Request.UserHostName;
                        Session["Estacion"] = Request.UserHostName;
                    }
                }

                nDT_Obj = null;
            }
            catch (Exception ex)
            {
                result2 = Convert.ToBoolean(0);
                msg2 = ex.ToString();
            }
            nObj = null;

            return Json(new
            {
                operacion = result,
                mensaje = msg,
                operacion2 = result2,
                mensaje2 = msg2
            });


        }

    }
}