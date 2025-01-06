using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace PROJ_OHIO.Controllers
{
    public class ConfiguracionController : Controller
    {
        private int GetUser()
        {
            var sesion = Convert.ToInt32(Session["UserID"]);
            return sesion;
        }

        private string GetHost()
        {
            var estacion = Convert.ToString(Session["Estacion"]);
            return estacion;
        }

        // GET: Configuracion
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

        [HttpPost]
        public ActionResult getData(int id)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("sp_parametro_getdataid", id).Tables[0];

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
            return Json(serializer.Serialize(lst), JsonRequestBehavior.AllowGet);

        }

        [HttpGet]
        public ActionResult List(string busqueda, int estado, int page, int filas)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("sp_parametro_listar_pagina", busqueda, estado, page, filas).Tables[0];
            //return View(nDT_Parametros);

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
            return Json(serializer.Serialize(lst), JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult Actualizar(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                string oper = "UPD";
                int id = Convert.ToInt32(formCollection["v_id_parametro"]);
                string _descripcion = Convert.ToString(formCollection["v_descripcion_parametro"]);
                string _valor = Convert.ToString(formCollection["v_valor_parametro"]);

                //int usuario = GetUser();
                //string estacion = GetHost();

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("sp_parametro_operacion", oper, id, _descripcion, _valor, 1, GetUser(), GetHost()).Tables[0];

                foreach (DataRow row in nDT_Parametros.Rows)
                {
                    result = Convert.ToBoolean(row["result"]);
                    msg = Convert.ToString(row["mensaje"]);
                }
                nDT_Parametros = null;
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