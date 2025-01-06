using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace PROJ_OHIO.Controllers
{
    public class PeriodoController : Controller
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
        public ActionResult getData(string anio)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("sp_periodo_getdataid", anio).Tables[0];

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
            nDT_Parametros = nObj.Obtener_Listado("sp_periodo_listar_pagina", busqueda, estado, page, filas).Tables[0];
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
        public ActionResult Guardar(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                string oper = "INS";
                int anio = Convert.ToInt32(formCollection["v_anio"]);
                int mes = Convert.ToInt32(formCollection["v_mes"]);
                
                //int usuario = GetUser();
                //string estacion = GetHost();
                
                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("sp_periodo_operacion", oper, anio, mes, 1, GetUser(), GetHost()).Tables[0];

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

        [HttpPost]
        public ActionResult Inactivo(int anio, int mes)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                string oper = "ACT";
                int activo = 1;
            
                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("sp_periodo_operacion", oper, anio, mes, activo, GetUser(), GetHost()).Tables[0];

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

        [HttpPost]
        public ActionResult Activo(int anio, int mes)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                string oper = "INA";
                int activo = 0;
                                
                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("sp_periodo_operacion", oper, anio, mes, activo, GetUser(), GetHost()).Tables[0];

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