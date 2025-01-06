using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using PROJ_OHIO.Clases;

namespace PROJ_OHIO.Controllers
{
    public class UniversalController : Controller
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
        // GET: Universal
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

        [HttpGet]
        public ActionResult List(string busqueda, int estado, int page, int filas)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("sp_vida_universal_listar_pagina", busqueda, estado, page, filas).Tables[0];
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

        [HttpGet]
        public ActionResult ListPeriodos()
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            DataTable nDT_Parametros;
            nDT_Parametros = nObj.Obtener_Listado("sp_vida_universal_listar_periodos").Tables[0];
            
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
        public ActionResult Cobranza(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                //string oper = "INS";
                //string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                //string[] words = periodo.Split('-');

                string moneda = Convert.ToString(formCollection["v_moneda_cobranza"]);
                int anio = Convert.ToInt32(formCollection["v_anio_cobranza"]);
                int mes = Convert.ToInt32(formCollection["v_mes_cobranza"]);

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("SP_PAGO_VIDA_OPERACION", anio, mes, moneda, GetUser(), GetHost()).Tables[0];

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
        public ActionResult Endoso(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                //string oper = "INS";
                //string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                //string[] words = periodo.Split('-');
                int anio = Convert.ToInt32(formCollection["v_anio_endoso"]);
                int mes = Convert.ToInt32(formCollection["v_mes_endoso"]);

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("SP_POLIZA_MOVIMIENTO_OPERACION", anio, mes, GetUser(), GetHost()).Tables[0];

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

                string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                string[] words = periodo.Split('-');
                int anio = Convert.ToInt32(words[0]);
                int mes = Convert.ToInt32(words[1]);

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("sp_vida_universal_operacion", oper, anio, mes, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, GetUser(), GetHost()).Tables[0];

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
        public ActionResult ExcelToDatabase()
        {
            bool result = false;
            string mensaje = "";
            try
            {
                if (Request.Files[0].ContentLength > 0)
                {
                    string extension = System.IO.Path.GetExtension(Request.Files[0].FileName).ToLower();
                    string connString = "";
                    
                    string[] validFileTypes = { ".xls", ".xlsx" };
                    DataTable dt = new DataTable();
                    
                    string _moneda = Request.Form["moneda"];
                    int _anio = Convert.ToInt32(Request.Form["anio"]);
                    int _mes = Convert.ToInt32(Request.Form["mes"]);

                    //string PathRaiz = ConfigurationManager.AppSettings["PathRuta"];
                    //var _PathAplicacion = HttpContext.Current.Request.PhysicalApplicationPath;
                    //var _PathAplicacion1 = HttpContext.Server.MapPath(PathRaiz);

                    string Directorio = Path.GetTempPath();  //Server.MapPath(PathRaiz); // Path.Combine(PathRaiz); // Path.Combine(Server.MapPath(PathRaiz));
                    string path1 = Path.Combine(Directorio, Request.Files[0].FileName);

                    //using (FileStream fs = new FileStream(path1, FileMode.Open))
                    //{
                    //    // Escritura de datos de usuario o de aplicación...
                    //}

                    if (!Directory.Exists(Directorio))
                        Directory.CreateDirectory(Directorio);

                    if (validFileTypes.Contains(extension))
                    {
                        //// Verifica si existe el archivo en la ruta especificada para que lo elimine
                        if (System.IO.File.Exists(path1))
                        {
                            //System.IO.FileStream fs = new System.IO.FileStream(path1, System.IO.FileMode.Open);
                            //fs.Close();
                            //var proceso = System.Diagnostics.Process.Start(path1);
                            //proceso.WaitForExit();
                            System.IO.File.Delete(path1);
                        }
                            
                        
                        //// Guarda el archivo en la ruta especificada
                        Request.Files[0].SaveAs(path1);

                        // HDR=No   ==> para que no se omita los encabezados o cabeceras
                        // IMEX=1   ==> para que no haya conflito en los datos que se leen

                        //if (extension.Trim() == ".xls")
                        //    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";

                        //if (extension.Trim() == ".xlsx")
                        //    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";

                        //Service _service = new Service();
                        //dt = _service.ImportExceltoDatable(connString);


                        //this.Response.Write("Welcome!  ::  " + dt.Rows.Count);

                        ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();
                        string msg = "";
                        DataTable nDT_Parametros;
                        nDT_Parametros = nObj.Obtener_Listado("SP_VIDA_UNIVERSAL_DATA_RESERVA_ACTUARIAL_ELIMINAR", _anio, _mes, _moneda).Tables[0];

                        foreach (DataRow row in nDT_Parametros.Rows)
                        {
                            result = Convert.ToBoolean(row["result"]);
                            msg = Convert.ToString(row["mensaje"]);
                        }

                        //if (dt.Rows.Count > 0)
                        //{
                        ExcelToDatabase _exceltodatatbase = new ExcelToDatabase();
                            //result = _exceltodatatbase.PolizasDatabase(dt, _moneda, _anio, _mes);
                            result = _exceltodatatbase.PolizasDatabaseExcel(dt, _moneda, _anio, _mes, path1);

                            //result = PolizasDatabase(dt, _moneda, _anio, _mes);

                            //this.Response.Write("result!  ::  " + result);
                            if (result)
                            {
                                mensaje = "Data importada satisfactoriamente";
                                //// al finalizar debe eliminar el archivo
                                if (System.IO.File.Exists(path1))
                                    System.IO.File.Delete(path1);
                            }                                
                            else
                                mensaje = "Hay algún problema al importar los datos.";
                        //}
                        //else
                        //{
                        //    result = false;
                        //    mensaje = "No hay registros para importar";
                        //}
                    }
                    else
                    {
                        result = false;
                        mensaje = "Cargue archivos en formato .xls o .xlsx";
                    }
                }
            }
            catch (Exception ex)
            {
                result = Convert.ToBoolean(0);
                mensaje = ex.ToString();
            }

            return Json(new
            {
                operacion = result,
                mensaje = mensaje
            });

        }

        [HttpPost]
        public ActionResult ProcesaVP(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                //string oper = "INS";
                //string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                //string[] words = periodo.Split('-');

                string moneda = Convert.ToString(formCollection["v_moneda_vp"]);
                int anio = Convert.ToInt32(formCollection["v_anio_vp"]);
                int mes = Convert.ToInt32(formCollection["v_mes_vp"]);

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("SP_VALOR_POLIZA", anio, mes, moneda, GetUser()).Tables[0];

                foreach (DataRow row in nDT_Parametros.Rows)
                {
                    result = Convert.ToBoolean(row["result"]);
                    msg = Convert.ToString(row["mensaje"]); // + Convert.ToString(row["procedimiento"]);
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
        public ActionResult ProcesaVU(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                //string oper = "INS";
                //string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                //string[] words = periodo.Split('-');

                string moneda = Convert.ToString(formCollection["v_moneda_reserva_vu"]);
                int anio = Convert.ToInt32(formCollection["v_anio_reserva_vu"]);
                int mes = Convert.ToInt32(formCollection["v_mes_reserva_vu"]);

                int flag_gastos = 0;
                int flag_mortalidad = 0;
                int flag_catastrofe = 0;
                int flag_itp = 0;
                int flag_caida_alza = 0;
                int flag_caida_baja = 0;
                int flag_caida_mas = 0;

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("SP_VIDA_UNIVERSAL", anio, mes, moneda, GetUser(), flag_gastos, flag_mortalidad, flag_catastrofe, flag_itp, flag_caida_alza, flag_caida_baja, flag_caida_mas).Tables[0];

                foreach (DataRow row in nDT_Parametros.Rows)
                {
                    result = Convert.ToBoolean(row["result"]);
                    msg = Convert.ToString(row["mensaje"]); // + Convert.ToString(row["procedimiento"]);
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
        public ActionResult ProcesaMoceVU(FormCollection formCollection)
        {
            ModeloNegocio.MyNegocio nObj = new ModeloNegocio.MyNegocio();

            bool result = false;
            string msg = "";
            bool result2 = true;
            string msg2 = "";

            try
            {
                //string oper = "INS";
                //string periodo = Convert.ToString(formCollection["v_periodo_universal"]);
                //string[] words = periodo.Split('-');

                string moneda = Convert.ToString(formCollection["v_moneda_moce_vu"]);
                int anio = Convert.ToInt32(formCollection["v_anio_moce_vu"]);
                int mes = Convert.ToInt32(formCollection["v_mes_moce_vu"]);

                //int flag_gastos = 0;
                //int flag_mortalidad = 0;
                //int flag_catastrofe = 0;
                //int flag_itp = 0;
                //int flag_caida_alza = 0;
                //int flag_caida_baja = 0;
                //int flag_caida_mas = 0;

                DataTable nDT_Parametros;
                nDT_Parametros = nObj.Obtener_Listado("SP_MOCE_VIDA_UNIVERSAL", anio, mes, moneda, GetUser()).Tables[0];

                foreach (DataRow row in nDT_Parametros.Rows)
                {
                    result = Convert.ToBoolean(row["result"]);
                    msg = Convert.ToString(row["mensaje"]); // + Convert.ToString(row["procedimiento"]);
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