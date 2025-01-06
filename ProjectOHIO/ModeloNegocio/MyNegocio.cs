using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ModeloNegocio
{
  public   class MyNegocio
    {

         private ModeloDatos.MyDatos nObj_Datos = new ModeloDatos.MyDatos();
 

        public int Registro_Inserta(string mProcedure, params object[] mArgumentos)
        {
            int Registro_InsertaRet = default(int);
            Registro_InsertaRet = nObj_Datos.Ejecutar_Procedure(mProcedure, mArgumentos);
            return Registro_InsertaRet;
        }

        public int Registro_Inserta_Id(string mProcedure, params object[] mArgumentos)
        {
            int Registro_Inserta_IdRet = default(int);
            Registro_Inserta_IdRet = nObj_Datos.Ejecutar_Procedure_Out(mProcedure, mArgumentos);
            return Registro_Inserta_IdRet;
        }

        public int Registro_Update(string mProcedure, params object[] mArgumentos)
        {
            int Registro_UpdateRet = default(int);
            Registro_UpdateRet = nObj_Datos.Ejecutar_Procedure(mProcedure, mArgumentos);
            return Registro_UpdateRet;
        }

        public int Registro_Delete(string mProcedure, params object[] mArgumentos)
        {
            int Registro_DeleteRet = default(int);
            Registro_DeleteRet = nObj_Datos.Ejecutar_Procedure(mProcedure, mArgumentos);
            return Registro_DeleteRet;
        }

        public int Ejecutar_Procedimiento(string mProcedure, params object[] mArgumentos)
        {
            int Ejecutar_ProcedimientoRet = default(int);
            Ejecutar_ProcedimientoRet = nObj_Datos.Ejecutar_Procedure(mProcedure, mArgumentos);
            return Ejecutar_ProcedimientoRet;
        }

        // Ejecutar_Prodecimiento_Str

        public string Ejecutar_Procedimiento_Str(string mProcedure, params object[] mArgumentos)
        {
            string Ejecutar_Procedimiento_StrRet = default(string);
            Ejecutar_Procedimiento_StrRet = nObj_Datos.Ejecutar_Prodecimiento_Str(mProcedure, mArgumentos);
            return Ejecutar_Procedimiento_StrRet;
        }

        public DataSet Obtener_Listado(string mProcedure, params object[] mArgumentos)
        {
            DataSet Obtener_ListadoRet = default(DataSet);
            Obtener_ListadoRet = nObj_Datos.Traer_DataSet(mProcedure, mArgumentos);
            return Obtener_ListadoRet;
        }


        public DataSet Obtener_DataSet(string mProcedure, params object[] mArgumentos)
        {
            DataSet Obtener_DataSetRet = default(DataSet);
            Obtener_DataSetRet = nObj_Datos.Traer_DataSet(mProcedure, mArgumentos);
            return Obtener_DataSetRet;
        }


        public string Registro_Inserta_Id_Str(string mProcedure, params object[] mArgumentos)
        {
            string Registro_Inserta_Id_StrRet = default(string);
            Registro_Inserta_Id_StrRet = nObj_Datos.Ejecutar_Procedure_Out_Str(mProcedure, mArgumentos);
            return Registro_Inserta_Id_StrRet;
        }

        /* Agregado Jimmy 05-04-2022*/
        public int Ejecutar_Procedimiento_result_int(string mProcedure, params object[] mArgumentos)
        {
            int Ejecutar_ProcedimientoRet = default(int);
            Ejecutar_ProcedimientoRet = nObj_Datos.Ejecutar_Procedure_result_int(mProcedure, mArgumentos);
            return Ejecutar_ProcedimientoRet;
        }


    } // MyNegocio
}
