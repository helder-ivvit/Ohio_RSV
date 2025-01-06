using System;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Web;
using System.Data;
 
using System.Configuration;


namespace ModeloDatos
{


    public class MyDatos
    {
         // Coleccion de pares de claves ('Pws','123')
        System.Collections.Hashtable mColComandos = new System.Collections.Hashtable();

        public  System.Data.SqlClient.SqlConnection nConex = 
              new System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings["MyConnectionOhio"].ConnectionString.ToString());


        protected System.Data.SqlClient.SqlCommand Comando(string Proc_Almacenado)
        {

            System.Data.SqlClient.SqlCommand mComando;
            try
            {
                if (mColComandos.Contains(Proc_Almacenado))

                    mComando = (System.Data.SqlClient.SqlCommand)mColComandos[Proc_Almacenado];

                else
                {
                    nConex.Open();
                    mComando = new System.Data.SqlClient.SqlCommand(Proc_Almacenado, nConex);

                    mComando.Connection = nConex;
                    mComando.CommandType = CommandType.StoredProcedure;
                    mComando.CommandTimeout = 0;
                    //mComando.CommandTimeout = 240; //  60 x 4 = 240;

                    System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(mComando);

                    mColComandos.Add(Proc_Almacenado, mComando);  // pares de Claves
                    nConex.Close();
     //               nConex.Dispose();

                }
            }
            catch (System.Data.SqlClient.SqlException ex)
            {

                throw new Exception(" Error Comando : " + ex.Message);

            }

            return mComando;
        }
        //++++++++++++++++++++++

        private DataRow Obtener_Columna(DataTable nDT, string nNombre)
        {
            DataRow nRow;
            nRow = null;  // ' inicializado : 12/01/2014
            foreach (DataRow irow in nDT.Rows)
            {
                if (nNombre == irow["ROUTINE_NAME"].ToString())
                {
                    nRow = irow;
                    break;
                }
            }
            return nRow;
        }

        //+++++++++++++++++++++++++++++++++++++

        protected void Info_Procedure(string Proc_Almacenado, System.Data.SqlClient.SqlConnection nconex)
        {


            // nConex.Open()
            DataTable nDT_Procedure = nconex.GetSchema(Proc_Almacenado);
            DataColumn nDC_Procedure = nDT_Procedure.Columns["ROUTINE_NAME"];

            if (!(nDC_Procedure == null))
            {
                foreach (DataRow nrow in nDT_Procedure.Rows)
                {
                    string nProceName = nrow[nDC_Procedure].ToString();
                    DataTable nDT_Params = nconex.GetSchema("Procedureparameters", new string[] { "0", "0", nProceName });

                    DataColumn nParamNameDataColum = nDT_Params.Columns["PARAMETER_NAME"];
                    DataColumn nParamTypeDataColum = nDT_Params.Columns["DATA_TYPE"];

                    foreach (DataRow nParamRow in nDT_Params.Rows)
                    {
                        var nParamName = nParamRow[nParamNameDataColum].ToString();
                        var nParamType = nParamRow[nParamTypeDataColum].ToString();
                    }
                }
            }
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++++

        // Carga los parametros para el proceimiento almacenado
        protected void Cargar_Parametros(System.Data.SqlClient.SqlCommand nComando, ref object[] nArgs)
        {
            int i;
            {
                var withBlock = nComando;
                var loopTo = nArgs.GetUpperBound(0);
                for (i = 0; i <= loopTo; i++)
                {
                    try
                    {
                        ((System.Data.SqlClient.SqlParameter)withBlock.Parameters[i+1]).Value = nArgs[i];
                    }
                    catch (Exception ex)
                    {
                        // Throw (ex)
                        throw new Exception(" Error Cargar_Parametros : " + ex.Message);
                    }
                }
            }
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++

        public int Ejecutar_Procedure(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            int nResp;
            try
            {
                nConex.Open();
                mComand.Connection = nConex;
                mComand.CommandType = CommandType.StoredProcedure;
                Cargar_Parametros(mComand, ref mArgumentos);
                nResp = mComand.ExecuteNonQuery();

                nConex.Close();
                // nConex.Dispose();
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                nResp = -1;
                var xx = ex.ToString(); 
                nConex.Close();


            }
            return nResp;
        }
        //++++++++++++++++++++
        public string Ejecutar_Prodecimiento_Str(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            string nResp = "1";// System.Conversions.ToString[1];
            int i = 0;
            try
            {
                nConex.Open();
                mComand.Connection = nConex;
                mComand.CommandType = CommandType.StoredProcedure;
                Cargar_Parametros(mComand, ref mArgumentos);
                i = mComand.ExecuteNonQuery();
                nResp = i.ToString();
                nConex.Close();
                nConex.Dispose();
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                nResp = ex.Message;
                throw new Exception(" Error Ejecutar_procedure : " + ex.Message + " --> " + mProc_Almacenado);

                nConex.Close();
            }
            return nResp;
        }
        //++++++++++++++++++++++++++++++++++++++++++

        public string Ejecutar_Procedure_Test(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            string nResp = "1"; //Conversions.ToString(1);
            int i = 0;
            try
            {
                nConex.Open();
                mComand.Connection = nConex;
                mComand.CommandType = CommandType.StoredProcedure;
                Cargar_Parametros(mComand, ref mArgumentos);
                i = mComand.ExecuteNonQuery();
                nResp = i.ToString();

                nConex.Close();
                nConex.Dispose();
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                nResp = ex.Message;
                nConex.Close();
                throw new Exception(" Error Ejecutar_procedure : " + ex.Message + " --> " + mProc_Almacenado);


            }
            return nResp;
        }

        //++++++++++++++++++++++++++++++++++++
        // para cargar parametros del procedimiento almacenado que devuelve un  parametro
        protected void Cargar_Parametros_Out(System.Data.SqlClient.SqlCommand nComando, ref object[] nArgs)
        {
            int i;
            {
                var withBlock = nComando;
                var loopTo = nArgs.GetUpperBound(0) // - 1
               ;
                for (i = 0; i < loopTo; i++)
                {
                    try
                    {
                        ((System.Data.SqlClient.SqlParameter)withBlock.Parameters[i]).Value = nArgs[i];
                    }
                    catch (Exception ex)
                    {
                        // Throw (ex)
                        throw new Exception(" Error Cargar_Parametros_Out : " + ex.Message);
                    }
                }
                // i = i + 1

 
                if (nComando.Parameters.Contains("ValorReturn"))
                {
                }
                else
                    nComando.Parameters.Add(
                        new System.Data.SqlClient.SqlParameter("ValorReturn", System.Data.SqlDbType.Int  )).Direction = ParameterDirection.Output;
            }
        }

        //++++++++++++++++++++++++++++++++
        public int Ejecutar_Procedure_Out(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            int nResp = 0;
            //string sResp = "";
            string Error;
            try
            {
                nConex.Open();
                mComand.Connection = nConex;
                mComand.CommandType = CommandType.StoredProcedure;
                Cargar_Parametros_Out(mComand, ref mArgumentos);

                // return (bool)cmd.Parameters[":elimino"].Value;
                nResp = mComand.ExecuteNonQuery(); // .ExecuteNonQuery'   ExecuteScalar
                                                   // ValorReturn
                nResp = (int)mComand.Parameters["ValorReturn"].Value; //.Value();
                nConex.Close();
                nConex.Dispose();
                return nResp;
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                Error = ex.Message;

            }

            return 0;
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++

        protected System.Data.SqlClient.SqlDataAdapter CrearDataAdapter(string mProc_Almacenado, params object[] nArgs)  // IDataAdapter
        {

            try {

                System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
                // Not Args IsNothing 
                if (nArgs.Length > 0)
                    Cargar_Parametros(mComand, ref nArgs);
                return new System.Data.SqlClient.SqlDataAdapter(mComand);

            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                var x = ex.ToString();

            }
            return null;
        }
        //+++++++++++++++++++++++++++++++++++++
        // TRAER DATOS SIN PARAMETROS
        public new DataSet Traer_DataSet(string mProc_Amacenado)
        {
       

            try {
                var mDataSet = new DataSet();
                CrearDataAdapter(mProc_Amacenado).Fill(mDataSet);
                return mDataSet;

            }
            catch (Exception ex)
            {
                var neax = ex.ToString(); 
            }
            return null;
        }
        //++++++++++++++++++++++++++++++++++++++++
        // TRAER DATOS CON PARAMETROS
        public new DataSet Traer_DataSet(string mProc_Almacenado, params object[] nArgs)
        {
            var mDataSet = new DataSet();
            // aqui aqui 
            try
            {
                
                CrearDataAdapter(mProc_Almacenado, nArgs).Fill(mDataSet);
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                var nex = ex.ToString();
            }

            return mDataSet;
        }
        //++++++++++++++++++++++++++++++++++++++
        public System.Data.SqlClient.SqlDataAdapter TraerDataAdapter(string mProc_Almacenado, params object[] nArgs)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            if (nArgs == null)  // IsNothing[nArgs]
                Cargar_Parametros(mComand, ref nArgs);

            return new System.Data.SqlClient.SqlDataAdapter(mComand);
        }
        //++++++++++++++++++++++++++++
        public void nDeriveParameters(System.Data.SqlClient.SqlCommand nComando)
        {
            string nComandText = nComando.CommandText;

            if (nComandText.IndexOf(".") == -1)
                nComandText = nComando.Connection.Database + "." + nComandText;
        }
        //+++++++++++++++++++++++++++

        protected void Cargar_Parametros_Out_Str(System.Data.SqlClient.SqlCommand nComando, ref object[] nArgs)
        {
            int i;
            {
                var withBlock = nComando;
                var loopTo = nArgs.GetUpperBound(0) // - 1
               ;
                for (i = 0; i <= loopTo; i++)
                {
                    try
                    {
                        ((System.Data.SqlClient.SqlParameter)withBlock.Parameters[i]).Value = nArgs[i];
                    }
                    catch (Exception ex)
                    {
                        // Throw (ex)
                        throw new Exception(" Error Cargar_Parametros_Out : " + ex.Message);
                    }
                }

                // ValorReturn  111
                if (nComando.Parameters.Contains("ValorReturn"))
                {
                }
                else
                    nComando.Parameters.Add(new System.Data.SqlClient.SqlParameter("ValorReturn", System.Data.SqlDbType.VarChar  )).Direction = ParameterDirection.Output;
            }
        }
        //++++++++++++++++++++++++++++

        public string Ejecutar_Procedure_Out_Str(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            string nResp;
            int i = 0;
            nConex.Open();
            mComand.Connection = nConex;
            mComand.CommandType = CommandType.StoredProcedure;
            Cargar_Parametros_Out_Str(mComand, ref mArgumentos);


            try
            {
                i = mComand.ExecuteNonQuery(); // .ExecuteNonQuery'   nResp
                nResp = i.ToString();
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                throw new Exception(" Error Ejecutar_procedure : " + ex.Message + " --> " + mProc_Almacenado);
            }

            // ValorRetu rn

            // ValorReturn
            nResp = mComand.Parameters["ValorReturn"].ToString(); //.Value;  // Nombre de campo definido como salida
                                                                  // mComand.Parameters.Remove("ValorReturn")
            nConex.Close();
            nConex.Dispose();

            return nResp;
        }

        /* Agregado Jimmy 05-04-2022*/
        public int Ejecutar_Procedure_result_int(string mProc_Almacenado, params object[] mArgumentos)
        {
            System.Data.SqlClient.SqlCommand mComand = Comando(mProc_Almacenado);
            int nResp;
            try
            {
                nConex.Open();
                mComand.Connection = nConex;
                mComand.CommandType = CommandType.StoredProcedure;
                Cargar_Parametros(mComand, ref mArgumentos);
                nResp = (int)mComand.ExecuteScalar();

                nConex.Close();
                // nConex.Dispose();
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                nResp = -1;
                var xx = ex.ToString();
                nConex.Close();


            }
            return nResp;
        }


    }  //   MyDatos


}
