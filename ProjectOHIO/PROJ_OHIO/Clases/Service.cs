using System;
using System.Data.OleDb;
using System.Data;

namespace PROJ_OHIO.Clases
{
    public class Service
    {        
        public DataTable ImportExceltoDatable(string connString="")
        {
            Console.WriteLine("Jimmy exporta");
            OleDbConnection oledbConn = new OleDbConnection(connString);
            DataTable dt = new DataTable();
            try
            {
                oledbConn.Open();
                using (OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn))
                {
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    oleda.SelectCommand = cmd;
                    DataSet ds = new DataSet();
                    oleda.Fill(ds);

                    dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        //table tblObj = new table();
                        foreach (DataRow row in dt.Rows)
                        {
                            //tblObj.Name = row["Name"].ToString();
                            //tblObj.Name = row["Address"].ToString();
                            //tblObj.Salary = (int)row["Salary"];
                            //tblObj.Age = (int)row["Age"];
                        }
                    }

                    oledbConn.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                dt = null;
                oledbConn = null;
            }
            finally
            {
                //oledbConn.Close();                
                oledbConn = null;
            }

            return dt;
        }

    }
}