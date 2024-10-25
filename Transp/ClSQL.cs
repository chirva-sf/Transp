using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Transp
{
    class ClSQL
    {
        private static SqlConnection ClCN;
        public static bool Error;

        public ClSQL(string ConnectionString)
        {
            ConnectSQL(ConnectionString);
        }

        public void ConnectSQL(string ConnectionString)
        {
            Error = false;
            try
            {
                ClCN = new SqlConnection(ConnectionString);
                ClCN.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Error = true;
            }
        }

        public void DisconnectSQL()
        {
            ClCN.Close();
        }

        public DataTable SelectSQL(string strSQL)
        {
            Error = false;
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(strSQL, ClCN);
                DataSet ds = new DataSet("clsql");
                da.Fill(ds, "cl_sql");
                return ds.Tables["cl_sql"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Error = true;
                return null;
            }
        }

        public DataRow SelectRow(string strSQL)
        {
            Error = false;
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(strSQL, ClCN);
                DataSet ds = new DataSet("clsql");
                da.Fill(ds, "cl_sql");
                DataTable dt = ds.Tables["cl_sql"];
                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Error = true;
                return null;
            }
        }

        public string SelectCell(string strSQL)
        {
            Error = false;
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(strSQL, ClCN);
                DataSet ds = new DataSet("clsql");
                da.Fill(ds, "cl_sql");
                DataTable dt = ds.Tables["cl_sql"];
                string str = "";
                if (dt.Rows.Count > 0)
                {
                    str = dt.Rows[0][0].ToString();
                }
                return str;
            }
            catch (Exception ex)
            {
                Error = true;
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public bool CheckConnection()
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select 1", ClCN);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int SelectIntCell(string strSQL)
        {
            string str = SelectCell(strSQL);
            int intVal;
            if (!int.TryParse(str, NumberStyles.Integer, null, out intVal)) intVal = 0;
            return intVal;
        }

        public double SelectDoubleCell(string strSQL)
        {
            string str = SelectCell(strSQL);
            double doubleVal;
            if (!Double.TryParse(str, (NumberStyles.Float | NumberStyles.AllowThousands), null, out doubleVal)) doubleVal = 0;
            return doubleVal;
        }

        public DateTime SelectDateCell(string strSQL)
        {
            string str = SelectCell(strSQL);
            DateTime dtVal;
            if (!DateTime.TryParse(str, out dtVal)) dtVal = DateTime.MinValue;
            return dtVal;
        }

        public void ExecuteSQL(string strSQL)
        {
            Error = false;
            try
            {
                SqlCommand cmd = new SqlCommand(strSQL, ClCN);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Error = true;
            }
        }

    }
}
