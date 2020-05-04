using System;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Configuration;

namespace DReportUtil.DBControl
{
    public class DBConnecter
    {
        private string TAG = "DBCONTROL";
        OracleConnection conn;
        public DBConnecter()
        {
            string dbip1 = ConfigurationManager.AppSettings["dbIP"];
            string dbIns = ConfigurationManager.AppSettings["dbInstance"];
            string constr = string.Format(@"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={0})(PORT=1521))(CONNECT_DATA=(SERVICE_NAME={1})));User Id=uni_dish;Password=F8zV7zeT8wY2awXa;", dbip1, dbIns);
            conn = new OracleConnection(constr);
        }
        public string iOpen()
        {
            string retVal = string.Empty;
            try
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                    retVal = "0000";
                }
                else
                {
                    retVal = "0000";
                }
            }
            catch (Exception ex)
            {
                retVal = string.Format("FFFFx[{0}]", ex.Message);
            }
            return retVal;
        }
        public string iClose()
        {
            string retVal = string.Empty;
            try
            {
                if (conn.State != ConnectionState.Open)
                {
                    retVal = "0000";
                }
                else
                {
                    conn.Close();
                    retVal = "0000";
                }
            }
            catch (Exception ex)
            {
                retVal = string.Format("FFFFx[{0}]", ex.Message);
            }
            return retVal;
        }
        public bool idbCheck(out string sttCode)
        {
            sttCode = string.Empty;
            bool retVal = false;
            try
            {
                string stateCode = iOpen();
                if (stateCode == "0000")
                {
                    string cclose = iClose();
                    if (cclose == "0000")
                    {
                        sttCode = string.Empty;
                        retVal = true;
                    }
                    else
                    {
                        retVal = false;
                        sttCode = cclose;
                    }
                }
                else
                {
                    retVal = false;
                    sttCode = stateCode;
                }
            }
            catch (Exception ex)
            {
                retVal = false;
                sttCode = string.Format("FFFFx[{0}]", ex.Message);
            }
            return retVal;
        }
        public string iDBCommand(string qry)
        {
            string retVal = string.Empty;
            try
            {
                string op = iOpen();
                OracleCommand cmd = new OracleCommand(qry, conn);
                int stt = cmd.ExecuteNonQuery();
                string cl = iClose();
                retVal = string.Format(@"{0}{1}", stt, "000");
            }
            catch (Exception ex)
            {
                retVal = string.Format("FFFFx[{0}]", ex.Message);
            }
            return retVal;
        }
        public DataTable getTable(string qry)
        {
            DataTable dt = new DataTable();
            try
            {
                if (qry.Length > 0)
                {
                    OracleDataAdapter da = new OracleDataAdapter(qry, conn);
                    da.Fill(dt);
                }
                else
                {
                    dt.Clear();
                }
            }
            catch (Exception ex)
            {
                //SystemLogWriter.LogWriter._error(TAG, ex.ToString());
                dt.Clear();
            }
            return dt;
        }
    }
}
