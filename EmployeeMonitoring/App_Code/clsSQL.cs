using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using System.Text;
using System.IO;

/// <summary>
/// Summary description for clsSQL
/// </summary>
public class clsSQL
{
	public clsSQL()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    public DataTable Bind(string strSql, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        Execute SQL Query ใส่ DataTable
        clsSQL.BindDT("SELECT * FROM member","MySQL","cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        DataTable dt = new DataTable();

        if (!string.IsNullOrEmpty(csSQL))
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlDataAdapter myDa_SQL = new SqlDataAdapter(strSql, myConn_SQL);

                myDa_SQL.Fill(dt);
                myConn_SQL.Dispose();
                myDa_SQL.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcDataAdapter myDa_ODBC = new OdbcDataAdapter(strSql, myConn_ODBC);

                myDa_ODBC.Fill(dt);
                myConn_ODBC.Dispose();
                myDa_ODBC.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlDataAdapter myDa_SQL = new SqlDataAdapter(strSql, myConn_SQL);

                myDa_SQL.Fill(dt);
                myConn_SQL.Dispose();
                myDa_SQL.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
        }
        else
        {
            return null;
        }
    }

	public DataTable Bind(string strSql, string[,] arrParameter, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        รัน Query ด้วย Parameter ใส่ DataTable
        strSQL.Append("SELECT email FROM member WHERE id=?ID");
        DataTable dt = new DataTable();
        dt = Bind(strSQL.ToString(), new string[,] { { "?ID", txtTest.Text } }, "MySQL", "cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        DataTable dt = new DataTable();
        int i;

        if (!string.IsNullOrEmpty(csSQL))
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlDataAdapter myDa_SQL = new SqlDataAdapter(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myDa_SQL.SelectCommand.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                myDa_SQL.Fill(dt);
                myConn_SQL.Dispose();
                myDa_SQL.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcDataAdapter myDa_ODBC = new OdbcDataAdapter(strSql, myConn_ODBC);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myDa_ODBC.SelectCommand.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                myDa_ODBC.Fill(dt);
                myConn_ODBC.Dispose();
                myDa_ODBC.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlDataAdapter myDa_SQL = new SqlDataAdapter(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myDa_SQL.SelectCommand.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                myDa_SQL.Fill(dt);
                myConn_SQL.Dispose();
                myDa_SQL.Dispose();
                if (dt.Rows.Count > 0 && dt != null)
                {
                    return dt;
                }
                else
                {
                    dt.Dispose();
                    return null;
                }
            }
        }
        else
        {
            return null;
        }
    }
	
    public string Return(string strSql, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        รัน Query เพื่อคืนค่าค่าเดียว
        clsSQL.Return("SELECT MAX(id) FROM member","MySQL","cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        string strReturn = "";

        if (!string.IsNullOrEmpty(csSQL))
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);
                try
                {
                    myConn_SQL.Open();
                    strReturn = myCmd_SQL.ExecuteScalar().ToString();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcCommand myCmd_ODBC = new OdbcCommand(strSql, myConn_ODBC);
                try
                {
                    myConn_ODBC.Open();
                    strReturn = myCmd_ODBC.ExecuteScalar().ToString();
                    myConn_ODBC.Close();
                    myCmd_ODBC.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_ODBC.Dispose();
                    myConn_ODBC.Close();
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);
                try
                {
                    myConn_SQL.Open();
                    strReturn = myCmd_SQL.ExecuteScalar().ToString();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
        }
        return strReturn;
    }

	public string Return(string strSql, string[,] arrParameter, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        รัน Query เพื่อคืนค่าค่าเดียว ผ่านการส่งค่าด้วย Parameter
        strSQL.Append("SELECT email FROM member WHERE id=?ID");
        lblMessage.Text = clsSQL.Return(strSQL.ToString(), new string[,] { { "?ID", txtTest.Text } }, "MySQL", "cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        string strReturn = "";
        int i;

        if (!string.IsNullOrEmpty(csSQL))
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_SQL.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                try
                {
                    myConn_SQL.Open();
                    strReturn = myCmd_SQL.ExecuteScalar().ToString();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcCommand myCmd_ODBC = new OdbcCommand(strSql, myConn_ODBC);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_ODBC.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                try
                {
                    myConn_ODBC.Open();
                    strReturn = myCmd_ODBC.ExecuteScalar().ToString();
                    myConn_ODBC.Close();
                    myCmd_ODBC.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_ODBC.Dispose();
                    myConn_ODBC.Close();
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_SQL.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }
                try
                {
                    myConn_SQL.Open();
                    strReturn = myCmd_SQL.ExecuteScalar().ToString();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
        }
        return strReturn;
    }
	
    public bool Execute(string strSql, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        Execute คำสั่ง Query
        clsSQL.Execute("DELETE FROM member WHERE id=1","MySQL","cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        bool boolReturn;

        if (!string.IsNullOrEmpty(csSQL))
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);
                try
                {
                    myConn_SQL.Open();
                    myCmd_SQL.ExecuteNonQuery();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcCommand myCmd_ODBC = new OdbcCommand(strSql, myConn_ODBC);
                try
                {
                    myConn_ODBC.Open();
                    myCmd_ODBC.ExecuteNonQuery();
                    myConn_ODBC.Close();
                    myCmd_ODBC.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_ODBC.Dispose();
                    myConn_ODBC.Close();
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);
                try
                {
                    myConn_SQL.Open();
                    myCmd_SQL.ExecuteNonQuery();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
        }
        else
        {
            boolReturn = false;
        }
        return boolReturn;
    }
	
	public bool Execute(string strSql, string[,] arrParameter, string strDBType, string appsetting_name)
    {
        #region Remark
        /*############################ Example ############################
        Execute คำสั่ง Query ส่งค่าด้วย Parameter
        clsSQL.Execute("UPDATE webboard_type SET type_name=?NAME WHERE type_id=?ID", new string[,] { { "?ID", txtTest.Text }, { "?NAME", "ใช้ Array 2 มิติ" } }, "MySQL", "cs");
        #################################################################*/
        #endregion

        string csSQL = System.Configuration.ConfigurationManager.AppSettings[appsetting_name];
        bool boolReturn;
        int i;

        if (!string.IsNullOrEmpty(csSQL) && arrParameter.Rank==2)
        {
            if (strDBType.ToLower() == "sql")
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_SQL.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }

                try
                {
                    myConn_SQL.Open();
                    myCmd_SQL.ExecuteNonQuery();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
            else if (strDBType.ToLower() == "odbc")
            {
                OdbcConnection myConn_ODBC = new OdbcConnection(csSQL);
                OdbcCommand myCmd_ODBC = new OdbcCommand(strSql, myConn_ODBC);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_ODBC.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }

                try
                {
                    myConn_ODBC.Open();
                    myCmd_ODBC.ExecuteNonQuery();
                    myConn_ODBC.Close();
                    myCmd_ODBC.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_ODBC.Dispose();
                    myConn_ODBC.Close();
                }
            }
            else
            {
                SqlConnection myConn_SQL = new SqlConnection(csSQL);
                SqlCommand myCmd_SQL = new SqlCommand(strSql, myConn_SQL);

                for (i = 0; i < arrParameter.Length / arrParameter.Rank; i++)
                {
                    myCmd_SQL.Parameters.AddWithValue(arrParameter[i, 0], arrParameter[i, 1]);
                }

                try
                {
                    myConn_SQL.Open();
                    myCmd_SQL.ExecuteNonQuery();
                    myConn_SQL.Close();
                    myCmd_SQL.Dispose();
                    boolReturn = true;
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    boolReturn = false;
                }
                finally
                {
                    myCmd_SQL.Dispose();
                    myConn_SQL.Close();
                }
            }
        }
        else
        {
            boolReturn = false;
        }
        return boolReturn;
    }
}
