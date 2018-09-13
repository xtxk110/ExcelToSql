using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.Data.Common;
using MySql.Data.MySqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

/// <summary>
/// 获取各数据库类型链接,默认MSSQL
/// </summary>
public class DB
{
    private static readonly DB db = new DB();

    public static DB Instance()
    {
        return db;
    }

    public IDbConnection CreateConnection(DbType dbType,string connStr)
    {
        IDbConnection connection = null;
        try
        {
            switch (dbType)
            {
                case DbType.MSSQL:
                    connection = new SqlConnection(connStr);
                    break;
                case DbType.MYSQL:
                    connection = new MySqlConnection(connStr);
                    break;
                case DbType.ORACELE:
                    connection = new OracleConnection(connStr);
                    break;
                case DbType.ACCESS:
                    connection = new OleDbConnection(connStr);
                    break;
                default:
                    connection = new SqlConnection(connStr);
                    break;
            }
            connection.Open();
        }catch(Exception e) { throw e; }
        return connection; 
    }
}
/// <summary>
/// 数据库操作类
/// </summary>
public class DbHelper
{
    /// <summary>
    /// 生成以日期(yyyyMMddHHmm)开头的16位随机数字符
    /// </summary>
    /// <returns></returns>
    public static string GetNewId()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(DateTime.Now.ToString("yyyyMMddHHmm").ToString());
        Random rd = new Random(Guid.NewGuid().GetHashCode());
        sb.Append(rd.Next(1000, 9999).ToString());

        return sb.ToString();
    }
    /// <summary>
    /// 数据库连接字符串
    /// </summary>
    public static string ConnStr = ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
    #region MSSQL
    /// <summary>
    /// MSSQL批量插入数据库
    /// </summary>
    /// <param name="dt">待插入数据表</param>
    /// <param name="table">目的数据表名称</param>
    /// <param name="message">操作输出信息</param>
    public static void InsetBatch(DataTable dt,string table,out string message)
    {
        message = "操作成功";
        string sql = " SELECT TOP 0 * FROM "+table;
        DataTable destination = new DataTable("table");
        try
        {
            using(SqlConnection conn=DB.Instance().CreateConnection(DbType.MSSQL,ConnStr) as SqlConnection)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                SqlCommandBuilder scb = new SqlCommandBuilder(sda);
                sda.InsertCommand = scb.GetInsertCommand();
                sda.UpdateCommand = scb.GetUpdateCommand();
                sda.DeleteCommand = scb.GetDeleteCommand();
                sda.Fill(destination);

                if(destination.Columns.Count == dt.Columns.Count)
                {
                    foreach(DataRow item in dt.Rows)
                    {
                        destination.Rows.Add(item.ItemArray);
                    }
                }else if (destination.Columns.Count == dt.Columns.Count + 1)
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        object[] dataRow = new object[item.ItemArray.Length + 1];
                        dataRow[0] = GetNewId();
                        item.ItemArray.CopyTo(dataRow, 1);
                        destination.Rows.Add(dataRow);
                    }
                }
                else
                {
                    message = "导入数据列与目标数据列不等,请检查";
                    return;
                }
                sda.UpdateBatchSize = 1000;
                sda.Update(destination);
            }
        }catch(Exception e) { message = e.Message; }
    }

    #endregion
}

