using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ExcelHelper
    {
        public static readonly ExcelHelper Default = new ExcelHelper();
        
        /// <summary>
        /// 数据源（Excel）链接对象
        /// </summary>
        public OleDbConnection getcon(string path)
        {
            string M_str_Oledbcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "; Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            OleDbConnection myCon = new OleDbConnection(M_str_Oledbcon);
            
            return myCon;
        }


        /// <summary>
        /// 读取数据 存储在DataTable对象中
        /// </summary>
        /// <param name="M_str_sqlstr">Sql 命令</param>
        public DataTable getTable(string M_str_sqlstr, string path)
        {
            OleDbConnection oleConn = this.getcon(path);
            OleDbDataAdapter myDa = new OleDbDataAdapter(M_str_sqlstr, oleConn);
            
            DataTable dt = new DataTable();
            myDa.Fill(dt);
            return dt;
        }

        /// <summary>
        /// 获取该数据源Excel的所有工作表（Sheet）
        /// </summary>
        public DataTable getSheets(string path)
        {
            OleDbConnection oleConn = this.getcon(path);
            DataTable dt = null;

            try
            {
                oleConn.Open();
                dt = oleConn.GetSchema("Tables");
            }

            catch (Exception e)
            {
                throw new ApplicationException("Error:" + e.Message);
            }
            finally
            {
                oleConn.Close();
            }

            return dt;

        }
        
        /// <summary>
        /// 获取数据 存储在OleDataReader中
        /// </summary>
        /// <param name="M_str_sqlstr">Sql 语句</param>
        /// <returns></returns>
        public OleDbDataReader getDataRead(string M_str_sqlstr,string path)
        {
            OleDbConnection oleConn = this.getcon(path);
            OleDbCommand myCom = new OleDbCommand(M_str_sqlstr, oleConn);
            oleConn.Open();
            OleDbDataReader myRead = myCom.ExecuteReader(CommandBehavior.CloseConnection);

            return myRead;
        }



    }
}
