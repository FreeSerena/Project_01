using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
    public class ExcelHelper
    {
        private string officeVersion;

        private string filePath;

        private string connStr;

        private ArrayList sheetNameAL;

        #region 获取office版本号
        /// <summary>
        /// 获取Office版本号
        /// </summary>
        /// <returns></returns>
        public string GetOfficeVersion()
        {
            Type type;
            object excel;
            object version = null;

            type = Type.GetTypeFromProgID("Excel.Application");

            if (type == null)
            {
                return "没有安装excel";
            }
            else
            {
                excel = Activator.CreateInstance(type);
                if (excel == null)
                {
                    return "创建对象出错";
                }
                else
                {
                    version = type.GetProperty("Version").GetValue(excel, null);
                    type.GetProperty("Visible").SetValue(excel, false, null);
                    type.GetMethod("Quit").Invoke(excel, null);
                    if (version != null)
                    {
                        //Excel版本号
                        officeVersion = version.ToString();
                        return string.Empty;
                    }
                    else
                    {
                        return "未知错误";
                    }
                }
            }
        }
        #endregion

        #region 获取文件连接字符
        public string GetExcelConnStr(string excelFilePath)
        {
            if (string.IsNullOrEmpty(officeVersion))
            {
                return "没有安装Office！";
            }
            if (string.IsNullOrEmpty(excelFilePath))
            {
                return "文件路径为空！";
            }

            filePath = excelFilePath;
            switch (officeVersion)
            {
                case "12.0"://office 2007
                case "14.0"://office 2010
                    connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filePath + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'");
                    break;
                case "11.0"://office 2003
                default://默认为office2003
                    connStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + filePath + "';Extended Properties='Excel 8.0;HDR=YES;IMEX=1'");
                    break;
            }
            return string.Empty;
        }
        #endregion

        #region 获取excel sheet
        public string GetSheetName()
        {
            if (string.IsNullOrEmpty(connStr))
            {
                return "excel 连接为空！";
            }
            var conn = new OleDbConnection(connStr);
            try
            {
                conn.Open();
                System.Data.DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                sheetNameAL = new ArrayList();
                object temp;
                foreach (DataRow item in dtSheetName.Rows)
                {
                    temp = item["TABLE_NAME"];
                    if (temp != null)
                    {
                        sheetNameAL.Add(temp.ToString().Replace("$", string.Empty));
                    }
                }
                if (dtSheetName.Rows.Count <= 0)
                {
                    return "Excel 文件为空！";
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception error)
            {
                return error.Message;
            }
            finally
            {
                conn.Close();
            }
        }
        #endregion

        #region 从excel表格中读取数据
        public string GetExcelData(out DataSet ds)
        {
            ds = new DataSet();
            if (string.IsNullOrEmpty(connStr))
            {
                return "Excel 连接为空";
            }
            if (sheetNameAL.Count <= 0)
            {
                return "Excel 文件为空！";
            }

            var conn = new OleDbConnection(connStr);
            conn.Open();
            try
            {
                string sqlText;
                System.Data.DataTable dt;
                foreach (var item in sheetNameAL)
                {
                    if (item != null && !string.IsNullOrEmpty(item.ToString()))
                    {
                        sqlText = "select * from [" + item.ToString() + "$]";
                        var oleDBAD = new OleDbDataAdapter(sqlText, conn);
                        dt = new System.Data.DataTable();
                        oleDBAD.Fill(dt);
                        dt.TableName = item.ToString();
                        ds.Tables.Add(dt);
                    }
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                conn.Close();
            }

        }
        #endregion
    }
}
