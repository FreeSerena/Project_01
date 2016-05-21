using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
    public class PointBLL
    {
        public ExcelHelper excelHelper = new ExcelHelper();
        public void CheckPoint(string filePath, out DataSet dt, Dictionary<int, bool> validateColumns)
        {
            string msg = string.Empty;

            //检测Office版本
            msg = excelHelper.GetOfficeVersion();
            if (!string.IsNullOrEmpty(msg))
            {
                throw new ApplicationException(msg);
            }

            //检测上传的Excel是否存在
            msg = excelHelper.GetExcelConnStr(filePath);
            if (!string.IsNullOrEmpty(msg))
            {
                throw new ApplicationException(msg);
            }

            msg = excelHelper.GetSheetName();

            if (!string.IsNullOrEmpty(msg))
            {
                throw new ApplicationException(msg);
            }
        }
    }
}
