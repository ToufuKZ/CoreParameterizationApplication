using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 铁芯装配三维参数化
{
    public class ExcelUtility
    {
        public string ReadFromExcelFile(string filePath)
        {
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(fs);
                }

                fs.Close();
                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);

                IRow row = sheet.GetRow(0);  //读取当前行数据
                                             //LastRowNum 是当前表的总行数-1（注意）

                row = sheet.GetRow(0);
                string value = row.GetCell(0).ToString();
                return value;
            }

            catch (Exception e)
            {
                //只在Debug模式下才输出
                return e.Message;
            }
        }
    }
}
