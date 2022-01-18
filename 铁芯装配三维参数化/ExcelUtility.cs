using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.Eval;
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
        internal IWorkbook GetWorkbook(string filePath)
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
                return wk;
                //读取当前表数据
                //ISheet sheet = wk.GetSheetAt(0);
                
                //IRow row = sheet.GetRow(0);

                //row = sheet.GetRow(0);
                //string value = row.GetCell(0).ToString();
                //return value;

            }

            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }
        }
        internal ISheet GetWKSheet(IWorkbook workbook, int sheetName) 
        {
            return workbook.GetSheetAt(sheetName);
        }
        internal object GetCellValue(ISheet sheet, string location)
        {
            var cr = new CellReference(location);
            var row = sheet.GetRow(cr.Row);
            var cell = row.GetCell(cr.Col);

            if (cell == null)
            {
                return string.Empty;
            }
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Error:
                    return ErrorEval.GetText(cell.ErrorCellValue);

                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Boolean:
                            return cell.BooleanCellValue;

                        case CellType.Error:
                            return ErrorEval.GetText(cell.ErrorCellValue);

                        case CellType.Numeric:
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                return cell.DateCellValue.ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                return cell.NumericCellValue;
                            }
                        case CellType.String:
                            string str = cell.StringCellValue;
                            if (!string.IsNullOrEmpty(str))
                            {
                                return str.ToString();
                            }
                            else
                            {
                                return string.Empty;
                            }
                        case CellType.Unknown:
                        case CellType.Blank:
                        default:
                            return string.Empty;
                    }
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.String:
                    string strValue = cell.StringCellValue;
                    return strValue.ToString().Trim();

                case CellType.Unknown:
                case CellType.Blank:
                default:
                    return string.Empty;
            }
        }

        public void LoadData(string filePath)
        {
            IWorkbook countBook = GetWorkbook(filePath);
            ISheet firstSheet = GetWKSheet(countBook, 0);
            ISheet secondSheet = GetWKSheet(countBook, 1);

            string 型号 = GetCellValue(firstSheet, "B1").ToString()+ GetCellValue(firstSheet, "C1").ToString()+"-"+ GetCellValue(firstSheet, "D1").ToString()+ "/"+GetCellValue(firstSheet, "F1").ToString();
            string 图号 = GetCellValue(firstSheet, "H1").ToString();
            int 容量 = Convert.ToInt32(GetCellValue(firstSheet, "D1"));
            int 低压匝数 = Convert.ToInt32(GetCellValue(firstSheet, "E7"));
            int 低压电压 = Convert.ToInt32(GetCellValue(firstSheet, "E3"));
            int 空载损耗 = Convert.ToInt32(GetCellValue(secondSheet, "B13"));
            string 连接组别;
            if (Convert.ToInt32(GetCellValue(firstSheet, "J3")) == 1)
            {
                连接组别 = "Yyn0";
            }
            else if(Convert.ToInt32(GetCellValue(firstSheet, "J4")) == 1)
            {
                连接组别 = "Dyn11";
            }
            else if (Convert.ToInt32(GetCellValue(firstSheet, "J5")) == 1)
            {
                连接组别 = "Dd0";
            }
            else
            {
                连接组别 = "Yd11";
            }
            
            int 轨距 = 820;
            int MO = Convert.ToInt32(GetCellValue(firstSheet, "H26"));
            int 窗高 = Convert.ToInt32(GetCellValue(firstSheet, "G40"));
            double 叠片系数 = Convert.ToDouble(GetCellValue(secondSheet, "B6"));
            double 有效截面积 = Convert.ToDouble(GetCellValue(secondSheet, "B3"));
            string 硅钢片 = GetCellValue(secondSheet, "B5").ToString();

            Console.WriteLine(连接组别);
            //int 轭片长 = (int)GetCellValue(firstSheet, "B5");
            //int 轭片宽 = (int)GetCellValue(firstSheet, "B5");
            //float K值 = (float)GetCellValue(firstSheet, "B5");
            //float 高压叠厚 = (float)GetCellValue(firstSheet, "B5");
            //float 低压叠厚 = (float)GetCellValue(firstSheet, "B5");
            //int 柱片宽 = (int)GetCellValue(firstSheet, "B5");
            //int B柱片长 = (int)GetCellValue(firstSheet, "B5");
            //int AC柱片长 = (int)GetCellValue(firstSheet, "B5");
        }
    }
}
