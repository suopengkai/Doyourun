using System.Collections.Generic;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;

namespace WindowsFormsExe导入
{
	/// <summary>
	/// 第三方
	/// </summary>
	public  static class NPOHelp
    {

        ///返回datatable   static 静态类才能调用
        public  static  DataTable  ExcelToDataTable(string FilName)
        {
            ///想实例化 有返回值
            var dt = new DataTable();
            //实例化 工作表
            IWorkbook work = null;
            ///文件后缀名字  .txt .xls  .xlsx
            var filExt = Path.GetExtension(FilName);
            ///打开 Excel
            using (FileStream fsw = File.OpenRead(FilName))
            {
                if (filExt==".xls") //格式
                {
                    //读取  .xls
                    work = new HSSFWorkbook(fsw);
                }
                else  if ( filExt==".xlsx")
                {
                    work = new HSSFWorkbook(fsw);

                }
                else
                {
                    return null;
                }
                //0列 编号  第一页
                ///头行
                var first = work.GetSheetAt(0);
                var header = first.GetRow(first.FirstRowNum);
                List<int> Clounnum = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    ///获取列偷单元格
                    var obj = GetCellValue(header.GetCell(i));
                  
                    ///如果没列头自己  定义名字
                    if (obj == null) obj = "Clounnum" + i;
                    ///添加自己的
                    dt.Columns.Add(new DataColumn(obj.ToString()));
                    Clounnum.Add(i);
                }
                ///获取除第一行外的行数  FirstRowNum 
                for (int i = (first.FirstRowNum + 1); i < first.LastRowNum; i++)
                {
                    ///新的行
                    DataRow dr = dt.NewRow();
                    foreach (var item in  Clounnum)
                    {
                        dr[item] = GetCellValue(first.GetRow(i).GetCell(item));

                    }
                    dt.Rows.Add(dr);
                }

                return dt;
            }


        }

        public static object GetCellValue(ICell cell)
        {
            if (cell == null) return null;

            switch (cell.CellType)
            {
                case CellType.Blank:  //空白
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;

                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula://公式

                default: return "=" + cell.CellFormula;
                   
            }

        }

    }
}
