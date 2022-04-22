
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace Doyourun
{
	public static class Helper
	{
        public static DataTable ConvertStreamToDataTable(Stream stream, string fileType, string sheetName, bool isFirstRowColumn)
        {
            IWorkbook workbook = null;
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;

            if (fileType == "xlsx")
            {
                workbook = new XSSFWorkbook(stream);//2007版本
            }
            else if (fileType == "xls")
            {
                try
                {
                    workbook = new HSSFWorkbook(stream);//2003版本
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Office 2007"))
                    {
                        throw new Exception("xlsx");
                    }
                    else
                    {
                        throw ex;
                    }
                }
            }
            else
            {
                throw new Exception("Excel文件类型不正确");
            }

            if (sheetName != "")
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null)
                {
                    //如果没有找到指定的sheetName对应的sheet
                    throw new Exception("请按照模板格式导入");
                }
            }
            else
            {
                sheet = null;
                throw new Exception("请按照模板格式导入");
            }

            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                if (isFirstRowColumn)
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = Convert.ToString(cell.StringCellValue);
                            if (cellValue != null)
                            {
                                DataColumn column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                            }
                        }
                    }
                    startRow = sheet.FirstRowNum + 1;
                }
                else
                {
                    startRow = sheet.FirstRowNum;
                }

                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                            if (row.GetCell(j).CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(row.GetCell(j)))//如果类型是日期形式的，转成合理的日期字符串
                            {
                                dataRow[j] = row.GetCell(j).DateCellValue.ToString();
                            }
                        }

                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;
        }
        ///// <summary>
        ///// 实体类转换成DataTable
        ///// </summary>
        ///// <param name="modelList">实体类列表</param>
        ///// <returns></returns>
        //public DataTable FillDataTable(List<T> modelList)
        //{
        //    if (modelList == null || modelList.Count == 0)
        //    {
        //        return null;
        //    }
        //    DataTable dt = CreateData(modelList[0]);
        //    foreach (T model in modelList)
        //    {
        //        DataRow dataRow = dt.NewRow();
        //        foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
        //        {
        //            dataRow[propertyInfo.Name] = propertyInfo.GetValue(model, null);
        //        }
        //        dt.Rows.Add(dataRow);
        //    }
        //    return dt;
        //}

        ///// <summary>
        ///// 根据实体类得到表结构
        ///// </summary>
        ///// <param name="model">实体类</param>
        ///// <returns></returns>
        //private DataTable CreateData(T model)
        //{
        //    DataTable dataTable = new DataTable(typeof(T).Name);
        //    foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
        //    {
        //        dataTable.Columns.Add(new DataColumn(propertyInfo.Name, propertyInfo.PropertyType));
        //    }
        //    return dataTable;
        //}
    }
}
