using Excel;
using Microsoft.AspNetCore.Http;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;

namespace WebApplication3
{
	public   class Help
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


		/// <summary>
		/// 将多个DataTable数据导入到excel中
		/// </summary>
		/// <param name="dts">要导入的数据集合</param>
		/// <param name="strExcelFileName">定义Excel文件名</param>
		/// <param name="indexType">给个 1 就行</param>
		
		public static bool GridToExcels(List<DataTable> dts, string strExcelFileName, int indexType)
        {
            bool BSave = false;
            try
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                DataSet set = new DataSet();
                foreach (DataTable dt in dts)
                {
                    ISheet sheet = workbook.CreateSheet(dt.TableName);
                    ICellStyle HeadercellStyle = workbook.CreateCellStyle();
                    HeadercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    
                    //字体
                    NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
                    headerfont.Boldweight = (short)FontBoldWeight.Bold;
                    HeadercellStyle.SetFont(headerfont);


                    //用column name 作为列名
                    int icolIndex = 0;
                    IRow headerRow = sheet.CreateRow(0);
                    foreach (DataColumn item in dt.Columns)
                    {
                        ICell cell = headerRow.CreateCell(icolIndex);
                        cell.SetCellValue(item.ColumnName);
                        cell.CellStyle = HeadercellStyle;
                        icolIndex++;
                    }


                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    //为避免日期格式被Excel自动替换，所以设定 format 为 『@』 表示一率当成text来看
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
                    cellfont.Boldweight = (short)FontBoldWeight.Normal;
                    cellStyle.SetFont(cellfont);
                    if (indexType == 1)
                    {
                        //建立内容行
                        int iRowIndex = 1;
                        int iCellIndex = 0;
                        foreach (DataRow Rowitem in dt.Rows)
                        {
                            IRow DataRow = sheet.CreateRow(iRowIndex);
                            foreach (DataColumn Colitem in dt.Columns)
                            {


                                ICell cell = DataRow.CreateCell(iCellIndex);
                                cell.SetCellValue(Rowitem[Colitem].ToString());
                                cell.CellStyle.FillBackgroundColor =16;
                                cell.CellStyle = cellStyle;
                                iCellIndex++;
                            }
                            iCellIndex = 0;
                            iRowIndex++;
                        }
                        //自适应列宽
                        for (int i = 0; i < icolIndex; i++)
                        {
                            sheet.AutoSizeColumn(i);
                        }
                    }

                }
                FileStream file = new FileStream(strExcelFileName, FileMode.OpenOrCreate);
                workbook.Write(file);
                file.Flush();
                file.Close();
                BSave = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return BSave;
        }

    }
}
