using Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;

namespace WebApplication3.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class FileController : ControllerBase
	{
		[HttpPost("UploadFile")]
		public void UploadFile()
		{

			var abc = Request.Form["dev"];
			IFormFile ff = Request.Form.Files["hh"];
			IFormFile file = Request.Form.Files[0];
			string type = "";
			if (ff.FileName.Contains(".xls"))
			{
				type = "xls";   
			}
			Stream stream = ff.OpenReadStream();

		
			byte[] buffer = new byte[file.Length];


			var str = Encoding.Default.GetString(buffer);
			Stream strs = new MemoryStream(buffer);
			var data2 = Help.ConvertStreamToDataTable(stream, type, "B", true);
			data2.TableName = "数据学校";
			Stream streams = ff.OpenReadStream();

			var data =Help.ConvertStreamToDataTable(streams, type,"P",true);
			data.TableName = "数据大厅";


			Stream stream3 = ff.OpenReadStream();
	        var data3 = Help.ConvertStreamToDataTable(stream3, type, "sheet3", true);
		

			DataSet ds = new DataSet();
			ds.Tables.Add(data);
			ds.Tables.Add(data2);
			ds.Tables.Add(data3);

			string filePath = "D:/ 两个sheet.tex";
			List<DataTable> dtlist = new List<DataTable>();
			for (int i = 0; i < ds.Tables.Count; i++)
			{
				dtlist.Add(ds.Tables[i]);
			}
		

			var c = Help.GridToExcels(dtlist, "filePath", 1);


			for (int i = 0; i < ds.Tables.Count; i++)
			{
				DataTable dt = new DataTable();
				dt = ds.Tables[i];
				int columcount = dt.Columns.Count;
				
				Application myexcel = new Application();
				Workbook myworkbook = myexcel.Application.Workbooks.Add(true);
				Excel.Worksheet myworksheet = (Excel.Worksheet)myexcel.Worksheets[dt.TableName];
				myworksheet.Name = dt.TableName;
				myworksheet.Rows.Font.Size = 33;
				myworksheet.Rows.Font.Name = "宋体";
				myworksheet.Rows.VerticalAlignment = XlHAlign.xlHAlignCenter;
				myworksheet.Rows.WrapText = true;

				
				//myworksheet.Copy(Type.Missing, myworkbook.Sheets[3]);

				
			}
			FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
			fs.Close();
			//File.Deltet(filePath);
			

			


		}
	}
}
