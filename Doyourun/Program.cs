using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Reflection;
using System.Data;
using System.Data.OleDb;
using System.Text;
using ExcelDataReader;
using NPOI.HSSF.UserModel;
using WindowsFormsExe����;

namespace Doyourun

{
	public class Program
	{
		public static void Main(string[] args)
		{
		CreateHostBuilder(args).Build();
			CreateHostBuilder(args).Build();
			IServiceCollection s = new ServiceCollection();
			//s.AddScoped<ITestOne, TestOne>()
		 //  .AddScoped<ITestTwo, TestTwo>()
		 //  .AddScoped<ITestThree, TestThree>()
		 //  .AddScoped<ITestApp, TestApp>()
		 //  .BuildServiceProvider()
		 //  .GetService<ITestApp>();

			string tr = @"
               <table border='1' width ='100%' style'color : red'>
			<tr height='20' align='center' style='font-weight:bold'> <td>��Ŀ</td> <td>���</td> <td>����</td></tr> 

			<tr height='20' align='center' > <td>�꿧</td> <td>44</td> <td>204</td></tr> 
			<tr height='20' align='center' > <td>����û</td> <td>33</td> <td>24</td></tr> 
			<tr height='20' align='center'> <td>��������Ĥ</td> <td>55</td> <td>4234</td></tr> 
			</table> 
			<table border='2' width ='100%' style'color : green'>
			<tr height='20' align='center' style='font-weight:bold'> <td>��Ŀ2</td> <td>���2</td> <td>����2</td></tr> 
			<tr height='20' align='center' > <td>�꿧</td> <td>44</td> <td>204</td></tr> 
			<tr height='20' align='center' > <td>����û</td> <td>33</td> <td>24</td></tr> 
			<tr height='20' align='center' style'color : red' >  <td>��������Ĥ</td> <td>55</td> <td>4234</td></tr> 
			</table>";
		

			var path = "E:/pp.xls";
			List<DataTable> stulist = new List<DataTable>();
			List<student> stulist1 = new List<student>();
			stulist1.Add(new student() { ID = "����", Name = "����" });
			
			
			DataTable dt1 = new DataTable();
			dt1.TableName = "����1";
			dt1.Columns.Add("����");
			dt1.Columns.Add("����");
		

			List<student> stulist2 = new List<student>();
			stulist2.Add(new student() { ID = "Ψһ", Name = "�º�" });
			DataTable dt2 = new DataTable();
			dt2.TableName = "����2";
			dt2.Columns.Add("Ψһ2");
			dt2.Columns.Add("�º�2");
			stulist.Add(dt1);
			stulist.Add(dt2);
			//SearchSheetToDT("","P",path);

				//var table=	NPOHelp.ExcelToDataTable(path);
			

			FileStream fs = new FileStream(path ,FileMode.Open, FileAccess.Read, FileShare.None);
			byte[] buffer = new byte[fs.Length];

			fs.Read(buffer, 0,buffer.Length);
			fs.Close();
			var  str = Encoding.Default.GetString(buffer);
			Stream strs = new MemoryStream(buffer);
			
			//DataTable dt = Helper.ConvertStreamToDataTable(strs,"xls","P",true);


			//using (var stream =new  FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
			//{
			//	//var workbook = new HSSFWorkbook(stream);
			//	var sheet = workbook.GetSheetAt(1);
			//	//var sheet2 = workbook.GetSheetAt(2);

			//	if (sheet.SheetName == "����1")
			//	{
			//		var row = sheet.GetRow(0);
			//		var cell = row.CreateCell(0);
			//		Console.WriteLine(cell.ToString());
			//	}
			//	//var c = reader.AsDaset();
			    
			//}


				_ = sho("C:/Users/������/Desktop/p.xlsx", tr);
			Console.ReadLine();
		}
		public static async Task<string> sho(string naem, string content)
		{

			File.Delete(naem);
		
			_ = File.WriteAllTextAsync(naem, content);
			string s = await File.ReadAllTextAsync(naem);
			 
			return s;
		}

		public static IHostBuilder CreateHostBuilder(string[] args) =>
			Host.CreateDefaultBuilder(args)
				.ConfigureWebHostDefaults(webBuilder =>
				{
					webBuilder.UseStartup<Startup>();
				});

		public static DataTable SearchSheetToDT(string strSearch, string sheetName,string path)
		{
			
			//���ӱ��ַ���
			string ExcelConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + @path + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2;ImportMixedTypes=Text'";
			using (OleDbConnection ole_conn = new OleDbConnection(ExcelConnection))
			{
				ole_conn.Open();
				using (OleDbCommand ole_cmd = ole_conn.CreateCommand())
				{
					ole_cmd.CommandText = strSearch;
					OleDbDataAdapter adapter = new OleDbDataAdapter(ole_cmd);
					DataSet ds = new DataSet();
					adapter.Fill(ds, sheetName);//sheetName����excel���sheet����
					DataTable dt = new DataTable();

					for (int i = 0; i < ds.Tables.Count; i++)
					{

						dt = ds.Tables[i];
					
					}

					return dt;
				}
			}
		}

		public interface ITestOne { }
		public interface ITestTwo { }
		public interface ITestThree { }

		public class TestOne : ITestOne { }
		public class TestTwo : ITestTwo { }
		public class TestThree : ITestThree { }

		public interface ITestApp { }
		public class TestApp : ITestApp
		{
			public TestApp(ITestOne testOne, ITestTwo testTwo, ITestThree testThree)
			{
				Console.WriteLine($"TestApp({testOne}, {testTwo}, {testThree})");
			}
		}
	}
}
