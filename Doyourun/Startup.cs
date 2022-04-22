using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Doyourun
{
	public class Startup
	{
		// This method gets called by the runtime. Use this method to add services to the container.
		// For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
		public void ConfigureServices(IServiceCollection services)
		{
			services.AddTransient<Program>();
			//services.AddMvc(mvc =>
			//{
			//	mvc.Filters.Add(typeof(Program));
			//});
			///ע��ӿں�ʵ����ӳ���ϵ
			services.AddScoped<Program,Program>();
		}

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
		{
			if (env.IsDevelopment())
			{
				app.UseDeveloperExceptionPage();
			}
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//���� 2312
			app.UseRouting();
			app.UseCookiePolicy(); //ע��cookie
		
			app.Run(async context => { 
				context.Response.ContentType = "text/html;charset=utf-8"; 
				await context.Response.WriteAsync("��Ƶ�յ�");
			});

			//_ = request.Cookies.TryGetValue("key", out value);
			app.UseEndpoints(endpoints =>
			{
				endpoints.MapGet("/", async context =>
				{
					await context.Response.WriteAsync("Hello World!");

					string filname = "C:/Users/������/Desktop/p.txt";
					File.Delete(filname);

					var filnames = "C:\\Users\\������\\Desktop\\pppp.txt";
					_ = File.WriteAllTextAsync(filnames, "�������롣������������");
					string s = await File.ReadAllTextAsync(filnames,System.Text.Encoding.Default);
		
					await context.Response.WriteAsync(s+"ʥ���ڿռ�");
				});

			});
			
		}
		
	}
}
