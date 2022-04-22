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
			///注册接口和实现类映射关系
			services.AddScoped<Program,Program>();
		}

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
		{
			if (env.IsDevelopment())
			{
				app.UseDeveloperExceptionPage();
			}
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//国标 2312
			app.UseRouting();
			app.UseCookiePolicy(); //注入cookie
		
			app.Run(async context => { 
				context.Response.ContentType = "text/html;charset=utf-8"; 
				await context.Response.WriteAsync("视频空调");
			});

			//_ = request.Cookies.TryGetValue("key", out value);
			app.UseEndpoints(endpoints =>
			{
				endpoints.MapGet("/", async context =>
				{
					await context.Response.WriteAsync("Hello World!");

					string filname = "C:/Users/索鹏凯/Desktop/p.txt";
					File.Delete(filname);

					var filnames = "C:\\Users\\索鹏凯\\Desktop\\pppp.txt";
					_ = File.WriteAllTextAsync(filnames, "所彭凯插入。。。《》《》");
					string s = await File.ReadAllTextAsync(filnames,System.Text.Encoding.Default);
		
					await context.Response.WriteAsync(s+"圣诞节空间");
				});

			});
			
		}
		
	}
}
