using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using IronPdf; // Add this using directive

namespace ConvertToPdfAndExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Initialize IronPDF License Key
            IronPdf.License.LicenseKey = "IRONSUITE.VAHISHIABBAS.GMAIL.COM.9914-1A6F44D8E3-BJNPBKTNWT2CROJ2-NGFRRSCPY6P2-H544YZA2VNEH-SQQJJKOZZ42J-ZPVJV7US736Z-FTNM7JYW2UGL-6EIMDC-TFBSMRGVFWKNUA-DEPLOYMENT.TRIAL-EXNOIF.TRIAL.EXPIRES.07.NOV.2024";

            // Add services to the container.
            builder.Services.AddRazorPages();

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (!app.Environment.IsDevelopment())
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            app.MapRazorPages();

            app.Run();
        }
    }
}
