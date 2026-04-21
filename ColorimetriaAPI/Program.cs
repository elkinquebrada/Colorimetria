using System;
using System.Web.Http;
using Microsoft.Owin.Hosting;
using Owin;
using Unity;
using Unity.AspNet.WebApi;
using ColorimetriaAPI.Services;

namespace ColorimetriaAPI
{
    public class Program
    {
        static void Main(string[] args)
        {
            string baseAddress = "http://localhost:5000/";
            try
            {
                using (WebApp.Start<Startup>(url: baseAddress))
                {
                    Console.WriteLine("API Colorimetría levantada en: " + baseAddress);
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();
            }
        }
    }

    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            HttpConfiguration config = new HttpConfiguration();
            var container = new UnityContainer();
            container.RegisterType<ColorimetryService>();
            config.DependencyResolver = new UnityHierarchicalDependencyResolver(container);
            config.MapHttpAttributeRoutes();
            config.Routes.MapHttpRoute("DefaultApi", "api/{controller}/{id}", new { id = RouteParameter.Optional });
            app.UseWebApi(config);
        }
    }
}