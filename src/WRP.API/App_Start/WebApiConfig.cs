using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace WRP.API
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{action}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            // Remove the XML formatter (only JSON)
            config.Formatters.Remove(config.Formatters.XmlFormatter);

        }
    }
}
