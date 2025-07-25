using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Provisioning.Model;
using SharePoint.CSOM.Extensions;
using System.Net;
using TestApp.Temp;
using AuthenticationManager = PnP.Framework.AuthenticationManager;

namespace TestApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            string siteUrl = "";
            string clientId = "";
            string tenantId = "";

            string certPath = @"";
            string certPassword = "!";

            string listId = "";
            string siteId = "";

            var authManager = new AuthenticationManager(clientId, certPath, certPassword, tenantId);

            using var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Trace);
            });

            var logger = loggerFactory.CreateLogger<Program>();

            // Configure global execution logic - now ExecuteScopeAsync gets the actual CSOM setup actions!
            //SharePoint.CSOM.Extensions.Configuration.CSOMConfiguration.Configure(async (context, setupAction) =>
            //{
            //    if (context is ClientContext clientContext)
            //    {
            //    }
            //    else
            //    {
            //        // Fallback for non-ClientContext scenarios
            //        setupAction();
            //        await context.ExecuteQueryAsync();
            //    }
            //});

            using (var context = authManager.GetContext(siteUrl))
            {
            }
        }
    }
}