namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using AngleSharp.Html.Parser.Tokens;
    using Azure.Core;
    using Azure.Identity;
    using DocumentFormat.OpenXml.Spreadsheet;
    using Microsoft.AspNetCore.SignalR;
    using Microsoft.Graph;
    using Microsoft.Graph.Models;
    //using Microsoft.Graph.Beta;
    //using Microsoft.Graph.Beta.Models.Networkaccess;
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Client.Platforms.Features.DesktopOs.Kerberos;
    using Microsoft.SharePoint.Client.WebParts;
    //using Microsoft.SharePoint.Client;
    //using Microsoft.SharePoint.Internal;
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Runtime.CompilerServices;
    using System.Text.Json;
    using System.Threading.Tasks;
    using static System.Net.WebRequestMethods;

    public class CheckListsOnWebHooksHTTPS
    {
        public async Task CheckAllSiteLists
            (
            string[] args, 
            string notificationUrl, 
            int expirationMinutes,  
            string tenantId,
            string clientSecret,
            string clientId,
            bool? onDocLibOnly = false
            )
        {

            //Future improvement :use more generic solution and Graphservice class
          //  var Domain = "lls6.sharepoint.com";
          //  var Resource = "https://" + Domain;
          //  var SiteName = "/sites/SP-EventReceivers-Test";
          //  var sitedef = Domain + ":" + SiteName;
          //  string SiteUrl = Resource + SiteName;

            // App registration details
           // var clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
     
           // var clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";

            // First page request: GET /sites?search=.
            //var sitesResponse = await GClient.Sites.Search(".").GetAsync(); // broad search to match all site collections
            //var HttpClient = graphService.GetHttpClient(tenantId, true);
            //HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //string notificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev/api/WebHookListener"; // Must be HTTPS and reachable

            //int expirationDays = 30; // Max 43200 minutes (30 days) for list webhooks

            //var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            //var scopes = new[] { "https://graph.microsoft.com/.default" };
            //Get all sites 
            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();
            RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService graphService = new RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService(config);
            GraphServiceClient GClient = await graphService.GetGraphClient(tenantId);
            //Token from Graph.Beta
            //var apptoken = await graphService.GetGraphCLientToken(tenantId, true, GraphService.TokenType.App);
            var apptoken = await graphService.GetGraphCLientToken(tenantId, false, GraphService.TokenType.App);
            Console.WriteLine($"App Access Token: {apptoken.Token} -> TokenType = {graphService.tokenType}");

            HttpClient HttpClient = new HttpClient();
            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apptoken.Token);
            HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));



            //var subscriptions = await GClient.Subscriptions.GetAsync();


            var sites = await GClient.Sites
                .GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Top = 999;    // max per page
                });

            string endpoint = "";
            foreach (var site in sites.Value)
            {
                //For debug only check one site
                if (site.WebUrl == "https://lls6.sharepoint.com/sites/SP-EventReceivers-Test")
                {
                    Console.WriteLine($"{site.DisplayName} -> {site.WebUrl}");
                    //Get site lists 
                    var lists = await GClient.Sites[site.Id].Lists.GetAsync();
                    foreach (var list in lists.Value)
                    {
                        Boolean reregister = false;
                        var drive = await GClient.Sites[site.Id].Lists[list.Id].Drive.GetAsync();
                        if (drive.DriveType == "documentLibrary")
                        {
                            Console.WriteLine($"  List: {list.DisplayName}");
                            try
                            {
                                // Get subscriptions for this list
                                //var subscriptions = await GClient.Sites[site.Id].Lists[list.Id].Subscriptions
                                endpoint = $"https://graph.microsoft.com/v1.0/drives/{drive.Id}/root/subscriptions";
                                //endpoint = $"https://graph.microsoft.com/v1.0/sites/{site.Id}/lists/{list.Id}/subscriptions";

                                // Call Graph directly
                                var resp = await HttpClient.GetAsync(endpoint);
                                var json = await resp.Content.ReadAsStringAsync();
                                using var doc = JsonDocument.Parse(json);
                                if (doc.RootElement.TryGetProperty("value", out var subs) && subs.GetArrayLength() > 0)
                                {
                                    Console.WriteLine($"    Webhook subscriptions detected: {subs.GetArrayLength()}");
                                    foreach (var sub in subs.EnumerateArray())
                                    {
                                        string ID = sub.GetProperty("id").GetString();
                                        string url = sub.GetProperty("notificationUrl").GetString();
                                        DateTime exp = sub.GetProperty("expirationDateTime").GetDateTime();
                                        Console.WriteLine($"      Id: {ID}");
                                        Console.WriteLine($"      NotificationUrl:{url}");
                                        Console.WriteLine($"      Expires: {exp}");
                                        //Check notificationURL and expiration. If expires today then re-register 
                                        var remainingtimespan = (exp - DateTime.Now).TotalMinutes;
                                        if (remainingtimespan <= 300000 && notificationUrl == url)
                                        {
                                            reregister = true;
                                        }
                                        Console.WriteLine($"remaining timespan in minutes  : {remainingtimespan} : reregister -> {reregister}");

                                    }
                                }
                                else
                                {
                                    Console.WriteLine("    No webhook subscriptions, to be added.");
                                    reregister = true;
                                }
                                if(reregister)
                                {
                                    RegisterWebhook RegWH = new RegisterWebhook();
                                    await RegWH.RegisterWebhookAsync(args, list.Id, site.WebUrl, notificationUrl, expirationMinutes, tenantId, clientSecret, clientId, false, true);
                                }
                            }
                            catch (ServiceException ex)
                            {
                                Console.WriteLine($"    Error retrieving subscriptions: {ex.Message}");
                            }
                        }
                    }
                }
            }
            Thread.Sleep(30000);
        }
    }
}
