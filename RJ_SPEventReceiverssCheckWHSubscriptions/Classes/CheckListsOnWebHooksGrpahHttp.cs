namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using AngleSharp.Html.Parser.Tokens;
    using Azure.Core;
    using Azure.Identity;
    using DocumentFormat.OpenXml.Office2010.Excel;
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
            //Get all sites 
            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();
            RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService graphService = new RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService(config);
            GraphServiceClient GClient = await graphService.GetGraphClient(tenantId);
            //Token from Graph.Beta
            //var apptoken = await graphService.GetGraphCLientToken(tenantId, true, GraphService.TokenType.App);
            var apptoken = await graphService.GetGraphCLientToken(tenantId, false, GraphService.TokenType.App);
            //            Console.WriteLine($"App Access Token: {apptoken.Token} -> TokenType = {graphService.tokenType}");

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
                string resourcefilter = "";
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
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine($"List: {list.DisplayName}");
                            Console.ForegroundColor = ConsoleColor.White;
                            resourcefilter = "/drives/" + drive.Id + "/root";

                            try
                            {
                                // Get subscriptions for this list
                                //var subscriptions = await GClient.Sites[site.Id].Lists[list.Id].Subscriptions
                                //Return all subscriptions. 
                                var subsResponse = await GClient.Subscriptions.GetAsync();

                                // Call Graph directly
                                var subscriptions = subsResponse?.Value ?? new List<Subscription>();


                                // Filter by resource
                                var targetSubs = subscriptions.Where(s => string.Equals(s.Resource, resourcefilter, StringComparison.OrdinalIgnoreCase)).ToList();
                                //fIRST REMPOVE SUBSCRIPTIONS ON DRIVES
                                if (targetSubs.Any())
                                {
                                    int subcounter = 0;
                                    Console.ForegroundColor = ConsoleColor.Magenta;
                                    Console.WriteLine($"Webhook DRIVE subscriptions detected: {targetSubs.Count}");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    foreach (var sub in targetSubs)
                                    {
                                        subcounter++;
                                        //  string ID = sub.GetProperty("id").GetString();
                                        // string url = sub.GetProperty("notificationUrl").GetString();
                                        //  DateTime exp = sub.GetProperty("expirationDateTime").GetDateTime();
                                        Console.WriteLine($"Id: {sub.Id}");
                                        Console.WriteLine($"Changetype: {sub.ChangeType}");
                                        Console.WriteLine($"NotificationUrl:{sub.NotificationUrl}");
                                        Console.WriteLine($"Expires: {sub.ExpirationDateTime}");
                                        //Check notificationURL and expiration. If expires today then re-register 
                                        var remainingTimespan = (sub.ExpirationDateTime?.Subtract(DateTime.Now)).GetValueOrDefault().TotalMinutes;
                                        //Just remove
                                        //if (remainingTimespan <= 0 && sub.NotificationUrl == notificationUrl)
                                        //Expired -> Remove
                                        try
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Subscription expired -> REMOVE!!! {sub.Id}");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            await GClient.Subscriptions[sub.Id].DeleteAsync();
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Subscription could not be removed, {ex.Message}");
                                        }
                                        //  reregister = true;
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.WriteLine($"(Subscription #{subcounter}, id {sub.Id}) remaining webhook timespan on {list.Name} in minutes  : {remainingTimespan} : reregister -> {reregister}");
                                        Console.ForegroundColor = ConsoleColor.White;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("No DRIVE webhook subscriptions detected -> ReRegister");
                                    reregister = true;
                                }
                                //Then check subscriptions on lists
                                // Get subscriptions for this list
                                //var subscriptions = await GClient.Sites[site.Id].Lists[list.Id].Subscriptions
                                //endpoint = $"https://graph.microsoft.com/v1.0/drives/{drive.Id}/root/subscriptions";
                                endpoint = $"https://graph.microsoft.com/v1.0/sites/{site.Id}/lists/{list.Id}/subscriptions";

                                // Call Graph directly
                                var resp = await HttpClient.GetAsync(endpoint);
                                var json = await resp.Content.ReadAsStringAsync();
                                using var doc = JsonDocument.Parse(json);
                                if (doc.RootElement.TryGetProperty("value", out var subs) && subs.GetArrayLength() > 0)
                                {
                                    Console.WriteLine($"Webhook LIST subscriptions detected: {subs.GetArrayLength()}");
                                    foreach (var sub in subs.EnumerateArray())
                                    {
                                        string ID = sub.GetProperty("id").GetString();
                                        string url = sub.GetProperty("notificationUrl").GetString();
                                        DateTime exp = sub.GetProperty("expirationDateTime").GetDateTime();
                                        Console.WriteLine($"Id: {ID}");
                                        Console.WriteLine($"NotificationUrl:{url}");
                                        Console.WriteLine($"Expires: {exp}");
                                        //Check notificationURL and expiration. If expires today then re-register
                                        var remainingtimespan = (exp - DateTime.Now).TotalMinutes;
                                        if (remainingtimespan <= 300000)
                                        {
                                            //reregister expired subscriptions
                                            try
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine($"Subscription id {ID} VOID, remove!!!");
                                                await GClient
                                                    .Sites[site.Id]
                                                    .Lists[list.Id]
                                                    .Subscriptions[ID]
                                                    .DeleteAsync();
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine($"Exception removng subscription {ex.Message}");
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                            reregister = true;
                                        }
                                        Console.WriteLine($"remaining timespan in minutes  : {remainingtimespan} : reregister -> {reregister}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("No webhook subscriptions on list detected");
                                    reregister = true;
                                }
                                if (reregister)
                                {
                                    RegisterWebhook RegWH = new RegisterWebhook();
                                    await RegWH.RegisterWebhookAsync(args, list.Id, site.WebUrl, notificationUrl, expirationMinutes, tenantId, clientSecret, clientId, false, false);
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
            Console.WriteLine("Webhook registrations completed");
        }
    }
}
