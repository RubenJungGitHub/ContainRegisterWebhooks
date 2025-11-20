namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using Azure.Identity;
    using Microsoft.Graph;
    //using Microsoft.Graph.Auth;
    using Microsoft.Graph.Models;
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Client.Platforms.Features.DesktopOs.Kerberos;
    using System;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using static System.Net.WebRequestMethods;


    public class RegisterWebhook
    {
        public async Task RegisterWebhookAsync
                        (
            string[] args,
            string listId,

            string SiteUrl,
            string notificationUrl,
            int expirationMinutes,
            string tenantId,
            string clientSecret,
            string clientId,
            bool? onListCreation = false,
            bool? onDocLibOnly = false
            )
        {
            //Future improvement :use more generic solution and Graphservice class
            var Domain = "lls6.sharepoint.com";
            // var Resource = "https://" + Domain;
            var SiteName = "/sites/SP-EventReceivers-Test";
            var sitedef = Domain + ":" + SiteName;
            //string SiteUrl = Resource + SiteName;

            // App registration details
            //  var clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
            //  var tenantId = "9a1b5f77-1f1a-40ac-b1a1-38617300f02a";
            //  var clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";

            //string ListTitle = "TestWebhook";
            //string ListTitle = "Documents";
            //string notificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev"; // Must be HTTPS and reachable

            //int expirationDays = 30; // Max 43200 minutes (30 days) for list webhooks

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService graphService = new RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService(config);
            GraphServiceClient GClient = await graphService.GetGraphClient(tenantId);
            var apptoken = await graphService.GetGraphCLientToken(tenantId, true, GraphService.TokenType.App);
            //  Console.WriteLine($"App Access Token: {apptoken.Token} -> TokenType = {graphService.tokenType}"); ;
            HttpClient HttpClient = new HttpClient();
            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apptoken.Token);
            HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));



            var graphClient = new GraphServiceClient(credential, scopes);
            var site = await graphClient
                .Sites[sitedef]
                .GetAsync();

            //Get subscriptions
            var subscriptions = await graphClient.Subscriptions
            .GetAsync();


            string[] parts = site.Id.Split(',');

            string hostname = parts[0];
            string siteCollectionId = parts[1];  // site collection GUIforsD
            string siteId = parts[2];           // actual site GUID
            var list = await graphClient.Sites[site.Id].Lists[listId].GetAsync();

            //To do  validation if null
            //var listId = list.Id;
            var drive = await graphClient.Sites[site.Id].Lists[listId].Drive.GetAsync();

            //Doublecheck
            if (drive.DriveType == "documentLibrary")
            {
                var endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{drive.Id}/root/subscriptions";
                // var subscriptionsRegistered = await graphClient.Subscriptions.GetAsync();
                var response = await HttpClient.GetAsync(endpoint);
                var content = await response.Content.ReadAsStringAsync();
                //ALL SUBSCRIPTIONS
                var subscriptionsRegistered = await graphClient.Subscriptions.GetAsync();
                if (subscriptionsRegistered.Value != null)
                {
                    // Iterate over each subscription
                    foreach (var registerredsubscription in subscriptionsRegistered.Value)
                    {
                        // Do something with each subscription
                        //Debug Remove Webhook subscription
                        try
                        {
                            int a = 1;
                            // await graphClient.Subscriptions[registerredsubscription.Id].DeleteAsync();
                        }
                        catch { }
                    }
                }
                else
                {
                    Console.WriteLine("No subscriptions found.");
                }

                //Dor documentlibs driveid is to be used!!!
                Subscription subscription;
                // Register webhook on specific list 
                if (!Convert.ToBoolean(onListCreation))
                {
                    if (Convert.ToBoolean(onDocLibOnly))
                    {
                        subscription = new Subscription
                        {
                            ChangeType = "updated",                      // or "created,updated,deleted"
                            Resource = $"/drives/{drive.Id}/root",       // correct resource path
                            NotificationUrl = notificationUrl,
                            ExpirationDateTime = DateTimeOffset.UtcNow.AddMinutes(expirationMinutes),
                            ClientState = clientSecret
                        };
                    }
                    else
                    {
                        /*
                         if (Convert.ToBoolean(onDocLibOnly) && list.DisplayName != "Documents")
                         {   // For document libraries, use the drive resource
                             subscription = new Subscription
                             {
                                 ChangeType = "updated",
                                 Resource = $"/sites/{site.Id}/drives/{drive.Id}/root",
                                 NotificationUrl = notificationUrl,
                                 ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(expirationMinutes),
                                 ClientState = clientSecret
                             };
                         }
                         else
                         {
                        */
                        subscription = new Subscription
                        {
                            ChangeType = "updated",
                            Resource = $"/sites/{site.Id}/lists/{listId}",
                            NotificationUrl = notificationUrl,
                            ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(expirationMinutes),
                            ClientState = clientSecret
                        };
                    }
                }
                else
                {
                    //ChatGptResponse
                    /*Bottom line . Not working 
                    Microsoft Graph webhooks cannot trigger on list creation.
                    ChangeType = "created" only works on list items or certain other resources.
                    For new lists, you must use polling, audit logs, or Power Automate.
                    */
                    subscription = new Subscription
                    {
                        ChangeType = "created,updated,deleted",
                        NotificationUrl = notificationUrl,
                        Resource = "/sites/{site-id}/lists",
                        ExpirationDateTime = DateTimeOffset.UtcNow.AddMinutes(expirationMinutes),
                        ClientState = clientSecret
                    };
                }

                try
                {
                    var createdSubscription = await graphClient.Subscriptions.PostAsync(subscription);
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Webhook type  {subscription.ChangeType} on {subscription.Resource} registered on list {list.DisplayName} ! Subscription ID: {createdSubscription.Id}");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
                {
                    Console.WriteLine($"Error creating subscription: {ex.Error.Message}");
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine($"Error registering webhook: {ex.Message}");
                }
            }
        }
    }
}

