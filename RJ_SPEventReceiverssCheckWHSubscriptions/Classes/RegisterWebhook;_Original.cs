namespace RJ_SPEventReceiversWebhookSubscribeOriginal.Classes
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

    public class RegisterWebhookOriginial
    {
        public async Task RegisterWebhookAsync(string[] args, string notificationUrl)
        {
            var Domain = "lls6.sharepoint.com";
            var Resource = "https://" + Domain;
            var SiteName = "/sites/SP-EventReceivers-Test";
            var sitedef = Domain + ":" +   SiteName;
            string SiteUrl = Resource + SiteName;

            // App registration details
            var clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
            var tenantId = "9a1b5f77-1f1a-40ac-b1a1-38617300f02a";
            var clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";

            string ListTitle = "Documents";

            //string notificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev/api/WebHookListener"; // Must be HTTPS and reachable
            //string notificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev"; // Must be HTTPS and reachable
            //string notificationUrl = "https://webhook.site/da6bf17b-bf55-4d80-add5-5b4a72c99265";
            //string notificationUrl = "https://contain-webhook-handler.free.beeceptor.com";




            int expirationDays = 30; // Max 43200 minutes (30 days) for list webhooks

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            //var graphClient = new GraphServiceClient(credential);
            //var site = await graphClient.Sites["root"].GetAsync();
            //Console.WriteLine($"Site retrieved: {site.Name}");


            var graphClient = new GraphServiceClient(credential, scopes);
            var site = await graphClient
                .Sites[sitedef]
                .GetAsync();

            //Get subscriptions
            var subscriptions = await graphClient.Subscriptions
            .GetAsync();


            string[] parts = site.Id.Split(',');

            string hostname = parts[0];           // contoso.sharepoint.com
            string siteCollectionId = parts[1];  // site collection GUIforsD
            string siteId = parts[2];           // actual site GUID
            var list = await graphClient.Sites[site.Id].Lists[ListTitle].GetAsync();

            //To do  validation if null
            var listId = list.Id;

            var subscriptionsRegistered = await graphClient.Subscriptions.GetAsync();
            if (subscriptionsRegistered != null)
            {
                // Iterate over each subscription
                foreach (var registerredsubscription in subscriptionsRegistered.Value)
                {
                    // Do something with each subscription
                    //Debug Remove Webhook subscription
                    try
                    {
                        await graphClient.Subscriptions[registerredsubscription.Id].DeleteAsync();
                    }
                    catch { }
                }
            }
            else
            {
                Console.WriteLine("No subscriptions found.");
            }



            // Register webhook
            var subscription = new Subscription
            {
                ChangeType = "updated",
                Resource = $"/sites/{site.Id}/lists/{listId}",
                NotificationUrl = notificationUrl,
                ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(expirationDays),
                ClientState = clientSecret
            };

            try
            {
                //========================================================================================
                //WORKING FOR ONE SUBSCRIPTION DON;T TOUCH!!!!
                var createdSubscription = await graphClient.Subscriptions.PostAsync(subscription);
                Console.WriteLine($"Webhook registered! Subscription ID: {createdSubscription.Id}");
                //========================================================================================
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                Console.WriteLine($"Error creating subscription: {ex.Error.Message}");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error registering webhook: {ex.Message}");
            }


            //validate subscription 
            var subs = await graphClient.Subscriptions.GetAsync();
            foreach (var s in subs.Value)
            {
                Console.WriteLine($"{s.Id} {s.Resource
                    } - {s.ExpirationDateTime} - {s.ClientState}");

                await graphClient.Subscriptions[s.Id].DeleteAsync();
            }
        }
    }
}
