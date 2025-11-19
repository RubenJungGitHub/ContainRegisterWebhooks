namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using Azure.Identity;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using System;
    using System.Threading.Tasks;

    public class WebhookService
    {
        private readonly GraphServiceClient _graphClient;
        public WebhookService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<Microsoft.Graph.Models.Subscription> RegisterWebhook(string driveId, string notificationUrl, int expirationDays = 30)
        {
            //Future improvement :use more generic solution and Graphservice class
            var Domain = "lls6.sharepoint.com";
            var Resource = "https://" + Domain;
            var SiteName = "/sites/SP-EventReceivers-Test";
            var sitedef = Domain + ":" + SiteName;
            string SiteUrl = Resource + SiteName;

            // App registration details
            var clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
            var tenantId = "9a1b5f77-1f1a-40ac-b1a1-38617300f02a";
            var clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";

            string ListTitle = "Documents";

            string NotificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev"; // Must be HTTPS and reachable

            int ExpirationDays = 30; // Max 43200 minutes (30 days) for list webhooks

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var scopes = new[] { "https://graph.microsoft.com/.default" };

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
            var list = await graphClient.Sites[site.Id].Lists[ListTitle].GetAsync();

            //To do  validation if null
            var listId = list.Id;
            var subscription = new Microsoft.Graph.Models.Subscription
            {
                ChangeType = "updated",
                Resource = $"/drives/{driveId}/root",
                NotificationUrl = NotificationUrl,
                ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(ExpirationDays),
                ClientState = Guid.NewGuid().ToString()
            };

            try
            {
                var createdSubscription = await _graphClient.Subscriptions.PostAsync(subscription);
                Console.WriteLine($"Webhook registered! Subscription ID: {createdSubscription.Id}");
                return createdSubscription;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating subscription: {ex.Message}");
                throw;
            }
        }
    }
}
