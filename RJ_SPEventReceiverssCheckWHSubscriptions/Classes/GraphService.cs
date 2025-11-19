namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using Azure;
    using Azure.Core;
    using Azure.Identity;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Bibliography;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.EMMA;
    using DocumentFormat.OpenXml.Math;
    using DocumentFormat.OpenXml.Office.CustomUI;
    using DocumentFormat.OpenXml.Office2010.Excel;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Microsoft.AspNetCore.Http.HttpResults;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Graph;
    using Microsoft.Graph.Auth;
    using Microsoft.Graph.Beta;
    using Microsoft.Graph.Beta.Models;
    using Microsoft.Graph.Models;
    using Microsoft.Graph.Models.TermStore;
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Client.Platforms.Features.DesktopOs.Kerberos;
    using Microsoft.Kiota.Abstractions.Serialization;
    using Microsoft.SharePoint.Client;
    using Microsoft.SharePoint.Client.Discovery;
    using Microsoft.SharePoint.Client.Taxonomy;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
//    using PnP.Core;
 //   using PnP.Core.Auth;
 //   using PnP.Core.Services;
  //  using PnP.Core.Services;
 //   using PnP.Framework;
    using System;
    using System.ClientModel.Primitives;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;
    using System.Text.Json;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using Xceed.Words.NET;
//    using static RJ_SPEventReceiversASPWebApp.Classes.GraphService;
    using static RJ_SPEventReceiversWebhookSubscribe.Classes.GraphService;
    using static System.Net.WebRequestMethods;

    public struct Parameters
    {
        // Hardcodes App registration details for POC. Should be stored in secure config file in JSON format. 
        public string clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
        public string tenantId; //= "9a1b5f77-1f1a-40ac-b1a1-38617300f02a";
        public string clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";
        public string Scopes;
        public TokenType tokenType = TokenType.App;
        //public string DelegatedUserScopes = "{ \"Sites.Read.All\" };";
        public Parameters(string tenantid, TokenType? tokenType = null)
        {
            tenantId = tenantid;
            if (tokenType == TokenType.App || tokenType == null)
            {
                Scopes = "https://graph.microsoft.com/.default";
            }
            else
            {
                string[] scopes = new string[] { "Sites.Read.All", "Sites.ReadWrite.All" };
                this.tokenType = TokenType.DelegatedUser;
            }

        }
    }

    public class GraphService
    {
        //Hardcoded for POC purposes but should be derrived from ChangedItems!!
        private string _FullSiteId;
        public string _Domain;
        private string _SiteID;
        private string _SubSiteId;
        private string _SiteURefForPermissions;
        public string _SiteUrl;
        public string _SiteUrlWithTenantID;
        private string _SiteRelativeUrl;
        public string _listTitle;
        private string _listId;
        private string _Tenant;
        private string _DriveID;
        private string _TaxonomyFieldInternalName = "ObjectClassification";
        public Microsoft.Graph.GraphServiceClient _graphClient;
        private Microsoft.Graph.Beta.GraphServiceClient _graphClientBeta;
        private Azure.Core.AccessToken _accessToken;
        public TokenType tokenType;
        public Azure.Identity.ClientSecretCredential _credential;
        public enum TokenType
        {
            App,
            DelegatedUser
        }

        public GraphService(IConfiguration? config)
        {
        }

        /// <summary>
        /// Fetch changed list items using the delta endpoint.
        /// </summary>
        public async Task<(List<string> ChangedItems, string? NewDeltaLink)> GetChangedItemsAsync(string tenantID, string siteId, string listId, string? deltaLink)
        {
            var changeditems = new List<string>();

            _accessToken = await GetGraphCLientToken(tenantID);
            // 3️⃣ Create HttpClient
            var HttpClient = await GetHttpClient(tenantID);
            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken.Token);
            HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


            // 4️⃣ Call delta endpoint
            var deltaUrl = deltaLink ?? $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/delta";

            HttpResponseMessage response = await HttpClient.GetAsync(deltaUrl);

            if (!response.IsSuccessStatusCode)
            {
                // If deltaLink expired or invalid, retry without deltaLink
                if (deltaLink != null && response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                {
                    Debug.WriteLine("Delta link expired. Starting fresh.");
                    deltaUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/delta";
                    response = await HttpClient.GetAsync(deltaUrl);
                    response.EnsureSuccessStatusCode(); // will throw if still failing
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Graph delta call failed: {response.StatusCode} - {errorContent}");
                }
            }

            /*
            //Old code 

            //var deltaUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/delta";
            var response = new HttpResponseMessage();
            try
            {
                response = await HttpClient.GetAsync(deltaUrl);
                response.EnsureSuccessStatusCode();
            }
            catch (Exception ex)
            {
                try
                {

                }
                catch (Exception exnolink)
                {
                    //Suggestion by CHATGPT. Unlikely 
                    //Delta queries in Microsoft Graph expire quickly(24 hours max). If you are passing a saved deltaLink that’s older, Graph will reject it with 400 Bad Request.
                    //Solution: Try calling the delta endpoint without the deltaLink to get a fresh delta token:


                    Debug.WriteLine($"Error calling delta endpoint: {exnolink.Message}");
                    Console.WriteLine($"Error calling delta endpoint: {exnolink.Message}");
                }
            }

            */
            var json = await response.Content.ReadAsStringAsync();
            string newDeltaLink = "";
            var data = System.Text.Json.JsonSerializer.Deserialize<JsonDocument>(json);
            var items = data.RootElement.GetProperty("value").EnumerateArray().ToList();

            if (data.RootElement.TryGetProperty("value", out JsonElement valueArray))
            {
                try
                {
                    changeditems = valueArray.EnumerateArray()
                                 .Select(item => item.GetProperty("id").GetString()!)
                                   .ToList();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing changed items: {ex.Message}");
                }
            }
            newDeltaLink = data.RootElement.TryGetProperty("@odata.deltaLink", out var link) ? link.GetString() : null;
            return (changeditems, newDeltaLink);
        }
        public async Task<string> GetItemContentAsync(string tenantID, string siteId, string listId, string ItemID)
        {
            var HttpClient = await GetHttpClient(tenantID);
            var contentUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/{ItemID}/driveItem/content";
            var response = await HttpClient.GetAsync(contentUrl);
            response.EnsureSuccessStatusCode();
            var fileBytes = await response.Content.ReadAsByteArrayAsync();
            var OriginalColor = Console.ForegroundColor;
            string body = "";
            try
            {
                using var stream = new MemoryStream(fileBytes);
                using var wordDoc = WordprocessingDocument.Open(stream, false);
                body = (wordDoc.MainDocumentPart.Document.Body.InnerText);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex);
            }
            Console.ForegroundColor = OriginalColor;


            return (body);
        }
        public async Task<(Microsoft.Graph.Models.ListItem, string FileContent)> GetListItemAsync(string tenantID, string siteId, string listId, string itemId)
        {
            Microsoft.Graph.Models.ListItem item = null;
            string filecontent = "";
            // Fetch the entire item, including its field values
            try
            {
                Microsoft.Graph.GraphServiceClient _graphClient = await GetGraphClient(tenantID);
                item = await _graphClient.Sites[siteId]
                                             .Lists[listId]
                                             .Items[itemId]
                                             .GetAsync(config =>
                                             {
                                                 config.QueryParameters.Expand = new[] { "fields" };
                                             });
                filecontent = await GetItemContentAsync(tenantID, siteId, listId, itemId);
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                Debug.WriteLine($"Isssue fetching list item {itemId} was not found in {listId} : {ex.Message}");
                Console.WriteLine($"Isssue fetching list item {itemId} was not found in {listId} : {ex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Isssue fetching list item {itemId} : {ex.Message}");
                Console.WriteLine($"Isssue fetching list item {itemId} : {ex.Message}");
            }
            return (item, filecontent);
        }
        public async Task<HttpClient> GetHttpClient(string tenantID, bool FromGraphBeta = false)
        {
            Parameters par = new Parameters(tenantID);
            this.tokenType = par.tokenType;
            var credential = new ClientSecretCredential(par.tenantId, par.clientId, par.clientSecret);
            this._credential = credential;
            var scopes = new[] { par.Scopes };
            if (FromGraphBeta)
            {
                Microsoft.Graph.Beta.GraphServiceClient _graphClient = new Microsoft.Graph.Beta.GraphServiceClient(this._credential, scopes);
            }
            else
            {
                Microsoft.Graph.GraphServiceClient _graphClient = new Microsoft.Graph.GraphServiceClient(this._credential, scopes);
            }

            _accessToken = await GetGraphCLientToken(tenantID);
            var HttpClient = new HttpClient();
            HttpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _accessToken.Token);
            return HttpClient;
        }
        public async Task<Microsoft.Graph.GraphServiceClient> GetGraphClient(string tenantID)
        {
            Parameters par = new Parameters(tenantID);
            this.tokenType = par.tokenType;
            Microsoft.Graph.GraphServiceClient _graphClient = new Microsoft.Graph.GraphServiceClient(
            this._credential = new ClientSecretCredential(par.tenantId, par.clientId, par.clientSecret),
            new[] { par.Scopes });
            return _graphClient;
        }
        public async Task<Microsoft.Graph.Beta.GraphServiceClient> GetGraphClientBeta(string tenantID)
        {
            Parameters par = new Parameters(tenantID);
            this.tokenType = par.tokenType;
            Microsoft.Graph.Beta.GraphServiceClient _graphClientBeta = new Microsoft.Graph.Beta.GraphServiceClient(
            this._credential = new ClientSecretCredential(par.tenantId, par.clientId, par.clientSecret),
            new[] { par.Scopes });
            return _graphClientBeta;
        }
        public async Task<List<(string Label, string Guid)>> GetAllTermsFromTermSetAsync(string tenantID, string siteId, string termSetId)
        {
            Microsoft.Graph.GraphServiceClient _graphClient = await GetGraphClient(tenantID);
            var allTerms = new List<(string Label, string Guid)>();

            try
            {
                // Call Graph to get the terms
                var response = await _graphClient
                    .Sites[siteId]
                    .TermStore
                    .Sets[termSetId]
                    .Terms
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Expand = new[] { "children" }; // also get children
                        config.QueryParameters.Top = 999; // adjust if needed
                    });

                if (response?.Value != null)
                {
                    foreach (var term in response.Value)
                    {
                        Debug.WriteLine($"Label: {term.Labels[0].Name}", $"GUID : {term.Id}");
                        AddTermRecursive(term, allTerms);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return allTerms;
        }
        // Helper to recursively add term and its children
        private void AddTermRecursive(Microsoft.Graph.Models.TermStore.Term term, List<(string Label, string Guid)> list)
        {
            var label = term.Labels?.FirstOrDefault()?.Name ?? "(no label)";
            var id = term.Id ?? "(no id)";
            list.Add((label, id));

            if (term.Children != null && term.Children.Any())
            {
                foreach (var child in term.Children)
                {
                    AddTermRecursive(child, list);
                }
            }
        }
        public async Task<Boolean> UpdateItemClassificationGraphAPI_SPCSOM(string tenantID, string fieldname, string siteId, string listId, Microsoft.Graph.Models.ListItem item, string termlabel, string termguid, string termsetguid)
        {
            //Code deleted, not working
            return false;
        }

        //According to ChatGPT 
        //Because Graph does NOT support full taxonmy updates it is not possible to update the tax fields using that approach. 
        //For that reason we'll need to update thje item using SP and not graph
        public async Task<string> GetSPCLientToken(string tenantID)
        {
            Azure.Core.AccessToken? Token;
            try
            {
                Parameters par = new Parameters(tenantID);
                this.tokenType = par.tokenType;
                //string[] scopes = new string[] { $"https://{tenantID}.sharepoint.com/.default" };
                string[] scopes = new string[] { $"https://{this._Domain}/.default" };
                //As per ChatGPT Suggestion
                //string[] scopes = new string[] { "https://00000003-0000-0ff1-ce00-000000000000/.default" };

                var app = ConfidentialClientApplicationBuilder
                            .Create(par.clientId)
                            .WithClientSecret(par.clientSecret)
                            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantID}"))
                            .Build();

                var result = await app.AcquireTokenForClient(scopes).WithForceRefresh(true).ExecuteAsync();
                string spToken = result.AccessToken;
                if (spToken == null || spToken == "")
                {
                    throw new InvalidOperationException("Failed to acquire Sharepoint AccessToken.");
                    return null;
                }
                else
                {
                    //Check token audience 
                    string payload = spToken.Split('.')[1]; // JWT middle part
                    payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '='); // pad base64
                    var jsonBytes = Convert.FromBase64String(payload);
                    var json = Encoding.UTF8.GetString(jsonBytes);

                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("aud", out var aud))
                    {
                        Console.WriteLine($"Token audience: {aud}");
                        Debug.WriteLine($"Token audience: {aud}");
                    }
                    else
                    {
                        Console.WriteLine("No audience found in token!");
                    }
                    return spToken;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error acquiring Sharepoint token: {ex.Message}");
                Debug.WriteLine($"Error acquiring Sharepoint token: {ex.Message}");
                return null;
            }
        }

        public async Task<Azure.Core.AccessToken> GetGraphCLientToken(string tenantID, bool FromGraphBeta = false, TokenType? tokenType = TokenType.App)
        {
            Azure.Core.AccessToken? Token;
            try
            {
                Parameters par = new Parameters(tenantID, tokenType);
                this.tokenType = par.tokenType;
                var credential = new ClientSecretCredential(par.tenantId, par.clientId, par.clientSecret);
                this._credential = credential;
                var scopes = new[] { par.Scopes };
                if (FromGraphBeta)
                {
                    Microsoft.Graph.Beta.GraphServiceClient _graphClient = new Microsoft.Graph.Beta.GraphServiceClient(this._credential, scopes);
                }
                else
                {
                    Microsoft.Graph.GraphServiceClient _graphClient = new Microsoft.Graph.GraphServiceClient(this._credential, scopes);
                }

                // 2️⃣ Acquire token manually
                var tokenRequestContext = new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
                Token = await credential.GetTokenAsync(tokenRequestContext);
                //Assigning it directly sometimes results in unauthorized.
                //Token = await credential.GetTokenAsync(tokenRequestContext);
                if (Token == null)
                {
                    throw new InvalidOperationException("Failed to acquire Graph AccessToken.");
                    return default(Azure.Core.AccessToken);
                }
                else
                {
                    return Token.Value;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error acquiring Graph token: {ex.Message}");
                Debug.WriteLine($"Error acquiring Graph token: {ex.Message}");
                return default(Azure.Core.AccessToken);
            }
        }

        //NOT WORKING
        public async Task<Boolean> UpdateItemClassificationSPAPI(string tenantID, string fieldname, string siteId, string listId, Microsoft.Graph.Models.ListItem item, string termlabel, string termguid, string termsetguid)
        {
            string spToken = await GetSPCLientToken(tenantID);

            if (string.IsNullOrEmpty(spToken))
            {
                Console.WriteLine("Invalid SharePoint token");
                return false;
            }

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", spToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // 1️⃣ Get Form Digest
                var contextInfoUrl = $"{this._SiteUrl}/_api/contextinfo";
                var digestResponse = await client.PostAsync(contextInfoUrl, null);
                digestResponse.EnsureSuccessStatusCode();

                var digestJson = JsonDocument.Parse(await digestResponse.Content.ReadAsStringAsync());
                string digestValue = digestJson.RootElement
                                               .GetProperty("d")
                                               .GetProperty("GetContextWebInformation")
                                               .GetProperty("FormDigestValue")
                                               .GetString();

                // 2️⃣ Build payload for taxonomy field
                var payload = new Dictionary<string, object>
                {
                    ["__metadata"] = new { type = $"SP.Data.{this._listTitle.Replace(" ", "_x0020_")}ListItem" },
                    [this._TaxonomyFieldInternalName] = new Dictionary<string, object>
                    {
                        ["__metadata"] = new { type = "SP.Taxonomy.TaxonomyFieldValue" },
                        ["Label"] = termlabel,
                        ["TermGuid"] = termguid,
                        ["WssId"] = -1
                    }
                };

                // 3️⃣ Serialize payload
                var jsonContent = new StringContent(
                    System.Text.Json.JsonSerializer.Serialize(payload),
                    Encoding.UTF8,
                    "application/json;odata=verbose"
                );

                // 4️⃣ Build update request
                var updateUrl = $"{this._SiteUrl}/_api/web/lists('{listId}')/items({item.Id})";
                using var request = new HttpRequestMessage(HttpMethod.Post, updateUrl);
                request.Headers.Add("X-RequestDigest", digestValue);
                request.Headers.Add("IF-MATCH", "*");            // Merge with any existing version
                request.Headers.Add("X-HTTP-Method", "MERGE");   // Required for update
                request.Content = jsonContent;

                // 5️⃣ Send update request
                var response = await client.SendAsync(request);
                if (!response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Update failed: {response.StatusCode}, {content}");
                    return false;
                }

                Console.WriteLine($"Item {item.Id} updated successfully!");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating item: {ex.Message}");
                return false;
            }
        }

        //According to ChatGPT 
        //Because Graph does NOT support full taxonmy updates it is not possible to update the tax fields using that approach. 
        //For that reason we'll need to update thje item using SP and not graph
        //Therefore at the moment this method is void
        public async Task<Boolean> UpdateItemClassificationGraphAPI(string tenantID, string fieldname, string siteId, string listId, Microsoft.Graph.Models.ListItem item, string termlabel, string termguid, string termsetguid)
        {
            try
            {
                string driveid = "";
                string driveitemid = "";
                string webUrl = "";
                string endpoint = "";
                string taxonomyFieldInternalName = "ObjectClassification_o"; // your field internal name


                //Try with graph Beta. 
                _accessToken = await GetGraphCLientToken(tenantID);

                //Dummy payload
                var payload = new Dictionary<string, object>
                {
                    ["ObjectClassification@odata.type"] = "SP.Taxonomy.TaxonomyFieldValue",
                    [this._TaxonomyFieldInternalName] = new
                    {
                        Label = "Confidential",
                        TermGuid = "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3",
                        WssId = -1
                    }
                };
                var json = System.Text.Json.JsonSerializer.Serialize(payload);

                using (HttpClient client = new HttpClient())
                {



                    //Simple test to change title, works. Taxonomy not yet!!!
                    //json = "{\"Title\": \"New Document Title\"}";

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken.Token);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    //Get lists
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists";
                    // HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.GetAsync(endpoint);
                    var jsonlists = await response.Content.ReadAsStringAsync();
                    var doclists = JsonDocument.Parse(jsonlists);
                    if (response.IsSuccessStatusCode)
                    {
                        var Lists = doclists.RootElement.GetProperty("value");

                        Console.WriteLine("Lists in the site :");
                        foreach (var list in Lists.EnumerateArray())
                        {
                            string name = list.GetProperty("name").GetString();
                            string displayName = list.GetProperty("displayName").GetString();
                            Debug.WriteLine($"List : {displayName} ({name})");
                            Console.WriteLine($"List : {displayName} ({name})");
                        }
                    }


                    //Check if internal name is correct 
                    //endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/fields?$filter=typeAsString eq 'TaxonomyFieldType'";
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/fields";
                    // HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                    response = await client.GetAsync(endpoint);
                    var jsonfields = await response.Content.ReadAsStringAsync();
                    var docfields = JsonDocument.Parse(jsonfields);
                    if (response.IsSuccessStatusCode)
                    {
                        var fields = docfields.RootElement.GetProperty("value");

                        Console.WriteLine("Columns in the list:");
                        foreach (var field in fields.EnumerateArray())
                        {
                            string name = field.GetProperty("name").GetString();
                            string displayName = field.GetProperty("displayName").GetString();
                            // fallback if type not present
                            string type = field.TryGetProperty("type", out var typeProp) ? typeProp.GetString() : "N/A";
                            Debug.WriteLine($"{displayName} ({name}) - {type}");
                        }
                    }

                    //Get listColumns
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{this._SiteID}/lists/{listId}/columns";
                    HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                    response = await client.GetAsync(endpoint);
                    //The next code is currently repeated and should be centralized in a common function. For the POC this is not yet done. 
                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        //Console.WriteLine($"Response ({response.StatusCode}): {responseBody}");
                        using JsonDocument doc = JsonDocument.Parse(responseBody);
                        var columns = doc.RootElement.GetProperty("value");

                        Console.WriteLine("Columns in the list:");
                        foreach (var column in columns.EnumerateArray())
                        {
                            string name = column.GetProperty("name").GetString();
                            string displayName = column.GetProperty("displayName").GetString();
                            // fallback if type not present
                            string type = column.TryGetProperty("type", out var typeProp) ? typeProp.GetString() : "N/A";
                            Debug.WriteLine($"{displayName} ({name}) - {type}");
                            Console.WriteLine($"{displayName} ({name}) - {type}");
                        }
                    }
                    else
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Failed to get taxonomyfields. Status: {response.StatusCode}");
                        Debug.WriteLine(responseBody);
                        return false;
                    }


                    //Get the corresponding driveID
                    driveid = await GetDocLibDriveID(tenantID, _listTitle, client);

                    //Now Get the correcponding DriveItemID 
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/{item.Id}?expand=driveItem";

                    content = new StringContent(json, Encoding.UTF8, "application/json");
                    response = await client.GetAsync(endpoint);
                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Console.WriteLine($"Response ({response.StatusCode}): {responseBody}");
                        using var jsonDoc = JsonDocument.Parse(responseBody);
                        var root = jsonDoc.RootElement;

                        // List Item ID
                        //var listItemId = root.GetProperty("id").GetString();

                        // DriveItem (may be missing if not a doclib)
                        if (root.TryGetProperty("driveItem", out var driveItem))
                        {
                            driveitemid = driveItem.GetProperty("id").GetString();
                            webUrl = driveItem.GetProperty("webUrl").GetString();

                            Debug.WriteLine($"✅ DriveItem ID: {driveitemid}");
                            Debug.WriteLine($"🌐 Web URL: {webUrl}");
                        }
                        else
                        {
                            Console.WriteLine("⚠️ No driveItem found — this list item is not a document library file.");
                        }
                    }
                    else
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Failed to get driveitemid. Status: {response.StatusCode}");
                        Debug.WriteLine(responseBody);
                        return false;
                    }

                    // ======= Graph endpoint =======
                    //THIS IS FOR LISTS. For doclibs DriveID needs to be used 
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveid}/items/{driveitemid}/listItem";
                    // Step 1: Get the List Item ID
                    var listItemResponse = await client.GetAsync(endpoint);

                    if (!listItemResponse.IsSuccessStatusCode)
                    {
                        string error = await listItemResponse.Content.ReadAsStringAsync();
                        Console.WriteLine($"Error getting list item: {listItemResponse.StatusCode}");
                        Console.WriteLine(error);
                        return false;
                    }

                    //Try update Title field and objectclassificationtext
                    //Step 2: Update the Title field as test
                    var listItemContent = await listItemResponse.Content.ReadAsStringAsync();
                    dynamic listItem = Newtonsoft.Json.JsonConvert.DeserializeObject(listItemContent);
                    string listItemId = listItem.id;

                    var body = new
                    {
                        Title = "Test Updated Document Title Ruben",
                        ObjectClassificationText = termlabel
                    };
                    content = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");
                    endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/{listItemId}/fields";
                    response = await client.PatchAsync(endpoint, content);
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("field Title updated successfully!");
                    }
                    else
                    {
                        var OriginalColor = Console.ForegroundColor;
                        Console.ForegroundColor = ConsoleColor.Red;
                        //string responseBody = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Failed to update title field. Status for file {listItem}:  {response.StatusCode}");
                        Console.WriteLine($"Failed to update title field. Status for file {listItem}:  {response.StatusCode}");
                        //Debug.WriteLine(responseBody);
                        Console.ForegroundColor = OriginalColor;
                        return false;
                    }


                    //Try update Title field and objectclassificationtext
                    //Step 3: Update the a Objectclassification choice field 
                    var bodyObjectClassificationDropDown = new
                    {
                        ObjectClassificationDropDown = termlabel,
                    };

                    content = new StringContent(JsonConvert.SerializeObject(bodyObjectClassificationDropDown), Encoding.UTF8, "application/json");

                    //response = await client.PatchAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{driveid}/items/{driveitemid}/fields", content);
                    response = await client.PatchAsync(endpoint, content);
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("field objectclassificationDropDown updated successfully!");
                    }
                    else
                    {
                        var OriginalColor = Console.ForegroundColor;
                        Console.ForegroundColor = ConsoleColor.Red;
                        //string responseBody = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Failed to update objectclassificationDropDown status for file {listItem}:  {response.StatusCode}");
                        Console.WriteLine($"Failed to update objectclassificationDropDown status for file {listItem}:  {response.StatusCode}");
                        //Debug.WriteLine(responseBody);
                        Console.ForegroundColor = OriginalColor;
                        return false;
                    }


                    //Step4 Update taxonmyField. 
                    //TESTPOCTESTPOC
                    //Tax fields not working. 401 or 400. ChatGPT states itaxonomyfields can be read by GRaph API but not updated

                    #region PayLoadtest
                    /*
                    var payload = new Dictionary<string, object>
                    {
                        ["fields"] = new Dictionary<string, object>
                        {
                            [fieldname] = new Dictionary<string, object>
                            {
                                ["Label"] = "Confidential",
                                ["TermGuid"] = "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3",
                                ["WssId"] = -1
                            }
                        }
                    };
                    /*
                    var payload = new Dictionary<string, object>
                    {
                        [fieldname] = new Dictionary<string, object>
                        {
                            ["Label"] = "Confidential",
                            ["TermGuid"] = "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3",
                            ["WssId"] = -1
                        }
                    };

                    var payload = new Dictionary<string, object>
                    {
                        ["fields"] = new Dictionary<string, object>
                        {
                            [$"{taxonomyFieldInternalName}@odata.type"] = "SP.Taxonomy.TaxonomyFieldValue",
                            [taxonomyFieldInternalName] = $"{termlabel}|{termguid}",
                            ["WssId"] = -1
                        }
                    };

                    var payload = new
                    {
                        fields = new Dictionary<string, object>
                        {
                            ["ObjectClassification@odata.type"] = "SP.Taxonomy.TaxonomyFieldValue",
                            ["ObjectClassification"] = $"{termlabel}|{termguid}"
                        }
                    };

                    var payload = new
                    {
                        fields = new Dictionary<string, object>
                        {
                            ["ObjectClassification@odata.type"] = "SP.Taxonomy.TaxonomyFieldValue",
                            ["ObjectClassification"] = new
                            {
                                Label = Confidential,
                                TermGuid = "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3",
                                WssId = -1
                            }
                        }
                    };
                    */
                    #endregion

                    payload = new Dictionary<string, object>
                    {
                        ["ObjectClassification@odata.type"] = "SP.Taxonomy.TaxonomyFieldValue",
                        [this._TaxonomyFieldInternalName] = new
                        {
                            Label = "Confidential",
                            TermGuid = "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3",
                            WssId = -1
                        }
                    };

                    json = System.Text.Json.JsonSerializer.Serialize(payload);
                    content = new StringContent(json, Encoding.UTF8, "application/json");
                    //endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{driveid}/items/{driveitemid}/fields";
                    response = await client.PatchAsync(endpoint, content);


                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("taxonomyfield updated successfully!");
                    }
                    else
                    {
                        var OriginalColor = Console.ForegroundColor;
                        Console.ForegroundColor = ConsoleColor.Red;
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Failed to update taxonomyfield field. Status for file {listItemId}:  {response.StatusCode}");
                        Console.WriteLine($"Failed to update taxonomyfield . Status for file {listItemId}:  {response.StatusCode}");
                        Debug.WriteLine(responseBody);
                        Console.ForegroundColor = OriginalColor;
                        //Ignore return because it is known that Grapch API'sface diffficulties returning hidden tax fiuelds making it impossible to update the guid of the new classification.
                        //return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error changing classification: {ex.Message}");
                return false;
            }
            return true;
        }

        public async Task<Boolean> UpdateItemClassificationGraphAPIAttempt2(string tenantID, string fieldname, string siteId, string listId, Microsoft.Graph.Models.ListItem item, string termlabel, string termguid, string termsetguid)
        {
            try
            {
                Parameters par = new Parameters(tenantID);
                this.tokenType = par.tokenType;
                var client = await GetGraphClient(par.tenantId);

                var fields = new Microsoft.Graph.Models.FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
        {
            {
                      this._TaxonomyFieldInternalName, new Dictionary<string, object>
                              {
                                 { "Label", "Confidential"},
                                { "TermGuid", "e5bc934d-989f-4dd4-9add-7ea4b6bb3cf3"},
                                { "WssId", -1 }
                            }
            }
        }
                };

                await client
                   .Sites[siteId]
                   .Lists[listId]
                   .Items[item.Id]
                   .Fields
                    .PatchAsync(fields);

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error changing classification: {ex.Message}");
                return false;
            }
            return true;
        }

        //NOT WORKING
        public async Task<Boolean> UpdateItemClassificationGraphAPIThrough_SP_RESTAPI(string tenantID, string fieldname, string siteId, string listId, Microsoft.Graph.Models.ListItem item, string termlabel, string termguid, string termsetguid)
        {
            try
            {
                //NOT WORKING
                //First check access 
                Parameters par = new Parameters(tenantID);
                this.tokenType = par.tokenType;
                var authManager = new PnP.Framework.AuthenticationManager();
                using (var context = authManager.GetACSAppOnlyContext(this._SiteUrl, par.clientId, par.clientSecret))
                {
                    var web = context.Web;
                    context.Load(web);
                    context.ExecuteQuery();
                    Console.WriteLine("Web title: " + web.Title);
                }

                using var http = new HttpClient();
                Microsoft.Graph.Beta.GraphServiceClient _graphClientBeta = await GetGraphClientBeta(tenantID);

                var endpoint = $"{_SiteUrl}/_api/web/lists(guid'{listId}')/items({item.Id})";
                var tokenContext = new TokenRequestContext(new[] { "https://lls6.sharepoint.com/.default" });
                //var tokenContext = new TokenRequestContext(new[] { $"{new Uri(SiteUrl).Scheme}://{new Uri(SiteUrl).Host}/.default" });
                var token = await this._credential.GetTokenAsync(tokenContext);

                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
                http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var listMetaUrl = $"{_SiteUrl}/_api/web/lists/getbytitle('{_listTitle}')?$select=ListItemEntityTypeFullName";
                var metaResp = await http.GetAsync(listMetaUrl);
                if (!metaResp.IsSuccessStatusCode)
                {
                    var err = await metaResp.Content.ReadAsStringAsync();
                    Console.WriteLine($"❌ Failed to get list metadata: {metaResp.StatusCode} - {err}");
                    return false;
                }
                dynamic meta = JsonConvert.DeserializeObject(await metaResp.Content.ReadAsStringAsync());
                string entityType = meta.d.ListItemEntityTypeFullName;

                var endpoint1 = $"{siteId}/_api/web/lists(guid'{listId}')/items({item.Id})";


                var payload = new Dictionary<string, object>
                {
                    ["__metadata"] = new { type = "SP.Data.ObjectClassification" },
                    [fieldname] = $"{termlabel}|{termguid}"
                };

                var json = Newtonsoft.Json.JsonConvert.SerializeObject(payload);

                var req = new HttpRequestMessage(HttpMethod.Post, endpoint);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
                req.Headers.Add("IF-MATCH", "*");
                req.Headers.Add("X-HTTP-Method", "MERGE");
                //req.Content = new StringContent(json, Encoding.UTF8, "application/json;odata=verbose");
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                var resp = await http.SendAsync(req);
                string result = await resp.Content.ReadAsStringAsync();

                if (!resp.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error: {resp.StatusCode} - {result}");
                    return false;
                }

                Console.WriteLine("✅ Managed metadata updated successfully!");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error changing classification: {ex.Message}");
                return false;
            }
        }

        public async Task<string> GetWebInfo(string tenantID, string siteId, string listId)
        {
            //string endpoint = "";
            try
            {
                // Extract site composite ID (everything between "/sites/" and "/lists/")
                // string endpoint = "";
                this._FullSiteId = siteId;
                var parts = siteId.Split(',');
                var siteCompositeId = $"{parts[0]},{parts[1]},{parts[2]}";
                if (parts.Length >= 3)
                {
                    Debug.WriteLine(siteCompositeId);
                }
                else
                {
                    Console.WriteLine("siteId does not contain enough parts.");
                }
                this._Domain = parts[0];
                this._SiteID = parts[0] + "," + parts[1];
                this._SubSiteId = parts[2];
                this._listId = listId;

                // Call Graph
                _accessToken = await GetGraphCLientToken(tenantID);
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken.Token);
                string endpoint = $"https://graph.microsoft.com/v1.0/sites/{siteCompositeId}";
                var response = await client.GetAsync(endpoint);
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                this._SiteUrl = doc.RootElement.GetProperty("webUrl").GetString();

                Debug.WriteLine($"Resolved SharePoint Site URL: {this._SiteUrl}");
                //this._SiteUrl = siteUrl;
                this._Tenant = parts[0];
                this._SiteRelativeUrl = "/sites/" + Regex.Split(this._SiteUrl, "/sites/", RegexOptions.IgnoreCase)[1];
                this._SiteURefForPermissions = _Tenant + ":" + _SiteRelativeUrl + ":";

                //get list name 
                this._listTitle = await GetListTitle(client);

                //Get corresponding drive ID
                this._DriveID = await GetDocLibDriveID(tenantID, this._listTitle, client);
                return this._SiteUrl;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching site URL: {ex.Message}");
                return null;
            }
        }

        public async Task<Boolean> AssignAppWritePermissionsOnSite(string tenantID)
        {
            //NOT WORKING
            Parameters par = new Parameters(tenantID);
            this.tokenType = par.tokenType;
            //string siteId = "YOUR_SITE_ID"; // Get via Graph: /sites/{hostname}:{site-relative-path}
            //string url = $"https://graph.microsoft.com/v1.0/{this._SiteRelativeUrl}/permissions";
            string url = $"https://graph.microsoft.com/v1.0/sites/{this._SiteURefForPermissions}:/permissions";
            _accessToken = await GetGraphCLientToken(tenantID);

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken.Token);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var body = new
                {
                    roles = new[] { "write" },
                    grantee = new
                    {
                        application = new { id = par.clientId }
                    }
                };

                var content = new StringContent(System.Text.Json.JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

                var response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                    Console.WriteLine("Site-level permission granted!");
                else
                    Console.WriteLine($"Error: {response.StatusCode} - {await response.Content.ReadAsStringAsync()}");
            }
            return false;
        }

        public async Task<Boolean> GetAppPermissionsOnSite(string tenantID)
        {
            _accessToken = await GetGraphCLientToken(tenantID);
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", this._accessToken.Token);

            var response = await client.GetAsync("https://graph.microsoft.com/v1.0/sites/lls6.sharepoint.com:/sites/SP-EventReceivers-Test");
            var content = await response.Content.ReadAsStringAsync();
            //Console.WriteLine(content);
            return true;
        }

        //According to ChaGPT this should be possible but then ChatGPT States it cant be done through GraphAPI > 5
        //Not working
        public async Task<Boolean> SetAppPermissionsOnSite(string tenantID)
        {
            Parameters par = new Parameters(tenantID);
            this.tokenType = par.tokenType;
            _accessToken = await GetGraphCLientToken(tenantID);
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", this._accessToken.Token);

            var permission = new Microsoft.Graph.Models.Permission
            {
                Roles = new List<string> { "write" },
                GrantedToIdentities = new List<Microsoft.Graph.Models.IdentitySet>
    {
        new Microsoft.Graph.Models.IdentitySet
        {
            Application = new Microsoft.Graph.Models.Identity
            {
                Id = par.clientId,
                DisplayName = "DotNetSPEventreceivers"
            }
        }
    }
            };

            //await _graphClient.Sites[this._SiteID].Permissions.AddAsync(permission);
            return true;

        }

        public async Task<bool> GetAppPermissionsOnList(string tenantID, string listName)
        {
            bool result = false;
            string driveid = "";
            var token = await GetGraphCLientToken(tenantID);
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
                string endpoint = $"https://graph.microsoft.com/v1.0/sites/{this._SiteURefForPermissions}/lists/{listName}/items";

                // Read test
                var readresponse = await client.GetAsync(endpoint);
                if (readresponse.IsSuccessStatusCode)
                {
                    Debug.WriteLine("Read access confirmed!");
                    Console.WriteLine("Read access confirmed!");

                    result = true;
                }
                else
                {
                    Console.WriteLine($"Failed to read list. Status: {readresponse.StatusCode}");
                    string error = await readresponse.Content.ReadAsStringAsync();
                    Debug.WriteLine(error);
                    result = false;
                }


                //Skip this . Tested and works but it will result in a infinite loop
                return result;

                // --- WRITE: create a test file in document library ---
                //Get the doclib driveid
                driveid = await GetDocLibDriveID(tenantID, listName, client);


                string fileName = "TestFileFromApp.txt";
                var fileContent = new ByteArrayContent(Encoding.UTF8.GetBytes("Hello from app!"));
                fileContent.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
                string writeEndpoint = $"https://graph.microsoft.com/v1.0/sites/{this._SiteID}/drives/{driveid}/root:/{fileName}:/content";
                HttpResponseMessage writeResponse = await client.PutAsync(writeEndpoint, fileContent);


                //var createResponse = await client.PostAsync(endpoint, jsonContent);
                if (writeResponse.IsSuccessStatusCode)
                {
                    Debug.WriteLine("Write access confirmed! Item created.");
                    result = true;
                }
                else
                {
                    Debug.WriteLine($"Failed to write to list. Status: {writeResponse.StatusCode}");
                    string error = await writeResponse.Content.ReadAsStringAsync();
                    Debug.WriteLine(error);
                    result = false;
                }
            }
            return result;
        }

        public async Task<String> GetDocLibDriveID(string tenantID, string listName, System.Net.Http.HttpClient client)
        {
            bool result = false;
            string driveid = "";
            try
            {
                var endpoint = $"https://graph.microsoft.com/v1.0/sites/{this._SiteID}/drives";
                HttpResponseMessage response = await client.GetAsync(endpoint);
                response.EnsureSuccessStatusCode();

                string content = await response.Content.ReadAsStringAsync();
                JObject json = JObject.Parse(content);
                response.EnsureSuccessStatusCode();
                foreach (var drive in json["value"])
                {
                    if (string.Equals(drive["name"].ToString(), listName, StringComparison.OrdinalIgnoreCase))
                    {
                        driveid = drive["id"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching drive ID: {ex.Message}");
                return null;
            }
            return driveid;
        }
        //public async Task<string> GetListTitle(Microsoft.Graph.GraphServiceClient client, string siteId, string listId)
        public async Task<string> GetListTitle(HttpClient client)
        {
            string Listendpoint = $"https://graph.microsoft.com/v1.0/sites/{this._SiteID}/lists/{this._listId}";
            try
            {
                // var list = await this._graphClient.Sites[this._SiteID].Lists[this._listId].GetAsync();
                HttpResponseMessage response = await client.GetAsync(Listendpoint);
                response.EnsureSuccessStatusCode();
                var json = await response.Content.ReadAsStringAsync();
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;
                string displayName = root.GetProperty("displayName").GetString();
                //Console.WriteLine($"List ID: {list.Id}");
                // Console.WriteLine($"List Title: {list.DisplayName}");

                return displayName;
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }
    }
}
