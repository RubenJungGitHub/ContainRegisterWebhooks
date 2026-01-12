using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;
using RJ_SPEventReceiversWebhookSubscribe;
using RJ_SPEventReceiversWebhookSubscribe.Classes;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Threading.Tasks;
using RJ_SPEventReceiversWebhookSubscribeOriginal.Classes;

/*
 How to run this:

Create a new ASP.NET Core Web API project (or minimal API).

Replace your Program.cs content with the above code.

Run your app (dotnet run).

Expose it with ngrok for HTTPS during testing, e.g., ngrok http 5000.

Use the ngrok URL plus /api/webhook as the notificationUrl when registering the webhook in SharePoint.
 */

//First try and register WebHook on SP list, then run listeners to capture events

//RegisterWebhook RegWebHook = new RegisterWebhook();
//RegWebHook.RegisterWebhookAsync(args).GetAwaiter().GetResult();

var builder = WebApplication.CreateBuilder(args);
// Optional: set URLs explicitly
builder.WebHost.UseUrls("https://localhost:5000");
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
//builder.Services.AddControllers();
//builder.WebHost.UseUrls("https://0.0.0.0:5000");
var app = builder.Build();
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();       // genereert swagger.json
    app.UseSwaggerUI();     // genereert UI op /swagger
}


const string ClientStateSecret = "bf6c26c2-cf09-405b-a57f-c6a2b58b2443";

app.UseHttpsRedirection();
//app.MapControllers();
// Map endpoints
app.MapGet("/api/WebHookListener", async (HttpRequest request, HttpResponse response) =>
{
    if (request.Query.TryGetValue("validationtoken", out var token))
    {
        response.StatusCode = 200;
        response.ContentType = "text/plain";
        await response.WriteAsync(token);
        return;
    }

    response.StatusCode = 200;
});

app.MapPost("/api/WebHookListener", async (HttpRequest request, HttpResponse response) =>
{
    using var reader = new StreamReader(request.Body);
    var body = await reader.ReadToEndAsync();
    Console.WriteLine("Webhook notification received:");
    Console.WriteLine(body);
    if (request.Query.ContainsKey("validationtoken"))
    {
        string validationToken = request.Query["validationtoken"];
        response.ContentType = "text/plain";
        await response.WriteAsync(validationToken);  // echo token
        return;
    }
    response.StatusCode = 200;
    await response.WriteAsync("Accepted");
});

//Set notificationURL 

// To  be securely configured, for POC not relevant yet
string notificationUrl = "https://gangrenous-kandis-unmunched.ngrok-free.dev/api/WebHookListener"; // FOR Ruben machine // Must be HTTPS and reachable
//string notificationUrl = "https://unsplit-zander-stressful.ngrok-free.dev/api/WebHookListener"; // FOR Frits machine // Must be HTTPS and reachable

int expirationTimeInMinutes = 10; //Max 43200 minutes (30 days) for list webhooks
// App registration details
var clientId = "f590b477-5bd7-47d6-8bda-36f77fa10afd";
var tenantId = "9a1b5f77-1f1a-40ac-b1a1-38617300f02a";
var clientSecret = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50";
// Start server without blocking
var host = app;

var serverTask = host.StartAsync();   // Start without blocking

await serverTask;                     // Wait for web server to be ready

//========================================================================================
//WORKING FOR ONE SUBSCRIPTION DON;T TOUCH!!!!
//RegisterWebhookOriginial regwh = new RegisterWebhookOriginial();
//await regwh.RegisterWebhookAsync(args, notificationUrl);
//========================================================================================


// ✅ Now safe to register webhook
try
{
    CheckListsOnWebHooksHTTPS CheckSubscriptions = new CheckListsOnWebHooksHTTPS();
    await CheckSubscriptions.CheckAllSiteLists(args, notificationUrl, expirationTimeInMinutes, tenantId, clientSecret, clientId, true);
    await host.WaitForShutdownAsync();
}
catch
{
    Console.WriteLine($"Ëxception running webhookmonitor {expirationTimeInMinutes}. RESTART SERVICE!!!");
}








// Classes to parse webhook payload
public class SharePointWebhookPayload
{
    public string ClientState { get; set; }
    public Notification[] Value { get; set; }
}

public class Notification
{
    public string SubscriptionId { get; set; }
    public string ClientState { get; set; }
    public string TenantId { get; set; }
    public string SiteUrl { get; set; }
    public string WebId { get; set; }
    public string ListId { get; set; }
}

