namespace  RJ_SPEventReceiversWebhookSubscribe.Controllers
{
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Graph;
    using  RJ_SPEventReceiversWebhookSubscribe.Classes;
    using System.Text.Json;
    using System.Text.Json.Serialization;

    [Route("api/[controller]")]
    [ApiController]
    public class WebhookListenerController : ControllerBase
    {
        private const string ClientState = "pE.8Q~ZQRGngJ1YliTP4EDC5bejaEl72LlBAzb50"; // Must match subscription


        // Validation Request from SharePoint
        
        [HttpGet]
        public IActionResult Get([FromQuery] string validationtoken)
        {
            if (!string.IsNullOrEmpty(validationtoken))
            {
                // Respond with the validation token to confirm webhook
                //return Ok();
                return Content(validationtoken, "text/plain");
            }
            return BadRequest();
        }
        

        [HttpPost]
        [Consumes("application/json", "text/json", "application/*+json")]
        //public async Task<IActionResult> Post([FromBody] dynamic? payload, [FromQuery] string? validationtoken = null)

        public IActionResult WebHookListener([FromQuery] string? validationtoken, [FromBody] NotificationRoot? notification)
    
            //public async Task<IActionResult> Post([FromQuery] string validationtoken)
            {

            Console.WriteLine("ListItemID : " + notification.Value[0].ResourceData.Id);

          if (!string.IsNullOrEmpty(validationtoken))
            {
                // Respond immediately with the validation token for registration
                //return Ok();
                return Content(validationtoken, "text/plain");
            }

            // Normal webhook processing
            Request.EnableBuffering(); // allows reading multiple times

            string body = "";
            using (var reader = new StreamReader(Request.Body, System.Text.Encoding.UTF8, leaveOpen: true))
            {
            //    body = await reader.ReadToEndAsync();
                Request.Body.Position = 0;
            }

            // Parse JSON to extract item IDs
            if (string.IsNullOrWhiteSpace(body))
            {
                Console.WriteLine("Received empty POST body. Ignoring.");
                Console.WriteLine("===============================================");
                return Ok(); // Nothing to do
            }

            // 3️⃣ Try to deserialize JSON only if body is not empty
            //NotificationRoot? notification = null;
            try
            {
                notification = JsonSerializer.Deserialize<NotificationRoot>(body);
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON deserialization failed: {ex.Message}");
                return BadRequest("Invalid JSON"); // Optional: could also just log
            }

            if (notification == null || notification.Value == null)
            {
                Console.WriteLine("No notifications found in payload.");
                return Ok();
            }

            // 4️⃣ Process each notification


            foreach (var val in notification.Value)
            {
             //   Console.WriteLine($"Item changed: {val.ResourceData.Id}");

                Console.WriteLine("Received notification:");
                Console.WriteLine(body);
                Console.WriteLine("===============================================");
                // Always return 200 OK quickly (<5s)
            }
            return Ok();
        }


    }

    public class NotificationPayload
    {
        public List<Notification> Value { get; set; }
    }

    public class Notification
    {
        public string SubscriptionId { get; set; }
        public string ClientState { get; set; }
        public string Resource { get; set; }
        public DateTime ExpirationDateTime { get; set; }
    }

}
