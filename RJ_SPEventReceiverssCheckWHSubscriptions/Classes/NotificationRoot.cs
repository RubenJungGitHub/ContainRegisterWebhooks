namespace RJ_SPEventReceiversWebhookSubscribe.Classes
{
    using System.Text.Json.Serialization;


    public class NotificationRoot
    {
        [JsonPropertyName("value")]
        public List<NotificationValue>? Value { get; set; }
    }

    public class NotificationValue
    {
        [JsonPropertyName("subscriptionId")]
        public string? SubscriptionId { get; set; }

        [JsonPropertyName("clientState")]
        public string? ClientState { get; set; }

        [JsonPropertyName("resource")]
        public string? Resource { get; set; }

        [JsonPropertyName("tenantId")]
        public string? TenantId { get; set; }

        [JsonPropertyName("resourceData")]
        public ResourceData? ResourceData { get; set; }

        [JsonPropertyName("subscriptionExpirationDateTime")]
        public string? SubscriptionExpirationDateTime { get; set; }

        [JsonPropertyName("changeType")]
        public string? ChangeType { get; set; }
    }

    public class ResourceData
    {
        [JsonPropertyName("@odata.type")]
        public string? ODataType { get; set; }

        [JsonPropertyName("id")]
        public string? Id { get; set; }
    }
}