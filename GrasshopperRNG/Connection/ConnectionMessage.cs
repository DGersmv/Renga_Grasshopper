using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Connection
{
    /// <summary>
    /// Base message class for communication protocol
    /// </summary>
    public class ConnectionMessage
    {
        [JsonProperty("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        [JsonProperty("command")]
        public string Command { get; set; }

        [JsonProperty("data")]
        public JObject Data { get; set; }

        [JsonProperty("timestamp")]
        public string Timestamp { get; set; } = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static ConnectionMessage FromJson(string json)
        {
            return JsonConvert.DeserializeObject<ConnectionMessage>(json);
        }
    }

    /// <summary>
    /// Response message from server
    /// </summary>
    public class ConnectionResponse
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("success")]
        public bool Success { get; set; }

        [JsonProperty("data")]
        public JObject Data { get; set; }

        [JsonProperty("error")]
        public string Error { get; set; }

        [JsonProperty("timestamp")]
        public string Timestamp { get; set; }

        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static ConnectionResponse FromJson(string json)
        {
            return JsonConvert.DeserializeObject<ConnectionResponse>(json);
        }
    }
}



