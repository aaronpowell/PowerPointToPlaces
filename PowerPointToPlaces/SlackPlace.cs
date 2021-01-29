using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PowerPointToPlaces
{
    class SlackPlace : ISendToPlace
    {
        // TODO: Enter your Slack WebHook URL here
        const string SLACK_URI = "";

        public bool CanSend(string note)
        {
            return note.StartsWith("SLACK:");
        }

        public async Task SendAsync(string note)
        {
            var client = new HttpClient();
            var text = note.Replace("SLACK:", "").Trim();
            Console.WriteLine($"\tSending {text} to slack");
            await client.PostAsync(SLACK_URI, new StringContent(JsonSerializer.Serialize(new { text })));
        }
    }
}
