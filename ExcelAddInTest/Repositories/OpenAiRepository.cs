using System;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Azure.AI.OpenAI;

namespace ExcelAddInTest.Repositories
{
    public class OpenAiRepository
    {
        private const string DataPlaceholder = "<<DATA>>";

        public static async Task<string> GetAnalysis(string data, string apiKey, string promptTemplate)
        {
            if (string.IsNullOrEmpty(apiKey))
            {
                throw new ArgumentNullException(nameof(apiKey));
            }

            if (string.IsNullOrWhiteSpace(promptTemplate))
            {
                throw new ArgumentNullException(nameof(promptTemplate));
            }

            var prompt = promptTemplate.Replace(DataPlaceholder, data);

            // needed :(
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            var client = new OpenAIClient(apiKey, new OpenAIClientOptions());

            // TODO: pass options as params
            var chatCompletionsOptions = new ChatCompletionsOptions
            {
                Messages =
                {
                    new ChatMessage(ChatRole.User, prompt),
                },
                Temperature = 0
            };

            var response = await client.GetChatCompletionsAsync(
                deploymentOrModelName: "gpt-3.5-turbo",
                chatCompletionsOptions);

            var chatCompletions = response.Value;

            // we expect only one response
            return chatCompletions?.Choices?.FirstOrDefault()?.Message.Content;
        }
    }
}
