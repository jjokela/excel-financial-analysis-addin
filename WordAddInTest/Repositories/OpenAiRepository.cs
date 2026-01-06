using System;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Azure.AI.OpenAI;

namespace WordAddInTest.Repositories
{
    public class OpenAiRepository
    {
        private const string DataPlaceholder = "<<DATA>>";

        public static async Task<string> GetAnalysis(string data, string apiKey, string promptTemplate, string modelName)
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

            var chatCompletionsOptions = new ChatCompletionsOptions
            {
                Messages =
                {
                    new ChatMessage(ChatRole.User, prompt),
                },
                Temperature = 0,
                DeploymentName = modelName
            };

            var response = await client.GetChatCompletionsAsync(
                chatCompletionsOptions);

            var chatCompletions = response.Value;

            // we expect only one response
            return chatCompletions?.Choices?.FirstOrDefault()?.Message.Content;
        }
    }
}
