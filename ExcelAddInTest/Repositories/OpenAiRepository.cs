using System;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Azure.AI.OpenAI;

namespace ExcelAddInTest.Repositories
{
    public class OpenAiRepository
    {
        public static async Task<string> GetAnalysis(string data, string apiKey)
        {
            if (string.IsNullOrEmpty(apiKey))
            {
                throw new ArgumentNullException(nameof(apiKey));
            }
            
            // needed :(
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // TODO: pass prompt as parameter, as well as model, temperature and other options
            var prompt =
                "You are an expert financial analyst." +
                "Analyze the following income statement data:" +
                $"{data}" +
                "What are some key insights and trends from this data?";

            var client = new OpenAIClient(apiKey, new OpenAIClientOptions());
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
