using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using TaskManager.Models;

namespace TaskManager.Services
{
    public class AIFalloverService : IAIService
    {
        private readonly OpenAIService _openAI;
        private readonly GeminiService _gemini;
        private readonly OpenRouterService _openRouter;
        private readonly IConfiguration _configuration;
        private readonly ILogger<AIFalloverService> _logger;

        public AIFalloverService(
            OpenAIService openAI, 
            GeminiService gemini, 
            OpenRouterService openRouter, 
            IConfiguration configuration,
            ILogger<AIFalloverService> logger)
        {
            _openAI = openAI;
            _gemini = gemini;
            _openRouter = openRouter;
            _configuration = configuration;
            _logger = logger;
        }

        private List<string> GetPriorityList()
        {
            var priority = _configuration.GetSection("AI:Priority").Get<string[]>();
            if (priority == null || !priority.Any())
            {
                return new List<string> { "OpenAI", "Gemini", "OpenRouter" };
            }
            return priority.ToList();
        }

        private IAIService? GetServiceByName(string hand)
        {
            return hand.ToLower() switch
            {
                "openai" => _openAI,
                "gemini" => _gemini,
                "openrouter" => _openRouter,
                _ => null
            };
        }

        public async Task<string> GetAssistanceAsync(string userPrompt, IEnumerable<TaskItem> contextTasks, IEnumerable<ChatMessage>? history = null, TaskItem? focusedTask = null)
        {
            var list = GetPriorityList();
            foreach (var name in list)
            {
                var service = GetServiceByName(name);
                if (service == null) continue;

                try
                {
                    var result = await service.GetAssistanceAsync(userPrompt, contextTasks, history, focusedTask);
                    if (!IsErrorResponse(result)) return result;
                    _logger.LogWarning($"{name} returned error response. Falling over...");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error calling {name}");
                }
            }
            return "All AI services failed to provide assistance.";
        }

        public async Task<string> DetermineActionAsync(string userPrompt, IEnumerable<ChatMessage> history)
        {
            var list = GetPriorityList();
            foreach (var name in list)
            {
                var service = GetServiceByName(name);
                if (service == null) continue;

                try
                {
                    var result = await service.DetermineActionAsync(userPrompt, history);
                    if (result != "[]" && !result.Contains("Failover Error")) return result;
                    _logger.LogWarning($"{name} failed to determine action. Falling over...");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error calling {name} for action determination");
                }
            }
            return "[]";
        }

        public async Task<string> SummarizeSearchResultsAsync(string query, string rawResults)
        {
            var list = GetPriorityList();
            foreach (var name in list)
            {
                var service = GetServiceByName(name);
                if (service == null) continue;

                try
                {
                    var result = await service.SummarizeSearchResultsAsync(query, rawResults);
                    if (result != rawResults && !result.Contains("Failure")) return result;
                }
                catch {}
            }
            return rawResults;
        }

        private bool IsErrorResponse(string response)
        {
            return response.Contains("Error:") || response.Contains("failed") || response.Contains("unavailable");
        }
    }
}
