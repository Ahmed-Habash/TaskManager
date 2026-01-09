using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using TaskManager.Models;
using System.IO;

namespace TaskManager.Services
{
    public class OpenRouterService : IAIService
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly ILogger<OpenRouterService> _logger;

        public OpenRouterService(HttpClient httpClient, IConfiguration configuration, ILogger<OpenRouterService> logger)
        {
            _httpClient = httpClient;
            _configuration = configuration;
            _logger = logger;
            
            _httpClient.BaseAddress = new Uri("https://openrouter.ai/api/v1/");
        }

        public async Task<string> GetAssistanceAsync(string userPrompt, IEnumerable<TaskItem> contextTasks, IEnumerable<ChatMessage>? history = null, TaskItem? focusedTask = null)
        {
            var apiKey = _configuration["OpenRouter:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return "OpenRouter API Key missing.";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            _httpClient.DefaultRequestHeaders.Add("HTTP-Referer", "https://taskmanager.local");
            _httpClient.DefaultRequestHeaders.Add("X-Title", "TaskManager AI");

            var systemPrompt = BuildSystemPrompt(contextTasks, focusedTask);
            var messages = new List<object> { new { role = "system", content = systemPrompt } };

            if (history != null)
            {
                foreach (var msg in history) messages.Add(new { role = msg.Role, content = msg.Content });
            }

            messages.Add(new { role = "user", content = userPrompt });

            var requestBody = new
            {
                model = "openai/gpt-4o-mini", // Reliable fallback
                messages = messages,
                max_tokens = 4000
            };

            try
            {
                var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync("chat/completions", content);
                
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger.LogError($"OpenRouter GetAssistance failed: {response.StatusCode} - {error}");
                    return $"OpenRouter Error: {response.StatusCode}";
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(jsonResponse);
                return doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? "No response.";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenRouter Assistance Failed");
                return "OpenRouter service error.";
            }
        }

        public async Task<string> DetermineActionAsync(string userPrompt, IEnumerable<ChatMessage> history)
        {
            var apiKey = _configuration["OpenRouter:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return "[]";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            _httpClient.DefaultRequestHeaders.Add("HTTP-Referer", "https://taskmanager.local");
            _httpClient.DefaultRequestHeaders.Add("X-Title", "TaskManager AI");

            var systemPrompt = @"You are an autonomous agent. Output a STRICT JSON ARRAY of action objects. 
No text outside the array. No markdown fences.
Example: [{""action"": ""CHAT"", ""parameters"": {""message"": ""Hello""}}]

Actions:
1. ASK_USER: question
2. SEARCH_WEB: query
3. SEND_EMAIL: to, subject, body, attachment_path?, smtp_user?, smtp_password? (Use system creds if user/pass omitted)
4. SET_EMAIL_CREDENTIALS: email, password (Use to REMEMBER credentials for future use)
5. CREATE_DOCUMENT: file_name, content, type, primary_color?, secondary_color?, image_keyword?
   - type='word'(default) or 'ppt'. 
   - PPT Rules: You MUST use format 'Slide X: Title\nVisual Keywords: k1, k2\n- Fact 1\n- Fact 2\n- Fact 3'. Provide 5+ slides (or the exact number if requested, e.g., 20) with detailed facts and full sentences. Match colors to topic. Minimum 3 bullet points per slide.
7. CREATE_SPREADSHEET: file_name, data
   - data: A JSON array of arrays representing rows and cells. Example: [[""Header1"",""Header2""],[""Row1Col1"",""Row1Col2""]].
8. CREATE_CHART: file_name, title, chart_type, labels, values
   - chart_type: 'bar', 'line', or 'pie'.
   - labels: Pipe-separated string of categories (e.g., 'Q1|Q2|Q3').
   - values: Pipe-separated string of numbers (e.g., '10|20|30').
9. CHAT: message

Examples:
- User: 'my email is a@b.com and pass is 123' -> [{""action"": ""SET_EMAIL_CREDENTIALS"", ""parameters"": {""email"": ""a@b.com"", ""password"": ""123""}}, {""action"": ""CHAT"", ""parameters"": {""message"": ""I've remembered your email credentials for future use.""}}]
- User: 'create a ppt about lions' -> [{""action"": ""CREATE_DOCUMENT"", ""parameters"": {""file_name"": ""Lions.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro\nVisual Keywords: lion\n- Lions are predators...\nSlide 2: Habitat...""}}]
- User: 'make a 20 slide presentation about space' -> [{""action"": ""CREATE_DOCUMENT"", ""parameters"": {""file_name"": ""Space_Exploration.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro...\nSlide 2: ...\n[ALL 20 SLIDES HERE]""}}]
- User: 'what is the weather' -> [{""action"": ""SEARCH_WEB"", ""parameters"": {""query"": ""current weather""}}]
- User: 'plot a graph of green birb coin' -> [{""action"": ""CREATE_CHART"", ""parameters"": {""file_name"": ""green_birb_coin_value.png"", ""title"": ""Green Birb Coin Value"", ""chart_type"": ""line"", ""labels"": ""Jan|Feb|Mar|Apr|May"", ""values"": ""0.001|0.002|0.0015|0.003|0.005""}}]
- User: 'chart the population of top 3 countries' -> [{""action"": ""CREATE_CHART"", ""parameters"": {""file_name"": ""population.png"", ""title"": ""Top 3 Countries Population"", ""chart_type"": ""bar"", ""labels"": ""India|China|USA"", ""values"": ""1428000000|1425000000|336000000""}}]
- User: 'search for ps5 prices and show me the best one' -> [{""action"": ""SEARCH_WEB"", ""parameters"": {""query"": ""ps5 prices best deal""}}]

Rules:
- PRIORITY 1: PLOTTING OVERRIDES SEARCHING. If the user request contains 'plot', 'graph', 'chart', or 'create a sheet':
  - YOU MUST GENERATE THE ARTIFACT IMMEDIATELY (CREATE_CHART / CREATE_SPREADSHEET).
  - YOU MUST NOT SEARCH, even if you don't have real-time data.
  - USE INTERNAL KNOWLEDGE Or ESTIMATATIONS. usage of 'Green Birb Coin' or 'Qatar House Prices' IS NOT AN EXCUSE TO SEARCH.
  - ONLY search if the user explicitly adds '...and search for data' or 'find data then plot'.
- DIRECT ACTION: If asked to 'make', 'create', 'write', 'generate', 'plot', 'graph', 'chart', 'visualize', or 'convert to file', you MUST use the appropriate CREATE_ action. NEVER just chat if the request is for a document, sheet, or chart.
- REFUSALS: If a request is prohibited (e.g., illegal, harmful, sensitive) or you cannot fulfill it, you MUST use the CHAT action to explain clearly to the user why you cannot fulfill the request. NEVER output an empty array '[]'.
- COMPLEX REQUESTS: If asked for many slides (e.g., 20), YOU MUST GENERATE ALL OF THEM in one single JSON action. Do not truncate.
- CHECK HISTORY for context/creds first.
- COMPOUND COMMANDS: You MUST BREAK DOWN complex requests into multiple independent actions in the single JSON array.
  - Example: 'Search for X and compare Y to Z' -> [{ ""action"": ""SEARCH_WEB"", ""parameters"": { ""query"": ""Search for X"" } }, { ""action"": ""CHAT"", ""parameters"": { ""message"": ""Here is the comparison of Y and Z: ..."" } }]
  - CRITICAL: If the user says 'search' or 'find', you MUST generate a SEARCH_WEB action. Do not skip it just because you can answer the other part of the request.
- SEARCH_WEB: Use ONLY if the user EXPLICITLY asks to 'search', 'find', 'price check', 'look up', 'where to buy', or 'find a store'.
- CHART/SHEET GENERATION: If the user asks to 'plot', 'graph', 'chart', or 'create a sheet' of something (even if specific like 'Green Birb Coin' or 'Pizza Prices'), DO NOT SEARCH. Use your internal knowledge to GENERATE THE DATA IMMEDIATELY and call CREATE_CHART or CREATE_SPREADSHEET.
  - If the item is fictional or niche (e.g. 'Green Birb Coin'), INVENT PLAUSIBLE DATA (estimates) and proceed directly to creation.
  - DO NOT output a search action unless the user's prompt starts with a question or 'search for'.
- GMAIL WARNING: If the user provides a @gmail.com address, you MUST warn them in a CHAT action that they likely need an 'App Password' for this to work, even if they give a regular password.
- REMEMBERING: If the user explicitly gives credentials, use SET_EMAIL_CREDENTIALS immediately. 
- PPT Images: Provide 2-3 HIGHLY CONCRETE and DESCRIPTIVE nouns for `image_keyword` and `Visual Keywords`. 
- NEVER use the word 'abstract'. Use 'photorealistic', 'macro photography', or 'detailed vector art' if you need a style.
- Example: 'Lion hunting in savanna' instead of 'Lion social'.
- NEVER use function-call syntax like ACTION(params). Output ONLY a STRICT JSON ARRAY.
- PRODUCT LINKS: For products, prioritize deep links to specific pages (e.g. including '/p/', '/product/'). For Temu, URLs MUST contain '-p.html' and NEVER '-s.html'. SKIP links where the snippet mentions 'sold out', 'discontinued', or 'removed'." ;

            var messages = new List<object> { new { role = "system", content = systemPrompt } };
            if (history != null) 
            {
                foreach (var h in history.TakeLast(10)) messages.Add(new { role = h.Role, content = h.Content });
            }
            messages.Add(new { role = "user", content = userPrompt });

            var requestBody = new { model = "openai/gpt-4o-mini", messages = messages, temperature = 0.3 };

            try
            {
                var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync("chat/completions", content);
                
                if (!response.IsSuccessStatusCode) 
                {
                    var statusCode = response.StatusCode;
                    var error = await response.Content.ReadAsStringAsync();
                    _logger.LogError($"OpenRouter DetermineAction failed: {statusCode} - {error}");
                    
                    var msg = $"OpenRouter Failover Error: {statusCode}.";
                    return JsonSerializer.Serialize(new[] { new { action = "CHAT", parameters = new { message = msg } } });
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var result = JsonDocument.Parse(jsonResponse);
                var contentStr = result.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? "[]";
                
                contentStr = contentStr.Trim();
                int firstBracket = contentStr.IndexOf('[');
                int lastBracket = contentStr.LastIndexOf(']');
                if (firstBracket >= 0 && lastBracket > firstBracket)
                    return contentStr.Substring(firstBracket, lastBracket - firstBracket + 1);
                
                return contentStr;
            }
            catch { return "[]"; }
        }

        public async Task<string> SummarizeSearchResultsAsync(string query, string rawResults)
        {
            var apiKey = _configuration["OpenRouter:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return rawResults;

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var messages = new List<object>
            {
                new { role = "system", content = "You are a helpful search result summarizer. \n\nYOUR GOAL: Extract the most relevant answers for the user's query.\n\nRULES:\n1. IF SHOPPING/BUYING: Identify specific products, prices, and links. Filter out 'out of stock' items. Format as: - [Product Name](Link) - Price - Store (Brief details).\n2. IF FACTUAL/STATS (e.g. 'stats of X', 'height of Y'): Provide the direct answer, facts, or statistics found. Citation style: 'Fact... [Source](Link)'.\n3. IF NEWS/GENERAL: Summarize the key points found.\n4. NO FLUFF: Start directly with the results. Do not say 'Here is what I found'.\n5. DIVERSITY: If multiple good sources exist, list 3-5 distinct ones.\n6. PRECISION: If the user asks for a specific metric (e.g. 'price', 'population'), prioritize results that contain that number." },
                new { role = "user", content = $"Query: {query}\n\nRaw Search Results:\n{rawResults}" }
            };

            var requestBody = new { model = "openai/gpt-4o-mini", messages = messages, max_tokens = 600, temperature = 0.3 };

            try
            {
                var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync("chat/completions", content);
                
                if (!response.IsSuccessStatusCode)
                {
                   if (response.StatusCode == (HttpStatusCode)402)
                       return "[OpenRouter Failure: 402 Payment Required - Balance exhausted].\n\n" + rawResults;
                   return rawResults;
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var result = JsonDocument.Parse(jsonResponse);
                return result.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? rawResults;
            }
            catch { return rawResults; }
        }

        private string BuildSystemPrompt(IEnumerable<TaskItem> tasks, TaskItem? focusedTask)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a helpful and intelligent productivity assistant for a Task Manager application.");
            sb.AppendLine("You have access to the user's current tasks. Use this information to provide context-aware advice, summaries, or motivation.");
            sb.AppendLine("Keep your responses concise, friendly, and actionable. Format responses in Markdown.");

            if (focusedTask != null)
            {
                sb.AppendLine("\n>>> FOCUSED TASK (The user is asking specifically about this task):");
                sb.AppendLine($"- ID: {focusedTask.Id}");
                sb.AppendLine($"  Title: {focusedTask.Title}");
                sb.AppendLine($"  Description: {focusedTask.Description ?? "None"}");
                sb.AppendLine($"  Status: {(focusedTask.IsCompleted ? "Completed" : "Pending")}");
                sb.AppendLine($"  Priority: {focusedTask.Priority}");
                sb.AppendLine($"  Due: {(focusedTask.DueDate.HasValue ? focusedTask.DueDate.Value.ToString("yyyy-MM-dd") : "None")}");
            }

            if (tasks != null && tasks.Any())
            {
                var relevantTasks = tasks.Where(t => !t.IsCompleted).OrderBy(t => t.DueDate).Take(20)
                                    .Concat(tasks.Where(t => t.IsCompleted).OrderByDescending(t => t.CreatedAt).Take(5));

                sb.AppendLine("\n>>> USER'S TASK LIST SUMMARY:");
                foreach (var task in relevantTasks)
                {
                    sb.AppendLine($"- [{task.Id}] {task.Title} (Due: {(task.DueDate.HasValue ? task.DueDate.Value.ToShortDateString() : "No Date")}, Pri: {task.Priority}, Status: {(task.IsCompleted ? "Done" : "Todo")})");
                }
            }
            
            return sb.ToString();
        }
    }
}
