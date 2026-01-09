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

namespace TaskManager.Services
{
    public class OpenAIService : IAIService
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly ILogger<OpenAIService> _logger;

        public OpenAIService(HttpClient httpClient, IConfiguration configuration, ILogger<OpenAIService> logger)
        {
            _httpClient = httpClient;
            _configuration = configuration;
            _logger = logger;
            
            _httpClient.BaseAddress = new Uri("https://api.openai.com/v1/");
        }

        public async Task<string> GetAssistanceAsync(string userPrompt, IEnumerable<TaskItem> contextTasks, IEnumerable<ChatMessage>? history = null, TaskItem? focusedTask = null)
        {
            var apiKey = _configuration["OpenAI:ApiKey"];
            if (string.IsNullOrEmpty(apiKey))
            {
                return "OpenAI API Key is missing. Please configure it in appsettings.json.";
            }

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var systemPrompt = BuildSystemPrompt(contextTasks, focusedTask);

            // Construct messages list starting with system prompt
            var messages = new List<object>
            {
                new { role = "system", content = systemPrompt }
            };

            // Add history if present
            if (history != null)
            {
                foreach (var msg in history)
                {
                    messages.Add(new { role = msg.Role, content = msg.Content });
                }
            }

            // Add current user prompt
            messages.Add(new { role = "user", content = userPrompt });

            var requestBody = new
            {
                model = "gpt-4o-mini", // Using a cost-effective model
                messages = messages,
                max_tokens = 4000
            };

            var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");

            try
            {
                var response = await CallPostAsyncWithRetry("chat/completions", content);
                
                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogWarning($"OpenAI GetAssistance failed ({response.StatusCode}).");
                    return $"OpenAI Error: {response.StatusCode}";
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                var result = JsonSerializer.Deserialize<JsonElement>(jsonResponse);
                
                // Navigate: choices[0].message.content
                if (result.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                {
                    var messageObj = choices[0].GetProperty("message");
                    return messageObj.GetProperty("content").GetString() ?? "No response generated.";
                }
                
                return "The AI returned an empty response.";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception calls OpenAI API");
                return "An internal error occurred while contacting the AI service.";
            }
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
                // Limit the context to avoid hitting token limits if there are huge numbers of tasks
                // Just taking top 50 mostly relevant ones
                var relevantTasks = tasks.Where(t => !t.IsCompleted).OrderBy(t => t.DueDate).Take(20)
                                    .Concat(tasks.Where(t => t.IsCompleted).OrderByDescending(t => t.CreatedAt).Take(5));

                sb.AppendLine("\n>>> USER'S TASK LIST SUMMARY:");
                foreach (var task in relevantTasks)
                {
                    sb.AppendLine($"- [{task.Id}] {task.Title} (Due: {(task.DueDate.HasValue ? task.DueDate.Value.ToShortDateString() : "No Date")}, Pri: {task.Priority}, Status: {(task.IsCompleted ? "Done" : "Todo")})");
                }
            }
            else
            {
                sb.AppendLine("\n>>> The user currently has no tasks.");
            }

            return sb.ToString();
        }

        public async Task<string> DetermineActionAsync(string userPrompt, IEnumerable<ChatMessage> history)
        {
            var apiKey = _configuration["OpenAI:ApiKey"];
            if (string.IsNullOrEmpty(apiKey))
            {
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_openai_failure.txt", $"[{DateTime.Now}] Key Missing");
                return JsonSerializer.Serialize(new[] { new { action = "error", parameters = new { message = "OpenAI API Key is missing." } } });
            }

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var systemPrompt = $@"You are an autonomous agent. Current Date: {DateTime.Now:yyyy-MM-dd}. Output a STRICT JSON ARRAY of action objects. 
No text outside the array. No markdown fences.
Example: [{{""action"": ""CHAT"", ""parameters"": {{""message"": ""Hello""}}}}]

Actions:
1. ASK_USER: question
2. SEARCH_WEB: query
3. SEND_EMAIL: to, subject, body, attachment_path?, smtp_user?, smtp_password? (Use system creds if user/pass omitted)
4. SET_EMAIL_CREDENTIALS: email, password (Use to REMEMBER credentials for future use)
5. CREATE_DOCUMENT: file_name, content, type, primary_color?, secondary_color?, image_keyword?
   - type='word'(default) or 'ppt'. 
   - PPT Rules: You MUST use format 'Slide X: Title\nVisual Keywords: k1, k2\n- Fact 1\n- Fact 2\n- Fact 3'. Provide 5+ slides (or the exact number if requested, e.g., 20) with detailed facts and full sentences. Match colors to topic. Minimum 3 bullet points per slide.
6. CREATE_SPREADSHEET: file_name, data
   - data: A JSON Array of Objects (keys=headers) OR Array of Arrays (rows).
   - Example: {{""file_name"":""users.xlsx"", ""data"":[{{""Name"":""Alice"", ""Age"":30}}, {{""Name"":""Bob"", ""Age"":25}}]}}
7. CREATE_CHART: file_name, title, chart_type, labels, values
   - chart_type: 'bar', 'pie', 'line', 'doughnut'.
   - labels: Array of strings (e.g. [""Q1"", ""Q2""]).
   - values: Array of numbers (e.g. [10, 20]).
8. CHAT: message

Examples:
- User: 'my email is a@b.com and pass is 123' -> [{{""action"": ""SET_EMAIL_CREDENTIALS"", ""parameters"": {{""email"": ""a@b.com"", ""password"": ""123""}}}}, {{""action"": ""CHAT"", ""parameters"": {{""message"": ""I've remembered your email credentials for future use.""}}}}]
- User: 'email hello to test@test.com' -> [{{""action"": ""SEND_EMAIL"", ""parameters"": {{""to"": ""test@test.com"", ""subject"": ""Hello"", ""body"": ""This is a test email.""}}}}]
- User: 'create a ppt about lions' -> [{{""action"": ""CREATE_DOCUMENT"", ""parameters"": {{""file_name"": ""Lions.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro\nVisual Keywords: lion\n- Lions are predators...\nSlide 2: Habitat...""}}}}]
- User: 'make a 20 slide presentation about space' -> [{{""action"": ""CREATE_DOCUMENT"", ""parameters"": {{""file_name"": ""Space_Exploration.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro...\nSlide 2: ...\n[ALL 20 SLIDES HERE]""}}}}]
- User: 'plot a chart of sales' -> [{{""action"": ""CREATE_CHART"", ""parameters"": {{""file_name"": ""sales.png"", ""title"": ""Sales Data"", ""chart_type"": ""bar"", ""labels"": [""Jan"", ""Feb""], ""values"": [100, 150]}}}}]
- User: 'create a sheet of users' -> [{{""action"": ""CREATE_SPREADSHEET"", ""parameters"": {{""file_name"": ""users.xlsx"", ""data"": [{{""id"":1, ""name"":""john""}}]}}}}]
- User: 'what is the weather' -> [{{""action"": ""SEARCH_WEB"", ""parameters"": {{""query"": ""current weather""}}}}]

Rules:
- DIRECT ACTION: If asked to 'make', 'create', 'write', 'generate', or 'convert to file', you MUST use the appropriate CREATE_ tool. NEVER just chat or ask for clarification if the request is for a file.
- REFUSALS: If a request is prohibited (e.g., illegal, harmful, sensitive) or you cannot fulfill it, you MUST use the CHAT action to explain clearly to the user why you cannot fulfill the request. NEVER output an empty array '[]'.
- COMPLEX REQUESTS: If asked for many slides (e.g., 20), YOU MUST GENERATE ALL OF THEM in one single JSON action. Do not truncate.
- CHECK HISTORY for context/creds first.
- SEARCH_WEB: Use ONLY for direct questions or if the topics is truly unknown/real-time (like current events, future movies, weather). For shopping queries, ALWAYS add 'buy' or 'specific product' to the query. For Temu, prefer 'site:temu.com ""product""' to find products. For regional queries (e.g. 'in Qatar'), prioritize local domains ending in the country code (e.g. '.qa') or include local currency in the query (e.g. 'QAR').
- GMAIL WARNING: If the user provides a @gmail.com address, you MUST warn them in a CHAT action that they likely need an 'App Password' for this to work, even if they give a regular password.
- REMEMBERING: If the user explicitly gives credentials, use SET_EMAIL_CREDENTIALS immediately. 
- PPT Images: Provide 2-3 HIGHLY CONCRETE and DESCRIPTIVE nouns for `image_keyword` and `Visual Keywords`. 
- NEVER use the word 'abstract'. Use 'photorealistic', 'macro photography', or 'detailed vector art' if you need a style.
- Example: 'Lion hunting in savanna' instead of 'Lion social'.
- NEVER use function-call syntax like ACTION(params). Output ONLY a STRICT JSON ARRAY.
- INTERNAL CONTEXT: You may see results ending in [INTERNAL_CONTEXT: ...]. This is for your memory only. NEVER repeat it.
- PRODUCT LINKS: For products, prioritize deep links to specific pages (e.g. including '/p/' or '/product/'). For Temu, URLs MUST contain '-p.html' and NEVER '-s.html'. SKIP links where the snippet mentions 'sold out', 'discontinued', or 'removed'." ;

            var messages = new List<object>
            {
                new { role = "system", content = systemPrompt }
            };

            if (history != null) 
            {
                var historyList = history.ToList();
                for (int i = 0; i < historyList.Count; i++)
                {
                    if (i < historyList.Count - 12) continue; // Only take last 12

                    var h = historyList[i];
                    // Skip if this matches the current prompt to avoid double-entry
                    if (i == historyList.Count - 1 && h.Role == "user" && h.Content == userPrompt) continue;

                    int limit = (i == historyList.Count - 1) ? 2000 : 400;
                    var contentToLink = h.Content?.Length > limit ? h.Content.Substring(0, limit) + "..." : h.Content;
                    messages.Add(new { role = h.Role, content = contentToLink });
                }
            }

            // Current prompt (always add it explicitly as the final instruction)
            messages.Add(new { role = "user", content = userPrompt });

            var requestBody = new
            {
                model = "gpt-4o-mini",
                messages = messages,
                max_tokens = 8000,
                temperature = 0.3
            };

            var requestJson = JsonSerializer.Serialize(requestBody);
            await File.WriteAllTextAsync("wwwroot/exports/debug_ai_request_openai.txt", $"[{DateTime.Now}] Request Body:\n{requestJson}");
            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");

            try
            {
                var response = await CallPostAsyncWithRetry("chat/completions", content);
                
                if (!response.IsSuccessStatusCode) 
                {
                    _logger.LogWarning($"OpenAI DetermineAction failed ({response.StatusCode}).");
                    await File.WriteAllTextAsync("wwwroot/exports/debug_ai_openai_failure.txt", $"[{DateTime.Now}] HTTP Error: {response.StatusCode}");
                    return "[]";
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_response_openai.txt", $"[{DateTime.Now}] Raw Response:\n{jsonResponse}");
                
                try 
                {
                    using var result = JsonDocument.Parse(jsonResponse);
                    if (result.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                    {
                        var messageContent = choices[0].GetProperty("message").GetProperty("content").GetString();
                        
                        // Sanitize markdown fences and surrounding text
                        messageContent = messageContent.Trim();
                        
                        int firstBracket = messageContent.IndexOf('[');
                        int lastBracket = messageContent.LastIndexOf(']');
                        if (firstBracket >= 0 && lastBracket > firstBracket)
                        {
                            messageContent = messageContent.Substring(firstBracket, lastBracket - firstBracket + 1);
                        }
                        else 
                        {
                            // RECOVERY: If AI returned function-call syntax instead of JSON
                            if (messageContent.Contains("CREATE_DOCUMENT", StringComparison.OrdinalIgnoreCase))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(messageContent, @"CREATE_DOCUMENT\(\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\""\)", System.Text.RegularExpressions.RegexOptions.Singleline);
                                if (match.Success)
                                {
                                    return JsonSerializer.Serialize(new[] { new { action = "CREATE_DOCUMENT", parameters = new { 
                                        file_name = match.Groups[1].Value,
                                        content = match.Groups[2].Value,
                                        type = match.Groups[3].Value,
                                        primary_color = match.Groups[4].Value,
                                        secondary_color = match.Groups[5].Value,
                                        image_keyword = match.Groups[6].Value
                                    }}});
                                }
                            }
                            
                            if (messageContent.Contains("CHAT", StringComparison.OrdinalIgnoreCase))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(messageContent, @"CHAT\(\""(.*?)\""\)", System.Text.RegularExpressions.RegexOptions.Singleline);
                                if (match.Success)
                                {
                                    return JsonSerializer.Serialize(new[] { new { action = "CHAT", parameters = new { message = match.Groups[1].Value }}});
                                }
                            }
                        }

                        // Validate JSON
                        try 
                        {
                            var doc = JsonDocument.Parse(messageContent); // Check if valid
                            if (doc.RootElement.ValueKind == JsonValueKind.Array && doc.RootElement.GetArrayLength() == 0)
                            {
                                // If AI returned an empty array, it means it didn't identify an action.
                                // Log this specifically.
                                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_empty.txt", $"[{DateTime.Now}] AI returned empty array.\nPrompt: {userPrompt}\nMessage Content: {messageContent}");
                            }
                            return messageContent;
                        }
                        catch (Exception ex)
                        {
                            // RECOVERY 2: If it's just raw text, treat it as a CHAT action
                            if (!messageContent.StartsWith("[") && !messageContent.StartsWith("{"))
                            {
                                return JsonSerializer.Serialize(new[] { new { action = "CHAT", parameters = new { message = messageContent }}});
                            }

                            _logger.LogWarning($"Invalid JSON from OpenAI: {messageContent.Substring(0, Math.Min(100, messageContent.Length))}. Returning empty array.");
                            await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_error.txt", $"[{DateTime.Now}] Parse Error: {ex.Message}\nContent that failed:\n{messageContent}\n\nFull Raw Response:\n{jsonResponse}");
                            return "[]";
                        }
                    }
                    
                    await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_empty.txt", $"[{DateTime.Now}] No choices or empty message.\nPrompt: {userPrompt}\nFull Response:\n{jsonResponse}");
                    return "[]";
                }
                catch(Exception ex)
                {
                    await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_exception.txt", $"[{DateTime.Now}] Exception: {ex.Message}\nRaw Response:\n{jsonResponse}");
                     return JsonSerializer.Serialize(new[] { new { action = "error", parameters = new { message = "Critical parsing error." } } });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception calls OpenAI API for Action");
                return JsonSerializer.Serialize(new[] { new { action = "error", parameters = new { message = ex.Message } } });
            }
        }
        public async Task<string> SimulateSearchAsync(string query)
        {
            var apiKey = _configuration["OpenAI:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return "Unable to search: API Key missing.";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var messages = new List<object>
            {
                new { role = "system", content = "You are a helpful search engine simulator. The user asked a query. Provide a detailed, helpful summary answer based on your knowledge. If the query is about prices or specific availability, give typical estimates and likely stores. Be concise but informative." },
                new { role = "user", content = query }
            };

            var requestBody = new
            {
                model = "gpt-4o-mini",
                messages = messages,
                max_tokens = 300
            };

            var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");

            try
            {
                var response = await CallPostAsyncWithRetry("chat/completions", content);
                if (!response.IsSuccessStatusCode) return "Search failed: OpenAI Error.";

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var result = JsonDocument.Parse(jsonResponse);
                
                if (result.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                {
                   return choices[0].GetProperty("message").GetProperty("content").GetString() ?? "No results found.";
                }
                return "No results found."; // Added return for success path without choices
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Search Simulation Failed");
                return "Search failed due to an error.";
            }
        }
        public async Task<string> SummarizeSearchResultsAsync(string query, string rawResults)
        {
            var apiKey = _configuration["OpenAI:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return rawResults;

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var messages = new List<object>
            {
                new { role = "system", content = "You are a specialized search analyzer. Provide a clean, bulleted list of 3-5 distinct best options. Do NOT just pick one.\n\nYOUR STRICT RULES:\n1. RELEVANCE: You MUST ONLY include products that match the user's intent. If the raw results contain irrelevant items (e.g. underwear when searching for microphones), IGNORE THEM COMPLETELY.\n2. DIRECT PRODUCT LINKS: You MUST provide the deep link to a SPECIFIC product page. NEVER link to a homepage, category page, or 'Collection'. A valid product link usually contains '/p/', '/product/', '/gp/', or a unique SKU in the URL. For Temu, specific product pages MUST end in '-p.html' (NOT '-s.html'). If a search result only leads to a category or search result page, SKIP IT.\n3. PRICE IS MANDATORY: You MUST find the price. If multiple prices are in a snippet, pick the one for the specific product. If no price is found, use 'Price: N/A' or 'See site'. NEVER mention 'placeholder', 'missing source data', or 'debug info' in your response. For regional searches, prioritize results in the local currency.\n4. NO DEAD LINKS: If a search result snippet or title mentions 'sold out', 'out of stock', 'discontinued', 'removed', or leads to a dead store (e.g. non-local stores for regional queries), SKIP IT ENTIRELY.\n5. SPECIFIC MODELS: Identify exact models (e.g., 'Boya BY-M1') instead of generic terms. If a title is messy, clean it up for the bullet point.\n6. DIVERSITY: Provide 3-5 distinct specific products. Do NOT stop at one.\n7. FORMAT: - [Product Name](Deep Link) - Price - Website\n8. CLICKABLE LINKS: Use standard Markdown: [Name](URL).\n9. NO INTRO: Start your response IMMEDIATELY with the first bullet point. Do NOT include any 'Here are' or 'I found' text." },


                new { role = "user", content = $"Query: {query}\n\nRaw Search Results:\n{rawResults}" }
            };

            var requestBody = new
            {
                model = "gpt-4o-mini",
                messages = messages,
                max_tokens = 600,
                temperature = 0.3
            };

            var requestJson = JsonSerializer.Serialize(requestBody);
            try { await File.AppendAllTextAsync("wwwroot/exports/debug_ai_summarization_requests.txt", $"[{DateTime.Now}] Query: {query}\nJSON: {requestJson}\n\n"); } catch {}
            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");

            try
            {
                var response = await CallPostAsyncWithRetry("chat/completions", content);
                if (!response.IsSuccessStatusCode) 
                {
                    _logger.LogWarning($"OpenAI Summarization failed ({response.StatusCode}).");
                    return rawResults;
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var result = JsonDocument.Parse(jsonResponse);
                
                if (result.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                {
                   var summary = choices[0].GetProperty("message").GetProperty("content").GetString() ?? rawResults;
                   await File.AppendAllTextAsync("wwwroot/exports/debug_ai_summarization_results.txt", $"[{DateTime.Now}] Query: {query}\nSummary: {summary}\n\n");
                   return summary;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Search Summarization Failed");
                return "[AI Summarization Error - Showing Raw Search Results]:\n\n" + rawResults;
            }
            return rawResults;
        }

        private async Task<HttpResponseMessage> CallPostAsyncWithRetry(string url, HttpContent content)
        {
            // The user wants an immediate switch to Gemini on quota (429), so we send it back immediately without retrying.
            var response = await _httpClient.PostAsync(url, content);
            return response;
        }
    }
}
