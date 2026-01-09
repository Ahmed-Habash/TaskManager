using System;
using System.Collections.Generic;
using System.Linq;
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
    public class GeminiService : IAIService
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly ILogger<GeminiService> _logger;

        private class GeminiPart { public string text { get; set; } }
        private class GeminiContent { public string role { get; set; } public List<GeminiPart> parts { get; set; } }
        private class GeminiRequest 
        { 
            public GeminiContent system_instruction { get; set; }
            public List<GeminiContent> contents { get; set; } 
        }

        public GeminiService(HttpClient httpClient, IConfiguration configuration, ILogger<GeminiService> logger)
        {
            _httpClient = httpClient;
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<string> GetAssistanceAsync(string userPrompt, IEnumerable<TaskItem> contextTasks, IEnumerable<ChatMessage>? history = null, TaskItem? focusedTask = null)
        {
            var apiKey = _configuration["Gemini:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return "Gemini API Key missing.";

            var systemPrompt = BuildSystemPrompt(contextTasks, focusedTask);
            var contents = new List<GeminiContent>();
            
            if (history != null)
            {
                foreach (var h in history)
                {
                    var role = h.Role == "user" ? "user" : "model";
                    if (contents.Count > 0 && contents.Last().role == role)
                    {
                        contents.Last().parts[0].text += "\n" + h.Content;
                    }
                    else
                    {
                        contents.Add(new GeminiContent { role = role, parts = new List<GeminiPart> { new GeminiPart { text = h.Content } } });
                    }
                }
            }

            if (contents.Count > 0 && contents.Last().role == "user")
            {
                contents.Last().parts[0].text += "\n" + userPrompt;
            }
            else
            {
                contents.Add(new GeminiContent { role = "user", parts = new List<GeminiPart> { new GeminiPart { text = userPrompt } } });
            }

            var requestBody = new GeminiRequest
            { 
                system_instruction = new GeminiContent { parts = new List<GeminiPart> { new GeminiPart { text = systemPrompt } } },
                contents = contents 
            };
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={apiKey}";

            try
            {
                var json = JsonSerializer.Serialize(requestBody);
                var response = await _httpClient.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json"));
                if (!response.IsSuccessStatusCode) 
                {
                    _logger.LogWarning($"Gemini GetAssistance failed ({response.StatusCode}).");
                    return $"Gemini Error: {response.StatusCode}";
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(jsonResponse);
                if (doc.RootElement.TryGetProperty("candidates", out var candidates) && candidates.GetArrayLength() > 0)
                {
                    return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString() ?? "No response.";
                }
            }
            catch { }
            return "Assistance failed.";
        }

        public async Task<string> DetermineActionAsync(string userPrompt, IEnumerable<ChatMessage> history)
        {
            var apiKey = _configuration["Gemini:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) 
            {
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_gemini_failure.txt", $"[{DateTime.Now}] Key Missing");
                return "[]";
            }

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
6. CHAT: message

Examples:
- User: 'my email is a@b.com and pass is 123' -> [{""action"": ""SET_EMAIL_CREDENTIALS"", ""parameters"": {""email"": ""a@b.com"", ""password"": ""123""}}, {""action"": ""CHAT"", ""parameters"": {""message"": ""I've remembered your email credentials for future use.""}}]
- User: 'create a ppt about lions' -> [{""action"": ""CREATE_DOCUMENT"", ""parameters"": {""file_name"": ""Lions.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro\nVisual Keywords: lion\n- Lions are predators...\nSlide 2: Habitat...""}}]
- User: 'make a 20 slide presentation about space' -> [{""action"": ""CREATE_DOCUMENT"", ""parameters"": {""file_name"": ""Space_Exploration.pptx"", ""type"": ""ppt"", ""content"": ""Slide 1: Intro...\nSlide 2: ...\n[ALL 20 SLIDES HERE]""}}]
- User: 'what is the weather' -> [{""action"": ""SEARCH_WEB"", ""parameters"": {""query"": ""current weather""}}]

Rules:
- DIRECT ACTION: If asked to 'make', 'create', 'write', 'generate', or 'convert to file', you MUST use CREATE_DOCUMENT. NEVER just chat or ask for clarification if the request is for a document.
- REFUSALS: If a request is prohibited (e.g., illegal, harmful, sensitive) or you cannot fulfill it, you MUST use the CHAT action to explain clearly to the user why you cannot fulfill the request. NEVER output an empty array '[]'.
- COMPLEX REQUESTS: If asked for many slides (e.g., 20), YOU MUST GENERATE ALL OF THEM in one single JSON action. Do not truncate.
- CHECK HISTORY for context/creds first.
- SEARCH_WEB: Use ONLY for direct questions or if the topics is truly unknown/real-time. For shopping queries, ALWAYS add 'buy' or 'specific product' to the query. For Temu, prefer 'site:temu.com ""product""' to find products. For regional queries (e.g. 'in Qatar'), prioritize local domains ending in the country code (e.g. '.qa') or include local currency in the query (e.g. 'QAR').
- GMAIL WARNING: If the user provides a @gmail.com address, you MUST warn them in a CHAT action that they likely need an 'App Password' for this to work, even if they give a regular password.
- REMEMBERING: If the user explicitly gives credentials, use SET_EMAIL_CREDENTIALS immediately. 
- PPT Images: Provide 2-3 HIGHLY CONCRETE and DESCRIPTIVE nouns for `image_keyword` and `Visual Keywords`. 
- NEVER use the word 'abstract'. Use 'photorealistic', 'macro photography', or 'detailed vector art' if you need a style.
- Example: 'Lion hunting in savanna' instead of 'Lion social'.
- NEVER use function-call syntax like ACTION(params). Output ONLY a STRICT JSON ARRAY.
- INTERNAL CONTEXT: You may see results ending in [INTERNAL_CONTEXT: ...]. This is for your memory only. NEVER repeat it.
- PRODUCT LINKS: For products, prioritize deep links to specific pages (e.g. including '/p/' or '/product/'). For Temu, URLs MUST contain '-p.html' and NEVER '-s.html'. SKIP links where the snippet mentions 'sold out', 'discontinued', or 'removed'.";

            var contents = new List<GeminiContent>();
            
            if (history != null)
            {
                var historyList = history.ToList();
                for (int i = 0; i < historyList.Count; i++)
                {
                    if (i < historyList.Count - 12) continue;
                    var h = historyList[i];

                    // Skip if this matches the current prompt to avoid double-entry
                    if (i == historyList.Count - 1 && h.Role == "user" && h.Content == userPrompt) continue;

                    int limit = (i == historyList.Count - 1) ? 2000 : 400;
                    var role = h.Role == "user" ? "user" : "model";
                    var text = h.Content?.Length > limit ? h.Content.Substring(0, limit) + "..." : h.Content;
                    
                    if (contents.Count > 0 && contents.Last().role == role)
                    {
                        contents.Last().parts[0].text += "\n" + text;
                    }
                    else
                    {
                        contents.Add(new GeminiContent { role = role, parts = new List<GeminiPart> { new GeminiPart { text = text } } });
                    }
                }
            }

            // Current prompt explicitly added at the end
            contents.Add(new GeminiContent { role = "user", parts = new List<GeminiPart> { new GeminiPart { text = userPrompt } } });

            var requestBody = new
            { 
                system_instruction = new { parts = new[] { new { text = systemPrompt } } },
                contents = contents,
                generation_config = new { response_mime_type = "application/json" }
            };
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={apiKey}";
            
            try
            {
                var json = JsonSerializer.Serialize(requestBody);
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_request_gemini.txt", $"[{DateTime.Now}] Request Body:\n{json}");
                var response = await _httpClient.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json"));
                
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning($"Gemini v1beta/system_instruction failed: {response.StatusCode}. Attempting legacy fallback...");
                    
                    // LEGACY FALLBACK: v1beta endpoint, no system_instruction, prompt merging
                    var legacyUrl = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={apiKey}";
                    var legacyPrompt = $"{systemPrompt}\n\nUSER REQUEST: {userPrompt}";
                    var legacyRequestBody = new { contents = new[] { new { role = "user", parts = new[] { new { text = legacyPrompt } } } } };
                    
                    var legacyResponse = await _httpClient.PostAsync(legacyUrl, new StringContent(JsonSerializer.Serialize(legacyRequestBody), Encoding.UTF8, "application/json"));
                    if (!legacyResponse.IsSuccessStatusCode)
                    {
                        var legacyError = await legacyResponse.Content.ReadAsStringAsync();
                        _logger.LogError($"Gemini Legacy Error: {legacyResponse.StatusCode} - {legacyError}");
                        await File.WriteAllTextAsync("wwwroot/exports/debug_ai_gemini_failure.txt", $"[{DateTime.Now}] Legacy HTTP Error: {legacyResponse.StatusCode} - {legacyError}");
                        return "[]";
                    }
                    response = legacyResponse; // Proceed with this response
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_response_gemini.txt", $"[{DateTime.Now}] Raw Response:\n{jsonResponse}");
                using var doc = JsonDocument.Parse(jsonResponse);
                
                if (doc.RootElement.TryGetProperty("candidates", out var candidates) && candidates.GetArrayLength() > 0)
                {
                    var content = candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
                    if (string.IsNullOrEmpty(content)) 
                    {
                        await File.WriteAllTextAsync("wwwroot/exports/debug_ai_gemini_failure.txt", $"[{DateTime.Now}] Empty text in response");
                        return "[]";
                    }

                    // Sanitize markdown fences and surrounding text
                    content = content.Trim();
                    // Improved regex to also match empty array []
                    var jsonMatch = System.Text.RegularExpressions.Regex.Match(content, @"\[\s*(\{.*\}|)\s*\]", System.Text.RegularExpressions.RegexOptions.Singleline);
                    if (jsonMatch.Success)
                    {
                        content = jsonMatch.Value;
                    }
                    else 
                    {
                        // RECOVERY: If AI returned function-call syntax instead of JSON
                        if (content.Contains("CREATE_DOCUMENT", StringComparison.OrdinalIgnoreCase))
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(content, @"CREATE_DOCUMENT\(\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\"",\s*\""(.*?)\""\)", System.Text.RegularExpressions.RegexOptions.Singleline);
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
                        
                        if (content.Contains("CHAT", StringComparison.OrdinalIgnoreCase))
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(content, @"CHAT\(\""(.*?)\""\)", System.Text.RegularExpressions.RegexOptions.Singleline);
                            if (match.Success)
                            {
                                return JsonSerializer.Serialize(new[] { new { action = "CHAT", parameters = new { message = match.Groups[1].Value }}});
                            }
                        }
                    }

                    try 
                    { 
                        var docJson = JsonDocument.Parse(content); 
                        if (docJson.RootElement.ValueKind == JsonValueKind.Array && docJson.RootElement.GetArrayLength() == 0)
                        {
                            await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_empty_gemini.txt", $"[{DateTime.Now}] Gemini returned empty array.\nPrompt: {userPrompt}\nFull Raw Content:\n{jsonResponse}");
                        }
                        return content; 
                    } 
                    catch 
                    { 
                        // RECOVERY 2: If it's just raw text, treat it as a CHAT action
                        if (!content.StartsWith("[") && !content.StartsWith("{"))
                        {
                            return JsonSerializer.Serialize(new[] { new { action = "CHAT", parameters = new { message = content }}});
                        }

                        await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_error_gemini.txt", $"[{DateTime.Now}] Gemini Invalid JSON:\n{content}\n\nFull Raw:\n{jsonResponse}");
                        return "[]"; 
                    }
                }
                
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_empty_gemini.txt", $"[{DateTime.Now}] Gemini No candidates or empty text.\nPrompt: {userPrompt}\nFull Response:\n{jsonResponse}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini Action Failed");
                await File.WriteAllTextAsync("wwwroot/exports/debug_ai_raw_exception_gemini.txt", $"[{DateTime.Now}] Exception: {ex.Message}\nPrompt: {userPrompt}");
            }
            return "[]";
        }

        public async Task<string> SummarizeSearchResultsAsync(string query, string rawResults)
        {
            var apiKey = _configuration["Gemini:ApiKey"];
            if (string.IsNullOrEmpty(apiKey)) return rawResults;

            var requestBody = new GeminiRequest
            {
                system_instruction = new GeminiContent { parts = new List<GeminiPart> { new GeminiPart { text = "You are a specialized search analyzer. Provide a clean, bulleted list of 3-5 distinct best options. Do NOT just pick one.\n\nYOUR STRICT RULES:\n1. RELEVANCE: You MUST ONLY include products that match the user's intent. If the raw results contain irrelevant items (e.g. underwear when searching for microphones), IGNORE THEM COMPLETELY.\n2. DIRECT PRODUCT LINKS: You MUST provide the deep link to a SPECIFIC product page. NEVER link to a homepage, category page, or 'Collection'. A valid product link usually contains '/p/', '/product/', '/gp/', or a unique SKU in the URL. For Temu, specific product pages MUST end in '-p.html' (NOT '-s.html'). If a search result only leads to a category or search result page, SKIP IT.\n3. PRICE IS MANDATORY: You MUST find the price for that specific product. If no price is found, use 'Price: N/A' or 'See site'. NEVER mention 'placeholder', 'missing source data', or 'debug info' in your response. Ensure you don't hallucinate reasoning about missing data. For regional searches, prioritize results in the local currency.\n4. NO DEAD LINKS: If a search result snippet or title mentions 'sold out', 'out of stock', 'discontinued', 'removed', or leads to a dead store (e.g. non-local stores for regional queries), SKIP IT ENTIRELY.\n5. SPECIFIC MODELS: Identify exact models (e.g., 'Boya BY-M1') instead of generic terms or categories. If a title is messy, clean it up for the bullet point.\n6. DIVERSITY: Provide 3-5 distinct specific products. Do NOT stop at one.\n7. FORMAT: - [Product Name](Deep Link) - Price - Website\n8. CLICKABLE LINKS: Use standard Markdown: [Name](URL).\n9. NO INTRO: Start your response IMMEDIATELY with the first bullet point. Do NOT include any 'Here are' or 'I found' text. Ensure each item is on a new line." } } },
                contents = new List<GeminiContent> { 
                    new GeminiContent { role = "user", parts = new List<GeminiPart> { new GeminiPart { text = $"Query: {query}\n\nResults:\n{rawResults}" } } } 
                }
            };
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={apiKey}";

            try
            {
                var response = await _httpClient.PostAsync(url, new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json"));
                if (!response.IsSuccessStatusCode) 
                {
                    _logger.LogWarning($"Gemini Summarization failed ({response.StatusCode}).");
                    return rawResults;
                }

                var jsonResponse = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(jsonResponse);
                if (doc.RootElement.TryGetProperty("candidates", out var candidates) && candidates.GetArrayLength() > 0)
                {
                    var summary = candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString() ?? rawResults;
                    await File.AppendAllTextAsync("wwwroot/exports/debug_ai_summarization_results_gemini.txt", $"[{DateTime.Now}] Query: {query}\nSummary: {summary}\n\n");
                    return summary;
                }
            }
            catch { }
            return rawResults;
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
            else
            {
                sb.AppendLine("\n>>> The user currently has no tasks.");
            }
            return sb.ToString();
        }
    }
}
