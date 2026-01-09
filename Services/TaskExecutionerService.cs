using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;
using System.Net;
using Microsoft.Extensions.Logging;

namespace TaskManager.Services
{
    public class TaskExecutionerService : ITaskExecutionerService
    {
        private readonly IWebHostEnvironment _environment;
        private readonly IAIService _aiService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<TaskExecutionerService> _logger;

        private static bool _googleQuotaExceeded = false;
        private static DateTime _lastQuotaReset = DateTime.MinValue;

        public TaskExecutionerService(IWebHostEnvironment environment, IAIService aiService, IConfiguration configuration, ILogger<TaskExecutionerService> logger)
        {
            _environment = environment;
            _aiService = aiService;
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<string> ExecuteAICommandAsync(string userPrompt, List<ChatMessage> history)
        {
            var jsonResponse = await _aiService.DetermineActionAsync(userPrompt, history);
            var sb = new StringBuilder();

            try 
            {
                using var doc = JsonDocument.Parse(jsonResponse);
                var root = doc.RootElement;
                
                if (root.ValueKind != JsonValueKind.Array)
                {
                    return "The AI returned an invalid response format.";
                }

                foreach (var actionElement in root.EnumerateArray())
                {
                    if (actionElement.TryGetProperty("action", out var actionProp))
                    {
                        var action = actionProp.GetString();
                        
                        JsonElement parameters = default;
                        if (actionElement.TryGetProperty("parameters", out var pProp))
                        {
                            parameters = pProp;
                        }
                        else
                        {
                            // Create an empty JsonElement if parameters are missing
                            using var docEmpty = JsonDocument.Parse("{}");
                            parameters = docEmpty.RootElement.Clone(); 
                        }

                        if (action == "ASK_USER")
                        {
                            return parameters.GetProperty("question").GetString() ?? "I need more information.";
                        }
                        else if (action == "SEARCH_WEB")
                        {
                            string query = "Generic Search";
                            if (parameters.TryGetProperty("query", out var qProp)) query = qProp.GetString() ?? "Generic Search";
                            
                            string searchResult;
                            try 
                            {
                                var googleApiKey = _configuration["Google:ApiKey"];
                                var googleCx = _configuration["Google:SearchEngineId"];
                                
                                if (!string.IsNullOrEmpty(googleApiKey) && !string.IsNullOrEmpty(googleCx) && googleCx != "YOUR_SEARCH_ENGINE_ID_HERE")
                                {
                                    using (var client = new System.Net.Http.HttpClient())
                                    {
                                        var url = $"https://www.googleapis.com/customsearch/v1?key={googleApiKey}&cx={googleCx}&q={Uri.EscapeDataString(query)}";
                                        var response = await client.GetAsync(url);
                                        if (response.IsSuccessStatusCode)
                                        {
                                            var searchJson = await response.Content.ReadAsStringAsync();
                                            using var searchDoc = JsonDocument.Parse(searchJson);
                                            var items = searchDoc.RootElement.TryGetProperty("items", out var itemsProp) ? itemsProp : default;
                                            
                                            if (items.ValueKind == JsonValueKind.Array)
                                            {
                                                var results = new StringBuilder();
                                                int processedCount = 0;
                                                int foundCount = 0;
                                                foreach (var item in items.EnumerateArray())
                                                {
                                                    if (processedCount >= 20) break; // Limit raw pool
                                                    var title = item.TryGetProperty("title", out var tProp) ? tProp.GetString() : "No Title";
                                                    var snippet = item.TryGetProperty("snippet", out var sProp) ? sProp.GetString() : "";
                                                    var link = item.TryGetProperty("link", out var lProp) ? lProp.GetString() : "";
                                                    
                                                    // Try to extract structured data (pagemap)
                                                    string structuredInfo = "";
                                                    var price = "";
                                                    var currency = "";
                                                    var availability = "";

                                                    if (item.TryGetProperty("pagemap", out var pagemap) && pagemap.ValueKind == JsonValueKind.Object)
                                                    {
                                                        // 1. Look in 'offer'
                                                        if (pagemap.TryGetProperty("offer", out var offerArr) && offerArr.ValueKind == JsonValueKind.Array && offerArr.GetArrayLength() > 0)
                                                        {
                                                            var offer = offerArr[0];
                                                            if (offer.TryGetProperty("price", out var p)) price = p.ValueKind == JsonValueKind.String ? p.GetString() : p.ToString();
                                                            if (offer.TryGetProperty("pricecurrency", out var c)) currency = c.ValueKind == JsonValueKind.String ? c.GetString() : c.ToString();
                                                            if (offer.TryGetProperty("availability", out var a)) availability = a.ValueKind == JsonValueKind.String ? a.GetString() : a.ToString();
                                                        }
                                                        // 2. Look in 'product'
                                                        if (string.IsNullOrEmpty(price) && pagemap.TryGetProperty("product", out var productArr) && productArr.ValueKind == JsonValueKind.Array && productArr.GetArrayLength() > 0)
                                                        {
                                                            var product = productArr[0];
                                                            if (product.TryGetProperty("price", out var p)) price = p.ValueKind == JsonValueKind.String ? p.GetString() : p.ToString();
                                                        }
                                                        // 3. Look in 'metatags'
                                                        if (string.IsNullOrEmpty(price) && pagemap.TryGetProperty("metatags", out var metaArr) && metaArr.ValueKind == JsonValueKind.Array && metaArr.GetArrayLength() > 0)
                                                        {
                                                            var meta = metaArr[0];
                                                            if (meta.TryGetProperty("og:price:amount", out var p)) price = p.GetString();
                                                            if (meta.TryGetProperty("og:price:currency", out var c) && string.IsNullOrEmpty(currency)) currency = c.GetString();
                                                            if (string.IsNullOrEmpty(price) && meta.TryGetProperty("product:price:amount", out var p2)) price = p2.GetString();
                                                        }
                                                    }

                                                    // 4. Regex Fallback: Try to find price in snippet (e.g., "$12.99" or "£10")
                                                    if (string.IsNullOrEmpty(price) && !string.IsNullOrEmpty(snippet))
                                                    {
                                                        var match = System.Text.RegularExpressions.Regex.Match(snippet, @"(\$|£|€|¥)\s?(\d+([.,]\d{2})?)");
                                                        if (match.Success)
                                                        {
                                                            price = match.Value;
                                                        }
                                                    }

                                                    var infoParts = new List<string>();
                                                    if (!string.IsNullOrEmpty(price)) infoParts.Add($"Price={price} {currency}");
                                                    if (!string.IsNullOrEmpty(availability)) infoParts.Add($"Availability={availability}");

                                                    if (infoParts.Any())
                                                    {
                                                        structuredInfo = $" [Structured Data: {string.Join(", ", infoParts)}]";
                                                    }

                                                    // FILTERING: Try to skip links that look like category or search result pages
                                                    // -s.html is Temu collection, /k- is search, /s/ is search, /gc/ and /c/ are categories
                                                    bool isGeneric = link.Contains("-s.html", StringComparison.OrdinalIgnoreCase) ||
                                                                     link.Contains("/k-", StringComparison.OrdinalIgnoreCase) || 
                                                                     link.Contains("/s/", StringComparison.OrdinalIgnoreCase) || 
                                                                     link.Contains("/gc/", StringComparison.OrdinalIgnoreCase) || 
                                                                     link.Contains("/c/", StringComparison.OrdinalIgnoreCase) || 
                                                                     link.Contains("search", StringComparison.OrdinalIgnoreCase) ||
                                                                     link.Contains("category", StringComparison.OrdinalIgnoreCase) ||
                                                                     link.Contains("/list", StringComparison.OrdinalIgnoreCase) ||
                                                                     title.Contains("sold on Temu", StringComparison.OrdinalIgnoreCase) ||
                                                                     title.EndsWith("Results", StringComparison.OrdinalIgnoreCase) ||
                                                                     title.Contains("Electronics - Temu", StringComparison.OrdinalIgnoreCase);
                                                    
                                                    bool isLikelyProduct = link.Contains("-p.html", StringComparison.OrdinalIgnoreCase) ||
                                                                           link.Contains("/p/", StringComparison.OrdinalIgnoreCase) ||
                                                                           link.Contains("/product/", StringComparison.OrdinalIgnoreCase) ||
                                                                           link.Contains("/gp/product/", StringComparison.OrdinalIgnoreCase);

                                                    // KEYWORD SHIELD: Ensure the title has at least one core noun from the query
                                                    // We skip common stop words and check for matches
                                                    var coreKeywords = query.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries)
                                                                            .Where(w => w.Length > 2 && !"the and for with from temu amazon shop buy cheap best cheapest".Split(' ').Contains(w))
                                                                            .ToList();
                                                    bool matchesKeywords = !coreKeywords.Any() || coreKeywords.Any(k => title.ToLower().Contains(k));

                                                    // DEAD LINK DETECTION: Skip if snippet suggests it's gone
                                                    bool isDead = snippet.Contains("sold out", StringComparison.OrdinalIgnoreCase) ||
                                                                  snippet.Contains("out of stock", StringComparison.OrdinalIgnoreCase) ||
                                                                  snippet.Contains("discontinued", StringComparison.OrdinalIgnoreCase) ||
                                                                  snippet.Contains("removed", StringComparison.OrdinalIgnoreCase) ||
                                                                  snippet.Contains("no longer available", StringComparison.OrdinalIgnoreCase);

                                                    // Only keep if NOT dead AND (NOT generic OR high-confidence product link) AND matches keywords
                                                    if (!isDead && (!isGeneric || isLikelyProduct) && matchesKeywords)
                                                    {
                                                        if (foundCount++ >= 15) break;
                                                        results.AppendLine($"- **{title}**: {snippet}{structuredInfo} ({link})");
                                                    }
                                                    processedCount++;
                                                }
                                                
                                                searchResult = await _aiService.SummarizeSearchResultsAsync(query, results.ToString());
                                            }
                                            else
                                            {
                                                searchResult = "No results found (0 items returned).";
                                            }
                                        }
                                        else
                                        {
                                            searchResult = $"Google Search failed. Status Code: {response.StatusCode}";
                                        }
                                    }
                                }
                                else
                                {
                                    searchResult = "Search unavailable: Google API Key or Search Engine ID is missing in configuration.";
                                }
                            }
                            catch (Exception ex)
                            {
                                searchResult = $"Search execution error: {ex.Message}";
                            }

                            string prefix = "[Google Search]";
                            sb.AppendLine($"{prefix} for '{query}':\n{searchResult}");
                        }
                        else if (action == "SET_EMAIL_CREDENTIALS")
                        {
                            string email = parameters.TryGetProperty("email", out var e) ? e.GetString() ?? "" : "";
                            string pass = parameters.TryGetProperty("password", out var p) ? p.GetString() ?? "" : "";
                            
                            if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(pass))
                            {
                                var settingsPath = Path.Combine(_environment.ContentRootPath, "user_settings.json");
                                JsonObject settingsRoot;
                                
                                if (File.Exists(settingsPath))
                                {
                                    var existingJson = await File.ReadAllTextAsync(settingsPath);
                                    try {
                                        settingsRoot = JsonNode.Parse(existingJson)?.AsObject() ?? new JsonObject();
                                    } catch {
                                        settingsRoot = new JsonObject();
                                    }
                                }
                                else
                                {
                                    settingsRoot = new JsonObject();
                                }
                                
                                settingsRoot["Email:Username"] = email;
                                settingsRoot["Email:Password"] = pass;
                                
                                try {
                                    await File.WriteAllTextAsync(settingsPath, settingsRoot.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
                                    sb.AppendLine($"Successfully remembered credentials for {email}.");
                                } catch (Exception ex) {
                                    _logger.LogWarning($"Failed to save user settings: {ex.Message}");
                                    sb.AppendLine($"Successfully set credentials for session, but failed to save to disk: {email}.");
                                }
                            }
                            else
                            {
                                sb.AppendLine("Failed to set credentials: Email and Password are required.");
                            }
                        }
                        else if (action == "CHAT")
                        {
                            string message = "";
                            if (parameters.ValueKind != JsonValueKind.Undefined && parameters.TryGetProperty("message", out var msgProp)) 
                                message = msgProp.GetString() ?? "";
                            
                            return string.IsNullOrEmpty(message) ? "I'm sorry, I couldn't understand that command. Please try again or rephrase." : message;
                        }
                        else if (action == "SEND_EMAIL")
                        {
                            string to = "";
                            if (parameters.TryGetProperty("to", out var toProp)) to = toProp.GetString() ?? "";
                            if (string.IsNullOrEmpty(to) && parameters.TryGetProperty("email", out var emailProp)) to = emailProp.GetString() ?? "";
                            if (string.IsNullOrEmpty(to) && parameters.TryGetProperty("recipient", out var recProp)) to = recProp.GetString() ?? "";
                            if (string.IsNullOrEmpty(to) && parameters.TryGetProperty("address", out var addrProp)) to = addrProp.GetString() ?? "";

                            string subject = "No Subject";
                            if (parameters.TryGetProperty("subject", out var subProp)) subject = subProp.GetString() ?? "No Subject";

                            string body = "No Content";
                            if (parameters.TryGetProperty("body", out var bodyProp)) body = bodyProp.GetString() ?? "No Content";
                            
                            string? attachment = null;
                            if (parameters.TryGetProperty("attachment_path", out var attProp))
                                attachment = attProp.GetString();

                            string? user = null;
                            if (parameters.TryGetProperty("smtp_user", out var userProp))
                                user = userProp.GetString();

                            string? pass = null;
                            if (parameters.TryGetProperty("smtp_password", out var passProp))
                                pass = passProp.GetString();

                            try {
                                await SendEmailAsync(to, subject, body, attachment, user, pass);
                                sb.AppendLine($"Sent email to {to}.");
                            } catch (Exception ex) {
                                sb.AppendLine($"Failed to send email: {ex.Message}");
                            }
                        }
                        else if (action == "CREATE_DOCUMENT")
                        {
                            string fileName = $"new_doc_{DateTime.Now.Ticks}.txt";
                            if (parameters.TryGetProperty("file_name", out var fProp)) fileName = fProp.GetString() ?? fileName;

                            string type = "word";
                            if (parameters.TryGetProperty("type", out var tProp)) type = tProp.GetString() ?? "word";

                            bool wasPdfRequest = type.ToLower().Contains("pdf");
                            if (wasPdfRequest)
                            {
                                type = "word";
                                if (fileName.EndsWith(".pdf")) fileName = fileName.Substring(0, fileName.Length - 4);
                            }

                            string content = "";
                            if (parameters.TryGetProperty("content", out var cProp)) content = cProp.GetString() ?? "";

                            string primaryColor = "1C1C1C";
                            if (parameters.TryGetProperty("primary_color", out var pColorProp)) primaryColor = pColorProp.GetString()?.Replace("#", "") ?? "1C1C1C";
                            
                            string secondaryColor = "0078D7";
                            if (parameters.TryGetProperty("secondary_color", out var sColorProp)) secondaryColor = sColorProp.GetString()?.Replace("#", "") ?? "0078D7";

                            string imageKeyword = "abstract";
                            if (parameters.TryGetProperty("image_keyword", out var imgProp)) imageKeyword = imgProp.GetString() ?? "abstract";

                            string fileUrl = await CreateFileAsync(fileName, content, type, primaryColor, secondaryColor, imageKeyword);
                            
                            if (wasPdfRequest)
                                sb.AppendLine($"Created Word document (Save as PDF from Word): [{Path.GetFileName(fileUrl)}]({fileUrl}) [INTERNAL_CONTEXT: {content}]");
                            else
                                sb.AppendLine($"Created {type} document: [{Path.GetFileName(fileUrl)}]({fileUrl}) [INTERNAL_CONTEXT: {content}]");
                        }
                        else if (action == "CREATE_SPREADSHEET")
                        {
                            string fileName = $"data_{DateTime.Now.Ticks}.xlsx";
                            if (parameters.TryGetProperty("file_name", out var fProp)) fileName = fProp.GetString() ?? fileName;
                            if (!fileName.EndsWith(".xlsx")) fileName += ".xlsx";

                            string dataJson = "";
                            if (parameters.TryGetProperty("data", out var dProp)) 
                            {
                                if (dProp.ValueKind == JsonValueKind.Array)
                                    dataJson = dProp.GetRawText();
                                else
                                    dataJson = dProp.GetString() ?? "";
                            }

                            string fileUrl = await CreateExcelAsync(fileName, dataJson);
                            sb.AppendLine($"Created Excel spreadsheet: [{Path.GetFileName(fileUrl)}]({fileUrl})");
                        }
                        else if (action == "CREATE_CHART")
                        {
                            string fileName = $"chart_{DateTime.Now.Ticks}.png";
                            if (parameters.TryGetProperty("file_name", out var fProp)) fileName = fProp.GetString() ?? fileName;
                            if (!fileName.EndsWith(".png")) fileName += ".png";

                            string title = "Data Visualization";
                            if (parameters.TryGetProperty("title", out var tProp)) title = tProp.GetString() ?? title;

                            string chartType = "bar"; // bar, line, pie
                            if (parameters.TryGetProperty("chart_type", out var ctProp)) chartType = ctProp.GetString() ?? chartType;

                            string labels = "";
                            if (parameters.TryGetProperty("labels", out var lProp))
                            {
                                if (lProp.ValueKind == JsonValueKind.Array)
                                    labels = string.Join("|", lProp.EnumerateArray().Select(x => x.ToString()));
                                else
                                    labels = lProp.GetString() ?? "";
                            }

                            string values = "";
                            if (parameters.TryGetProperty("values", out var vProp))
                            {
                                if (vProp.ValueKind == JsonValueKind.Array)
                                    values = string.Join("|", vProp.EnumerateArray().Select(x => x.ToString()));
                                else
                                    values = vProp.GetString() ?? "";
                            }

                            string fileUrl = await CreateChartAsync(fileName, title, chartType, labels, values);
                            sb.AppendLine($"Generated {chartType} chart: ![{title}]({fileUrl})");
                        }
                        else if (action == "UNKNOWN" || action == "error")
                        {
                            var msg = "";
                            if (parameters.ValueKind != JsonValueKind.Undefined && parameters.TryGetProperty("message", out var msgProp)) 
                            {
                                msg = msgProp.GetString();
                            }
                            
                            if (string.IsNullOrEmpty(msg) && actionElement.TryGetProperty("message", out var msgRoot))
                            {
                                msg = msgRoot.GetString();
                            }
                            
                            sb.AppendLine($"AI Error: {msg}");
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                return $"Error executing actions: {ex.Message}";
            }

            var finalResult = sb.ToString().Trim();
            if (string.IsNullOrEmpty(finalResult))
            {
                // Fallback: If no action was taken, return a slightly more diagnostic message
                // AND log the raw response to see what was missed
                await File.WriteAllTextAsync("wwwroot/exports/debug_unhandled_actions.txt", $"[{DateTime.Now}] Unhandled AI Response:\n{jsonResponse}");
                return "I'm sorry, I couldn't identify a specific action for that request. If you are asking for sensitive content, I may be prohibited from generating it. Please try rephrasing or asking for something else. [System V2.2]";
            }

            return finalResult;
        }

        public async Task<string> CreateFileAsync(string fileName, string content, string type = "word", string primaryColor = "1C1C1C", string secondaryColor = "0078D7", string imageKeyword = "")
        {
            var folderName = "exports";
            var path = Path.Combine(_environment.WebRootPath, folderName);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, fileName);

            if (type.ToLower().Contains("word"))
            {
                if (!fileName.EndsWith(".docx")) fileName += ".docx";
                filePath = Path.Combine(path, fileName);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    var lines = content.Replace("\r\n", "\n").Split('\n');

                    foreach (var line in lines)
                    {
                        Paragraph para = body.AppendChild(new Paragraph());
                        DocumentFormat.OpenXml.Wordprocessing.Run run = para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(line) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }
            }
            else if (type.ToLower().Contains("presentation") || type.ToLower().Contains("ppt") || type.ToLower().Contains("powerpoint"))
            {
                if (!fileName.EndsWith(".pptx")) fileName += ".pptx";
                filePath = Path.Combine(path, fileName);
                await CreatePowerPointAsync(filePath, content, primaryColor, secondaryColor, imageKeyword);
            }
            else
            {
                if (!fileName.EndsWith(".txt") && !fileName.Contains(".")) fileName += ".txt";
                filePath = Path.Combine(path, fileName);
                await File.WriteAllTextAsync(filePath, content);
            }

            return $"/{folderName}/{fileName}";
        }

        public async Task<string> CreateExcelAsync(string fileName, string dataJson)
        {
            var folderName = "exports";
            var path = Path.Combine(_environment.WebRootPath, folderName);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, fileName);

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

                // Parse dataJson (Expects a simple array of arrays or similar)
                try {
                    using var doc = JsonDocument.Parse(dataJson);
                    if (doc.RootElement.ValueKind == JsonValueKind.Array) {
                        bool isFirstRow = true;
                        List<string> headers = new List<string>();

                        foreach (var rowElement in doc.RootElement.EnumerateArray()) {
                            Row row = new Row();
                            
                            // Case 1: Array of Arrays (Standard)
                            if (rowElement.ValueKind == JsonValueKind.Array) {
                                foreach (var cellValue in rowElement.EnumerateArray()) {
                                    string val = cellValue.ToString();
                                    // Handle implicit CSV strings inside single-cell arrays (e.g. ["A,B,C"])
                                    if (rowElement.GetArrayLength() == 1 && val.Contains(",")) 
                                    {
                                        foreach(var part in val.Split(',')) 
                                            row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(part.Trim()) });
                                    }
                                    else 
                                    {
                                        row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(val) });
                                    }
                                }
                            }
                            // Case 2: Array of Objects (Key-Value)
                            else if (rowElement.ValueKind == JsonValueKind.Object) {
                                // If first row, generate header
                                if (isFirstRow) {
                                    Row headerRow = new Row();
                                    foreach (var prop in rowElement.EnumerateObject()) {
                                        headers.Add(prop.Name);
                                        headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(prop.Name) });
                                    }
                                    sheetData.Append(headerRow);
                                    isFirstRow = false; // Prevent logic from adding another header
                                }
                                
                                // Add values in order of headers
                                foreach (var header in headers) {
                                    string val = "";
                                    if (rowElement.TryGetProperty(header, out var valProp)) val = valProp.ToString();
                                    row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(val) });
                                }
                            }
                            // Case 3: Array of Strings (CSV Rows?)
                            else if (rowElement.ValueKind == JsonValueKind.String) {
                                var val = rowElement.GetString() ?? "";
                                if (val.Contains(",")) {
                                     foreach(var part in val.Split(',')) 
                                        row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(part.Trim()) });
                                } else {
                                     row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(val) });
                                }
                            }

                            sheetData.Append(row);
                        }
                    }
                } catch (Exception ex) {
                    _logger.LogError(ex, "Error parsing Excel data JSON");
                    Row errorRow = new Row();
                    errorRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue("Error parsing data") });
                    sheetData.Append(errorRow);
                }

                workbookPart.Workbook.Save();
            }

            return $"/{folderName}/{fileName}";
        }

        public async Task<string> CreateChartAsync(string fileName, string title, string chartType, string labels, string values)
        {
            var folderName = "exports";
            var path = Path.Combine(_environment.WebRootPath, folderName);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, fileName);

            // Use QuickChart.io for simple, high-quality chart generation
            // Example: https://quickchart.io/chart?c={type:'bar',data:{labels:['Q1','Q2'], datasets:[{label:'Users',data:[50,60]}]}}
            
            // Use colorful backgrounds for all chart types
            object backgroundColor = new[] {
                "rgba(255, 99, 132, 0.7)",
                "rgba(54, 162, 235, 0.7)",
                "rgba(255, 206, 86, 0.7)",
                "rgba(75, 192, 192, 0.7)",
                "rgba(153, 102, 255, 0.7)",
                "rgba(255, 159, 64, 0.7)",
                "rgba(199, 199, 199, 0.7)",
                "rgba(83, 102, 255, 0.7)",
                "rgba(255, 99, 255, 0.7)",
                "rgba(99, 255, 132, 0.7)"
            };

            var chartConfig = new {
                type = chartType,
                data = new {
                    labels = labels.Split('|'),
                    datasets = new[] {
                        new {
                            label = title,
                            data = values.Split('|').Select(v => decimal.TryParse(v, out var d) ? d : 0).ToArray(),
                            backgroundColor = backgroundColor,
                            borderColor = "rgba(0,0,0,0.1)",
                            borderWidth = 1
                        }
                    }
                },
                options = new {
                    title = new {
                        display = true,
                        text = title
                    }
                }
            };

            var jsonConfig = JsonSerializer.Serialize(chartConfig);
            var url = $"https://quickchart.io/chart?c={Uri.EscapeDataString(jsonConfig)}&width=800&height=400";

            try {
                using (var client = new HttpClient()) {
                    var bytes = await client.GetByteArrayAsync(url);
                    await File.WriteAllBytesAsync(filePath, bytes);
                }
            } catch (Exception ex) {
                _logger.LogError(ex, "Error generating chart image");
                // Create a dummy placeholder if it fails
                await File.WriteAllTextAsync(filePath, "Chart generation failed.");
            }

            return $"/{folderName}/{fileName}";
        }

        private async Task CreatePowerPointAsync(string filepath, string content, string primaryColor, string secondaryColor, string imageKeyword)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation))
            {
                PresentationPart presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new Presentation();
                
                // Add Mandatory Presentation Properties
                var pp = presentationPart.AddNewPart<PresentationPropertiesPart>();
                pp.PresentationProperties = new PresentationProperties();
                pp.PresentationProperties.Save();

                // Add View Properties
                var vp = presentationPart.AddNewPart<ViewPropertiesPart>();
                vp.ViewProperties = new ViewProperties();
                vp.ViewProperties.Save();

                SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                var masterShapeTree = new ShapeTree();
                InitializeShapeTree(masterShapeTree);
                slideMasterPart.SlideMaster = new SlideMaster(new CommonSlideData(masterShapeTree), new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink });
                
                // Add Mandatory Theme Part to Slide Master
                var themePart = slideMasterPart.AddNewPart<ThemePart>();
                themePart.Theme = new A.Theme() { Name = "Default Theme" };
                themePart.Theme.Append(new A.ThemeElements(
                    new A.ColorScheme(
                        new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                        new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                        new A.Dark2Color(new A.RgbColorModelHex() { Val = "1C1C1C" }),
                        new A.Light2Color(new A.RgbColorModelHex() { Val = "E0E0E0" }),
                        new A.Accent1Color(new A.RgbColorModelHex() { Val = primaryColor }),
                        new A.Accent2Color(new A.RgbColorModelHex() { Val = secondaryColor }),
                        new A.Accent3Color(new A.RgbColorModelHex() { Val = "FF0000" }),
                        new A.Accent4Color(new A.RgbColorModelHex() { Val = "00FF00" }),
                        new A.Accent5Color(new A.RgbColorModelHex() { Val = "0000FF" }),
                        new A.Accent6Color(new A.RgbColorModelHex() { Val = "FFFF00" }),
                        new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
                        new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" })
                    ) { Name = "Office" },
                    new A.FontScheme(
                        new A.MajorFont(new A.LatinFont() { Typeface = "Calibri" }, new A.EastAsianFont() { Typeface = "" }, new A.ComplexScriptFont() { Typeface = "" }),
                        new A.MinorFont(new A.LatinFont() { Typeface = "Calibri" }, new A.EastAsianFont() { Typeface = "" }, new A.ComplexScriptFont() { Typeface = "" })
                    ) { Name = "Office" },
                    new A.FormatScheme(
                        new A.FillStyleList(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 })),
                        new A.LineStyleList(new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }))),
                        new A.EffectStyleList(new A.EffectStyle(new A.EffectList())),
                        new A.BackgroundFillStyleList(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }))
                    ) { Name = "Office" }
                ));
                themePart.Theme.Save();

                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                var layoutShapeTree = new ShapeTree();
                InitializeShapeTree(layoutShapeTree);
                slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(layoutShapeTree));
                
                // Link Layout to Master in the XML
                var slideLayoutIdList = new SlideLayoutIdList();
                slideLayoutIdList.Append(new SlideLayoutId() { Id = 2147483649u, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) });
                slideMasterPart.SlideMaster.Append(slideLayoutIdList);

                slideLayoutPart.SlideLayout.Save();
                slideMasterPart.SlideMaster.Save();

                // BUILD PRESENTATION - STRICT SCHEMA ORDER
                var masterIdList = new SlideMasterIdList();
                masterIdList.Append(new SlideMasterId() { Id = 2147483648u, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });
                presentationPart.Presentation.Append(masterIdList);
                
                var slideIdList = new SlideIdList();
                presentationPart.Presentation.Append(slideIdList);

                presentationPart.Presentation.Append(new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 });
                presentationPart.Presentation.Append(new NotesSize() { Cx = 6858000, Cy = 9144000 });
                presentationPart.Presentation.Append(new DefaultTextStyle());

                var slides = ParseSlides(content);
                uint slideId = 256;

                foreach (var slideData in slides)
                {
                    // ... (image search logic remains unchanged) ...
                    string? localImagePath = null;
                    try
                    {
                        var words = new List<string>();
                        if (!string.IsNullOrEmpty(imageKeyword))
                        {
                            words.Add(imageKeyword);
                        }

                        // CLEANING: Strip AI-generated garbage from the content keywords
                        string rawKeywords = !string.IsNullOrWhiteSpace(slideData.Keywords) ? slideData.Keywords : slideData.Title;
                        string cleanedKeywords = System.Text.RegularExpressions.Regex.Replace(rawKeywords, @"(?i)(Visual Keywords:|Slide \d+:|abstract|photorealistic|macro photography)", "").Trim();

                        var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "a", "an", "the", "to", "in", "on", "at", "by", "for", "with", "is", "of", "and", "or", "some", "any" };
                        var genericTitles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "conclusion", "introduction", "intro", "summary", "overview", "q&a", "thanks", "thank you", "questions", "references" };
                        
                        var slideWords = cleanedKeywords.Split(new[] { ' ', ',', '.', ':', ';', '!', '?' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Where(w => w.Length > 2 && !stopWords.Contains(w) && !genericTitles.Contains(w));

                        if (genericTitles.Contains(slideData.Title.Trim()) && !string.IsNullOrEmpty(imageKeyword))
                        {
                            // If the title is generic, try to find a meaningful word in the body
                            var bodyWords = (slideData.Body ?? "").Split(new[] { ' ', ',', '.', ':', ';', '!', '?' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Where(w => w.Length > 4 && !stopWords.Contains(w) && !genericTitles.Contains(w))
                                           .Take(2);
                            
                            words.Add(imageKeyword);
                            words.AddRange(bodyWords);
                        }
                        else 
                        {
                            words.AddRange(slideWords);
                            if (!string.IsNullOrEmpty(imageKeyword) && !words.Contains(imageKeyword, StringComparer.OrdinalIgnoreCase))
                                words.Insert(0, imageKeyword);
                        }

                        // Final keyword string: limit to 2-3 specific nouns for better search match
                        var slideImgKeyword = string.Join(" ", words.Distinct(StringComparer.OrdinalIgnoreCase).Take(3));
                        localImagePath = Path.Combine(_environment.WebRootPath, "exports", $"thumb_{Guid.NewGuid()}.jpg");
                        
                        var googleApiKey = _configuration["Google:ApiKey"];
                        var googleCx = _configuration["Google:SearchEngineId"];
                        
                        // Static flag to skip if quota already hit in this session/process (optional, but good for speed)
                        bool canUseGoogle = !string.IsNullOrEmpty(googleApiKey) && !string.IsNullOrEmpty(googleCx) && googleCx != "YOUR_SEARCH_ENGINE_ID_HERE";

                        if (canUseGoogle)
                        {
                            // Reset quota flag if it's a new day
                            if (DateTime.Now.Date > _lastQuotaReset.Date)
                            {
                                _googleQuotaExceeded = false;
                                _lastQuotaReset = DateTime.Now;
                            }

                            if (_googleQuotaExceeded)
                            {
                                canUseGoogle = false; // Skip if we already know quota is gone
                            }
                        }

                        if (canUseGoogle)
                        {
                            try 
                            {
                                using (var client = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(10) })
                                {
                                    var searchUrl = $"https://www.googleapis.com/customsearch/v1?key={googleApiKey}&cx={googleCx}&q={Uri.EscapeDataString(slideImgKeyword)}&searchType=image&num=1";
                                    var response = await client.GetAsync(searchUrl);
                                    if (response.IsSuccessStatusCode)
                                    {
                                        var searchJson = await response.Content.ReadAsStringAsync();
                                        using var searchDoc = JsonDocument.Parse(searchJson);
                                        if (searchDoc.RootElement.TryGetProperty("items", out var items) && items.ValueKind == JsonValueKind.Array && items.GetArrayLength() > 0)
                                        {
                                            var imageUrl = items[0].TryGetProperty("link", out var l) ? l.GetString() : null;
                                            if (!string.IsNullOrEmpty(imageUrl))
                                            {
                                                var bytes = await client.GetByteArrayAsync(imageUrl);
                                                await File.WriteAllBytesAsync(localImagePath, bytes);
                                            }
                                        }
                                        else if (!string.IsNullOrEmpty(imageKeyword) && slideImgKeyword != imageKeyword)
                                        {
                                            // Retry with just the broad imageKeyword
                                            var retryUrl = $"https://www.googleapis.com/customsearch/v1?key={googleApiKey}&cx={googleCx}&q={Uri.EscapeDataString(imageKeyword)}&searchType=image&num=1";
                                            var retryResp = await client.GetAsync(retryUrl);
                                            if (retryResp.IsSuccessStatusCode)
                                            {
                                                var retryJson = await retryResp.Content.ReadAsStringAsync();
                                                using var retryDoc = JsonDocument.Parse(retryJson);
                                                if (retryDoc.RootElement.TryGetProperty("items", out var rItems) && rItems.ValueKind == JsonValueKind.Array && rItems.GetArrayLength() > 0)
                                                {
                                                    var imageUrl = rItems[0].TryGetProperty("link", out var l) ? l.GetString() : null;
                                                    if (!string.IsNullOrEmpty(imageUrl))
                                                    {
                                                        var bytes = await client.GetByteArrayAsync(imageUrl);
                                                        await File.WriteAllBytesAsync(localImagePath, bytes);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var errorContent = await response.Content.ReadAsStringAsync();
                                        if (response.StatusCode == System.Net.HttpStatusCode.TooManyRequests || errorContent.Contains("quotaExceeded", StringComparison.OrdinalIgnoreCase))
                                        {
                                            _googleQuotaExceeded = true;
                                            _lastQuotaReset = DateTime.Now;
                                        }
                                        _logger.LogWarning($"Google Image Search failed for '{slideData.Title}': {response.StatusCode} - {errorContent}");
                                        var failureLogPath = Path.Combine(_environment.WebRootPath, "exports", "debug_image_search_failure.txt");
                                        var failureLogDir = Path.GetDirectoryName(failureLogPath);
                                        if (!Directory.Exists(failureLogDir)) Directory.CreateDirectory(failureLogDir!);
                                        await File.AppendAllTextAsync(failureLogPath, $"[{DateTime.Now}] Slide: {slideData.Title} | Keyword: {slideImgKeyword} | Status: {response.StatusCode} | Body: {errorContent}\n");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                 _logger.LogError(ex, "Error during image download");
                            }
                        }

                        if (!File.Exists(localImagePath))
                        {
                            using (var client = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(8) })
                            {
                                var randomLock = System.Random.Shared.Next(1, 10000);
                                var bytes = await client.GetByteArrayAsync($"https://loremflickr.com/800/600/{slideImgKeyword.Replace(" ", ",")}?lock={randomLock}");
                                await File.WriteAllBytesAsync(localImagePath, bytes);
                            }
                        }
                    }
                    catch { localImagePath = null; }

                    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.AddPart(slideLayoutPart);
                    
                    var slideShapeTree = new ShapeTree();
                    InitializeShapeTree(slideShapeTree);
                    var slide = new Slide(new CommonSlideData(slideShapeTree));
                    var shapeTree = slide.CommonSlideData.ShapeTree;
                    
                    AddBackground(shapeTree, primaryColor); 
                    AddTextBox(shapeTree, slideData.Title, 500000, 300000, 8000000, 1000000, 4400, "FFFFFF", true); 
                    AddRectangle(shapeTree, 500000, 1200000, 3000000, 50000, secondaryColor); 

                    if (!string.IsNullOrWhiteSpace(slideData.Body))
                    {
                        AddTextBox(shapeTree, slideData.Body, 500000, 1500000, 5500000, 5000000, 2000, "E0E0E0", false); 
                    }

                    slidePart.Slide = slide;

                    if (localImagePath != null && File.Exists(localImagePath))
                    {
                        AddImage(slidePart, localImagePath, 6200000, 1500000, 2500000, 2500000);
                    }
                    else
                    {
                        AddPlaceholderImage(shapeTree, 6200000, 1500000, 2500000, 2500000);
                    }

                    slidePart.Slide.Save();
                    slideIdList.Append(new SlideId() { Id = slideId++, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
                }

                presentationPart.Presentation.Save();
            }
        }

        private void InitializeShapeTree(ShapeTree shapeTree)
        {
            shapeTree.Append(new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()));
            shapeTree.Append(new P.GroupShapeProperties(new A.Transform2D()));
        }

        private void AddImage(SlidePart slidePart, string imagePath, long x, long y, long cx, long cy)
        {
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                imagePart.FeedData(stream);
            }

            var pic = new P.Picture();
            pic.NonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)(uint)System.Random.Shared.Next(1000, 999999), Name = "TopicImage" },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            pic.BlipFill = new P.BlipFill(
                new A.Blip() { Embed = slidePart.GetIdOfPart(imagePart) },
                new A.Stretch(new A.FillRectangle())
            );

            pic.ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = x, Y = y },
                    new A.Extents() { Cx = cx, Cy = cy }
                ),
                new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle }
            );

            slidePart.Slide.CommonSlideData.ShapeTree.Append(pic);
        }

        private List<(string Title, string Keywords, string Body)> ParseSlides(string content)
        {
            var result = new List<(string, string, string)>();
            if (string.IsNullOrEmpty(content)) return result;

            var lines = content.Replace("\r\n", "\n").Split('\n');
            string? currentTitle = null;
            string? currentKeywords = null;
            var currentBody = new StringBuilder();

            foreach (var line in lines)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^Slide\s+\d+[:|-]", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    if (currentTitle != null)
                    {
                        result.Add((currentTitle, currentKeywords ?? "", currentBody.ToString().Trim()));
                    }
                    var rawTitle = line.Trim();
                    currentTitle = System.Text.RegularExpressions.Regex.Replace(rawTitle, @"^Slide\s+\d+[:|-]\s*", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (string.IsNullOrWhiteSpace(currentTitle)) currentTitle = rawTitle;
                    
                    currentKeywords = null;
                    currentBody.Clear();
                }
                else if (line.Trim().StartsWith("Visual Keywords:", StringComparison.OrdinalIgnoreCase))
                {
                    currentKeywords = line.Trim().Substring("Visual Keywords:".Length).Trim();
                }
                else
                {
                    currentBody.AppendLine(line);
                }
            }

            if (currentTitle != null)
            {
                result.Add((currentTitle, currentKeywords ?? "", currentBody.ToString().Trim()));
            }
            else if (result.Count == 0 && currentBody.Length > 0)
            {
                result.Add(("Generated Content", "", currentBody.ToString().Trim()));
            }

            return result;
        }

        private void AddBackground(ShapeTree shapeTree, string hexColor)
        {
            var shape = new P.Shape();
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)(uint)System.Random.Shared.Next(100, 1000000), Name = "Background" },
               new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());

            shape.ShapeProperties = new P.ShapeProperties();
            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = 0, Y = 0 };
            transform2D.Extents = new A.Extents() { Cx = 9144000, Cy = 6858000 }; 
            shape.ShapeProperties.Append(transform2D);

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.AdjustValueList = new A.AdjustValueList();
            shape.ShapeProperties.Append(presetGeometry);

            var solidFill = new A.SolidFill();
             var srgbClr = new A.RgbColorModelHex() { Val = hexColor };
            solidFill.Append(srgbClr);
            shape.ShapeProperties.Append(solidFill);

            shapeTree.Append(shape);
        }

        private void AddRectangle(ShapeTree shapeTree, long x, long y, long cx, long cy, string hexColor)
        {
             var shape = new P.Shape();
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)(uint)System.Random.Shared.Next(100, 1000000), Name = "Rect" },
               new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());

            shape.ShapeProperties = new P.ShapeProperties();
            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = cx, Cy = cy };
            shape.ShapeProperties.Append(transform2D);

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.AdjustValueList = new A.AdjustValueList();
            shape.ShapeProperties.Append(presetGeometry);

            var solidFill = new A.SolidFill();
            var srgbClr = new A.RgbColorModelHex() { Val = hexColor };
            solidFill.Append(srgbClr);
            shape.ShapeProperties.Append(solidFill);

            shapeTree.Append(shape);
        }

        private void AddPlaceholderImage(ShapeTree shapeTree, long x, long y, long cx, long cy)
        {
             var shape = new P.Shape();
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)(uint)System.Random.Shared.Next(100, 1000000), Name = "ImagePlaceholder" },
               new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());

            shape.ShapeProperties = new P.ShapeProperties();
            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = cx, Cy = cy };
            shape.ShapeProperties.Append(transform2D);

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.AdjustValueList = new A.AdjustValueList();
            shape.ShapeProperties.Append(presetGeometry);

            var solidFill = new A.SolidFill();
            var srgbClr = new A.RgbColorModelHex() { Val = "333333" };
            solidFill.Append(srgbClr);
            shape.ShapeProperties.Append(solidFill);
            
             var ln = new A.Outline() { Width = 12700 }; 
             ln.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = "666666" }));
             shape.ShapeProperties.Append(ln);

            shape.TextBody = new P.TextBody();
            shape.TextBody.BodyProperties = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };
            shape.TextBody.ListStyle = new A.ListStyle();

            var paragraph = new A.Paragraph();
            paragraph.ParagraphProperties = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            var run = new A.Run();
            run.RunProperties = new A.RunProperties() { Language = "en-US", FontSize = 1400 }; 
            run.RunProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = "888888" }));
            run.Text = new A.Text("[Visual/Image]");

            paragraph.Append(run);
            shape.TextBody.Append(paragraph);

            shapeTree.Append(shape);
        }

        private void AddTextBox(ShapeTree shapeTree, string text, long x, long y, long cx, long cy, int fontSize, string hexColor = "000000", bool isBold = false)
        {
           var shape = new P.Shape();
           
           shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)(uint)System.Random.Shared.Next(100, 1000000), Name = "TextBox" },
               new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());

           shape.ShapeProperties = new P.ShapeProperties();
           
           var transform2D = new A.Transform2D();
           transform2D.Offset = new A.Offset() { X = x, Y = y };
           transform2D.Extents = new A.Extents() { Cx = cx, Cy = cy };
           shape.ShapeProperties.Append(transform2D);

           var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
           presetGeometry.AdjustValueList = new A.AdjustValueList();
           shape.ShapeProperties.Append(presetGeometry);

           shape.ShapeProperties.Append(new A.NoFill());

           shape.TextBody = new P.TextBody();
           shape.TextBody.BodyProperties = new A.BodyProperties();
           shape.TextBody.ListStyle = new A.ListStyle();

           var lines = text.Split('\n');
           foreach(var line in lines)
           {
               var cleanLine = line.Trim();
               if(string.IsNullOrEmpty(cleanLine)) continue;

               var paragraph = new A.Paragraph();
               var run = new A.Run();
               run.RunProperties = new A.RunProperties() { Language = "en-US", FontSize = fontSize, Bold = isBold }; 
               
               var solidFill = new A.SolidFill();
               var srgbClr = new A.RgbColorModelHex() { Val = hexColor };
               solidFill.Append(srgbClr);
               run.RunProperties.Append(solidFill);

               run.Text = new A.Text(cleanLine);
               paragraph.Append(run);
               shape.TextBody.Append(paragraph);
           }

           shapeTree.Append(shape);
        }

        public async Task SendEmailAsync(string to, string subject, string body, string? attachmentPath = null, string? username = null, string? password = null)
        {
            string? finalUser = null;
            string? finalPass = null;
            try
            {
                _logger.LogInformation($"Attempting to send email to {to} with subject '{subject}'");

                // Check for SendGrid API Key (HTTP Fallback)
                var sendGridApiKey = _configuration["Email:SendGridApiKey"];
                var senderEmail = _configuration["Email:SenderEmail"];
                var senderName = _configuration["Email:SenderName"] ?? "TaskManager AI";

                // Use SendGrid ONLY if NO specific username/password was passed in
                if (!string.IsNullOrEmpty(sendGridApiKey) && string.IsNullOrEmpty(username) && string.IsNullOrEmpty(password))
                {
                    _logger.LogInformation("No dynamic credentials provided. Using SendGrid HTTP API fallback.");
                    await SendViaSendGridHttpAsync(sendGridApiKey, to, subject, body, senderEmail, senderName, attachmentPath);
                    await LogEmailSimulation(to, subject, body, attachmentPath, "SendGridAPI", "****", "SUCCESS: Sent via SendGrid HTTP");
                    return;
                }

                // --- Standard SMTP Logic (MailKit) ---
                // This will be used if: 
                // 1. Specific username/password were passed (Dynamic)
                // 2. OR SendGrid is not configured (Default SMTP)
                // Ensure exports directory exists for debug logging
                var debugPath = Path.Combine(_environment.WebRootPath, "exports", "debug_email.txt");
                var debugDir = Path.GetDirectoryName(debugPath);
                if (!Directory.Exists(debugDir)) Directory.CreateDirectory(debugDir!);
                
                // Fallback to configuration if not provided
                var smtpServer = _configuration["Email:SmtpServer"] ?? "smtp.gmail.com";
                var portStr = _configuration["Email:Port"] ?? "465"; 
                if (!int.TryParse(portStr, out int port)) 
                {
                    port = 465;
                    _logger.LogWarning($"Failed to parse Email:Port '{portStr}', defaulting to 465.");
                }
                
                var configUser = _configuration["Email:Username"];
                var configPass = _configuration["Email:Password"];

                bool isPlaceholder = configUser == "YOUR_EMAIL@gmail.com" || configPass == "YOUR_APP_PASSWORD_HERE";

                if (string.IsNullOrEmpty(configUser) || string.IsNullOrEmpty(configPass) || isPlaceholder)
                {
                    var settingsPath = Path.Combine(_environment.ContentRootPath, "user_settings.json");
                    if (File.Exists(settingsPath))
                    {
                        try
                        {
                            var settingsJson = await File.ReadAllTextAsync(settingsPath);
                            var settingsNode = JsonNode.Parse(settingsJson);
                            if (settingsNode != null)
                            {
                                if ((string.IsNullOrEmpty(configUser) || configUser == "YOUR_EMAIL@gmail.com") && settingsNode["Email:Username"] != null) 
                                {
                                    configUser = settingsNode["Email:Username"]?.GetValue<string>();
                                }
                                if ((string.IsNullOrEmpty(configPass) || configPass == "YOUR_APP_PASSWORD_HERE") && settingsNode["Email:Password"] != null)
                                {
                                    configPass = settingsNode["Email:Password"]?.GetValue<string>();
                                }
                            }
                        }
                        catch (Exception ex) 
                        { 
                             _logger.LogError(ex, "Error parsing user_settings.json for email credentials");
                        }
                    }
                }

                if (string.IsNullOrEmpty(senderEmail)) senderEmail = configUser;
                
                finalUser = !string.IsNullOrEmpty(username) ? username : configUser;
                finalPass = !string.IsNullOrEmpty(password) ? password : configPass;
                
                _logger.LogInformation($"Email Configuration: Server={smtpServer}:{port}, Sender={senderEmail}, Target={to}");

                if (string.IsNullOrEmpty(finalUser) || string.IsNullOrEmpty(finalPass))
                {
                   var msg = "FAILED: No Credentials Provided (User or Password missing)";
                   _logger.LogWarning(msg);
                   await LogEmailSimulation(to, subject, body, attachmentPath, finalUser, finalPass, msg);
                   throw new Exception(msg); 
                }

                using (var message = new MimeKit.MimeMessage())
                {
                    message.From.Add(new MimeKit.MailboxAddress(senderName, senderEmail ?? finalUser));
                    message.To.Add(MimeKit.MailboxAddress.Parse(to));
                    message.Subject = subject;

                    var builder = new MimeKit.BodyBuilder();
                    builder.HtmlBody = body;

                    if (!string.IsNullOrEmpty(attachmentPath))
                    {
                        var fullPath = attachmentPath;
                        if (!File.Exists(fullPath))
                        {
                            var exportsPath = Path.Combine(_environment.WebRootPath, "exports", attachmentPath);
                            if (File.Exists(exportsPath)) fullPath = exportsPath;
                        }

                        if (File.Exists(fullPath))
                        {
                            builder.Attachments.Add(fullPath);
                        }
                    }

                    message.Body = builder.ToMessageBody();

                    using (var client = new MailKit.Net.Smtp.SmtpClient())
                    {
                        client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                        var options = port == 465 ? MailKit.Security.SecureSocketOptions.SslOnConnect : MailKit.Security.SecureSocketOptions.StartTls;
                        
                        await client.ConnectAsync(smtpServer, port, options);
                        await client.AuthenticateAsync(finalUser, finalPass);
                        await client.SendAsync(message);
                        await client.DisconnectAsync(true);

                        _logger.LogInformation("Email sent successfully via MailKit.");
                        await LogEmailSimulation(to, subject, body, attachmentPath, finalUser, finalPass, "SUCCESS: Sent via MailKit");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to send email to {to}");
                await LogEmailSimulation(to, subject, body, attachmentPath, finalUser, finalPass, $"ERROR: {ex.Message}");
                throw new Exception($"Failed to send email: {ex.Message}");
            }
        }

        private async Task SendViaSendGridHttpAsync(string apiKey, string to, string subject, string body, string? senderEmail, string senderName, string? attachmentPath)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                var payload = new
                {
                    personalizations = new[] { new { to = new[] { new { email = to } } } },
                    from = new { email = !string.IsNullOrEmpty(senderEmail) ? senderEmail : "no-reply@taskmanager.com", name = senderName },
                    subject = subject,
                    content = new[] { new { type = "text/html", value = body } }
                };

                var json = JsonSerializer.Serialize(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("https://api.sendgrid.com/v3/mail/send", content);
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"SendGrid API Status: {response.StatusCode}, Details: {error}");
                }
            }
        }

        private async Task LogEmailSimulation(string to, string subject, string body, string? attachmentPath, string? username, string? password, string status)
        {
            try {
                var emailLogPath = Path.Combine(_environment.WebRootPath, "exports", "email_logs.txt");
                var directory = Path.GetDirectoryName(emailLogPath);
                if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);

                var authInfo = (username != null && password != null) ? $"[Authenticated as {username}]" : "[No Credentials]";
                var attInfo = attachmentPath != null ? $"[Attachment: {attachmentPath}]" : "";

                var logEntry = $"[{DateTime.Now}] {status} {authInfo} To: {to}, Subject: {subject}, Body: {body} {attInfo}\n----------------------------------------\n";
                await File.AppendAllTextAsync(emailLogPath, logEntry);
            } catch (Exception ex) {
                _logger.LogWarning($"Could not log email simulation: {ex.Message}");
            }
        }
    }
}
