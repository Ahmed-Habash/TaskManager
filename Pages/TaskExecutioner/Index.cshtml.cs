using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using TaskManager.Services;

namespace TaskManager.Pages.TaskExecutioner
{
    [IgnoreAntiforgeryToken]
    public class IndexModel : PageModel
    {
        private readonly ITaskExecutionerService _service;
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ITaskExecutionerService service, ILogger<IndexModel> logger)
        {
            _service = service;
            _logger = logger;
            ConversationHistory = new List<ChatMessage>();
        }

        [BindProperty]
        [Required]
        public string Command { get; set; }

        [BindProperty]
        public string HistoryJson { get; set; }

        public List<ChatMessage> ConversationHistory { get; set; }

        public void OnGet()
        {
            // Initial greeting
            ConversationHistory.Add(new ChatMessage { Role = "assistant", Content = "Hello! I can send emails, create files, and search the web. How can I help you today?" });
            HistoryJson = JsonSerializer.Serialize(ConversationHistory);
        }

        public async Task<IActionResult> OnPostExecuteCommandAsync()
        {
            if (!string.IsNullOrEmpty(HistoryJson))
            {
                try 
                {
                    ConversationHistory = JsonSerializer.Deserialize<List<ChatMessage>>(HistoryJson) ?? new List<ChatMessage>();
                }
                catch
                {
                    ConversationHistory = new List<ChatMessage>();
                }
            }
            else 
            {
                 ConversationHistory = new List<ChatMessage>();
            }

            if (!ModelState.IsValid)
            {
                return Page();
            }

            ConversationHistory.Add(new ChatMessage { Role = "user", Content = Command });

            try
            {
                var result = await _service.ExecuteAICommandAsync(Command, ConversationHistory);
                ConversationHistory.Add(new ChatMessage { Role = "assistant", Content = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing command");
                ConversationHistory.Add(new ChatMessage { Role = "assistant", Content = $"Sorry, I encountered an error: {ex.Message}" });
            }

            HistoryJson = JsonSerializer.Serialize(ConversationHistory);
            Command = string.Empty;

            return Page();
        }

        public async Task<IActionResult> OnPostChatAsync([FromBody] ChatRequest request)
        {
            try
            {
                if (request == null)
                    return new JsonResult(new { success = false, message = "Invalid request body (null)." });

                if (string.IsNullOrEmpty(request.Command))
                    return new JsonResult(new { success = false, message = "Empty command." });

                var history = new List<ChatMessage>();
                if (!string.IsNullOrEmpty(request.HistoryJson))
                {
                    try { history = JsonSerializer.Deserialize<List<ChatMessage>>(request.HistoryJson) ?? new List<ChatMessage>(); }
                    catch { }
                }

                history.Add(new ChatMessage { Role = "user", Content = request.Command });

                var response = await _service.ExecuteAICommandAsync(request.Command, history);

                history.Add(new ChatMessage { Role = "assistant", Content = response });

                return new JsonResult(new 
                { 
                    success = true, 
                    reply = response, 
                    historyJson = JsonSerializer.Serialize(history) 
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "AJAX Chat Error");
                return new JsonResult(new { success = false, message = $"Server Error: {ex.Message}" });
            }
        }
    }

    public class ChatRequest 
    {
        public string Command { get; set; }
        public string HistoryJson { get; set; }
    }
}
