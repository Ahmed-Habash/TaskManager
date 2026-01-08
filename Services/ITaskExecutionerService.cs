using System.Threading.Tasks;

namespace TaskManager.Services
{
    public interface ITaskExecutionerService
    {
        Task SendEmailAsync(string to, string subject, string body, string? attachmentPath = null, string? username = null, string? password = null);
        Task<string> CreateFileAsync(string fileName, string content, string type = "word", string primaryColor = "1C1C1C", string secondaryColor = "0078D7", string imageKeyword = "");
        Task<string> ExecuteAICommandAsync(string userPrompt, List<ChatMessage> history);
    }
}
