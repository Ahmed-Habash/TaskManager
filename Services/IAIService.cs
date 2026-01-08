using System.Collections.Generic;
using System.Threading.Tasks;
using TaskManager.Models;

namespace TaskManager.Services
{
    public class ChatMessage
    {
        public string Role { get; set; } = string.Empty; // "user" or "assistant"
        public string Content { get; set; } = string.Empty;
    }

    public interface IAIService
    {
        /// <summary>
        /// Sends a prompt to the AI with context about the current tasks and conversation history.
        /// </summary>
        /// <param name="userPrompt">The user's question or command.</param>
        /// <param name="contextTasks">The list of tasks to provide as context.</param>
        /// <param name="history">Optional history of previous messages.</param>
        /// <param name="focusedTask">Optional: A specific task the user is asking about.</param>
        /// <returns>The AI's response.</returns>
        Task<string> GetAssistanceAsync(string userPrompt, IEnumerable<TaskItem> contextTasks, IEnumerable<ChatMessage>? history = null, TaskItem? focusedTask = null);

        /// <summary>
        /// Analyzes the user's prompt and determines the action to take.
        /// Returns a JSON string describing the action and parameters.
        /// </summary>
        Task<string> DetermineActionAsync(string userPrompt, IEnumerable<ChatMessage> history);
        Task<string> SimulateSearchAsync(string query);
        Task<string> SummarizeSearchResultsAsync(string query, string rawResults);
    }
}
