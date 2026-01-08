using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Tasks
{
    public class BoardModel : PageModel
    {
        private readonly AppDbContext _context;
        private readonly TaskManager.Services.IAIService _aiService;

        public BoardModel(AppDbContext context, TaskManager.Services.IAIService aiService)
        {
            _context = context;
            _aiService = aiService;
        }

        public List<Folder> Folders { get; set; } = new();
        public List<TaskItem> UnfolderredTasks { get; set; } = new();
        public string Filter { get; set; } = "all"; // all, pending, completed

        public async Task OnGetAsync(string filter = "all")
        {
            Filter = filter;
            var userId = GetOrSetUserId();

            Folders = await _context.Folders
                .Where(f => f.UserId == userId)
                .OrderBy(f => f.Order)
                .Include(f => f.Tasks.Where(t => t.UserId == userId).OrderBy(t => t.Order))
                .ToListAsync();

            UnfolderredTasks = await _context.Tasks
                .Where(t => t.FolderId == null && t.UserId == userId)
                .OrderBy(t => t.Order)
                .ToListAsync();

            // Apply filter
            if (filter == "pending")
            {
                foreach (var folder in Folders)
                {
                    folder.Tasks = folder.Tasks.Where(t => !t.IsCompleted).ToList();
                }
                UnfolderredTasks = UnfolderredTasks.Where(t => !t.IsCompleted).ToList();
            }
            else if (filter == "completed")
            {
                foreach (var folder in Folders)
                {
                    folder.Tasks = folder.Tasks.Where(t => t.IsCompleted).ToList();
                }
                UnfolderredTasks = UnfolderredTasks.Where(t => t.IsCompleted).ToList();
            }
        }

        public async Task<IActionResult> OnPostUpdateTask([FromBody] UpdateTaskRequest request)
        {
            if (request == null || request.TaskIds == null)
                return BadRequest();

            var userId = GetOrSetUserId();
            
            // Fetch all involved tasks to minimize DB roundtrips (and ensure ownership)
            var tasksToUpdate = await _context.Tasks
                .Where(t => request.TaskIds.Contains(t.Id) && t.UserId == userId)
                .ToListAsync();

            // Create a dictionary for fast lookup
            var taskDict = tasksToUpdate.ToDictionary(t => t.Id);

            // Iterate through the ID list from client to apply order
            for (int i = 0; i < request.TaskIds.Count; i++)
            {
                var id = request.TaskIds[i];
                if (taskDict.TryGetValue(id, out var task))
                {
                    task.FolderId = request.FolderId;
                    task.Order = i; // Strict 0-based ordering
                }
            }

            await _context.SaveChangesAsync();
            return new JsonResult(new { success = true });
        }

        public async Task<IActionResult> OnPostToggleComplete([FromBody] ToggleCompleteRequest request)
        {
            if (request == null)
                return BadRequest();

            var userId = GetOrSetUserId();
            var task = await _context.Tasks.FirstOrDefaultAsync(t => t.Id == request.TaskId && t.UserId == userId);
            if (task == null) return NotFound();

            task.IsCompleted = !task.IsCompleted;
            await _context.SaveChangesAsync();

            return new JsonResult(new { success = true, isCompleted = task.IsCompleted });
        }

        public async Task<IActionResult> OnPostAskAI([FromBody] AskAIRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Prompt))
                return BadRequest("Prompt is required.");

            var userId = GetOrSetUserId();

            // Fetch tasks for context
            var tasks = await _context.Tasks
                .Where(t => t.UserId == userId)
                .ToListAsync();

            TaskItem? focusedTask = null;
            if (request.FocusedTaskId.HasValue)
            {
                focusedTask = tasks.FirstOrDefault(t => t.Id == request.FocusedTaskId.Value);
            }

            var response = await _aiService.GetAssistanceAsync(request.Prompt, tasks, request.History, focusedTask);

            return new JsonResult(new { success = true, reply = response });
        }

        private string GetOrSetUserId()
        {
            // 1. Check if user is logged in
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrEmpty(userId))
            {
                return userId;
            }

            // 2. Check for Guest Cookie
            const string CookieName = "GuestId";
            if (Request.Cookies.TryGetValue(CookieName, out var guestId))
            {
                return guestId;
            }

            // 3. Generate new Guest ID
            guestId = Guid.NewGuid().ToString();
            
            // Set cookie options
            var cookieOptions = new CookieOptions
            {
                Expires = DateTime.Now.AddYears(1),
                HttpOnly = true,
                IsEssential = true, // Required for GDPR if applicable, though this is a utility app
                SameSite = SameSiteMode.Lax
            };

            Response.Cookies.Append(CookieName, guestId, cookieOptions);
            return guestId;
        }

        public class ToggleCompleteRequest
        {
            public int TaskId { get; set; }
        }

        public class UpdateTaskRequest
        {
            public int? FolderId { get; set; }
            public List<int> TaskIds { get; set; } = new();
        }

        public class AskAIRequest
        {
            public string Prompt { get; set; } = string.Empty;
            public int? FocusedTaskId { get; set; }
            public List<TaskManager.Services.ChatMessage>? History { get; set; }
        }
    }
}
