using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Tasks
{
    public class EditModel : PageModel
    {
        private readonly AppDbContext _context;

        public EditModel(AppDbContext context)
        {
            _context = context;
        }

        [BindProperty]
        public TaskItem TaskItem { get; set; } = default!;

        public SelectList Folders { get; set; } = null!;

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var userId = GetOrSetUserId();

            var taskitem = await _context.Tasks.FirstOrDefaultAsync(m => m.Id == id && m.UserId == userId);
            if (taskitem == null)
            {
                return NotFound();
            }
            TaskItem = taskitem;

            var folders = await _context.Folders
                .Where(f => f.UserId == userId)
                .OrderBy(f => f.Order)
                .ToListAsync();
            Folders = new SelectList(folders, "Id", "Name", TaskItem.FolderId);

            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            var userId = GetOrSetUserId();

            if (!ModelState.IsValid)
            {
                var folders = await _context.Folders
                    .Where(f => f.UserId == userId)
                    .OrderBy(f => f.Order)
                    .ToListAsync();
                Folders = new SelectList(folders, "Id", "Name", TaskItem.FolderId);
                return Page();
            }

            // Verify task belongs to user
            var existingTask = await _context.Tasks.FirstOrDefaultAsync(t => t.Id == TaskItem.Id && t.UserId == userId);
            if (existingTask == null)
                return NotFound();

            // Ensure task belongs to user and preserve UserId
            existingTask.Title = TaskItem.Title;
            existingTask.Description = TaskItem.Description;
            existingTask.Priority = TaskItem.Priority;
            existingTask.Category = TaskItem.Category;
            existingTask.IsCompleted = TaskItem.IsCompleted;
            existingTask.DueDate = TaskItem.DueDate;
            
            // Verify folder belongs to user if specified
            if (TaskItem.FolderId.HasValue)
            {
                var folder = await _context.Folders.FirstOrDefaultAsync(f => f.Id == TaskItem.FolderId && f.UserId == userId);
                if (folder == null)
                    existingTask.FolderId = null;
                else
                    existingTask.FolderId = TaskItem.FolderId;
            }
            else
            {
                existingTask.FolderId = null;
            }

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!TaskItemExists(TaskItem.Id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return RedirectToPage("./Board");
        }

        private bool TaskItemExists(int id)
        {
            return _context.Tasks.Any(e => e.Id == id);
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

            // 3. Generate new guest ID (fallback)
            guestId = Guid.NewGuid().ToString();
            var cookieOptions = new CookieOptions
            {
                Expires = DateTime.Now.AddYears(1),
                HttpOnly = true,
                IsEssential = true,
                SameSite = SameSiteMode.Lax
            };
            Response.Cookies.Append(CookieName, guestId, cookieOptions);
            return guestId;
        }
    }
}

