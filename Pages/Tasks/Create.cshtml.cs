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
    public class CreateModel : PageModel
    {
        private readonly AppDbContext _context;

        public CreateModel(AppDbContext context)
        {
            _context = context;
        }

        [BindProperty]
        public TaskItem TaskItem { get; set; } = new();

        public SelectList Folders { get; set; } = null!;

        public async Task<IActionResult> OnGetAsync()
        {
            var userId = GetOrSetUserId();

            var folders = await _context.Folders
                .Where(f => f.UserId == userId)
                .OrderBy(f => f.Order)
                .ToListAsync();
            Folders = new SelectList(folders, "Id", "Name", null);
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

            TaskItem.CreatedAt = DateTime.Now; // Auto-set to today's date
            TaskItem.UserId = userId;
            
            // Verify folder belongs to user if specified
            if (TaskItem.FolderId.HasValue)
            {
                var folder = await _context.Folders.FirstOrDefaultAsync(f => f.Id == TaskItem.FolderId && f.UserId == userId);
                if (folder == null)
                    TaskItem.FolderId = null;
            }
            
            var maxOrder = await _context.Tasks
                .Where(t => t.FolderId == TaskItem.FolderId && t.UserId == userId)
                .OrderByDescending(t => t.Order)
                .Select(t => t.Order)
                .FirstOrDefaultAsync();
            TaskItem.Order = maxOrder + 1;

            _context.Tasks.Add(TaskItem);
            await _context.SaveChangesAsync();

            return RedirectToPage("./Board");
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

            // 3. Generate new Guest ID (should already exist if they came from Board/Index, but just in case)
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

