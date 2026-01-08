using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Folders
{
    public class DeleteModel : PageModel
    {
        private readonly AppDbContext _context;

        public DeleteModel(AppDbContext context)
        {
            _context = context;
        }

        [BindProperty]
        public Folder Folder { get; set; } = default!;

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var userId = GetOrSetUserId();

            var folder = await _context.Folders
                .Where(f => f.Id == id && f.UserId == userId)
                .Include(f => f.Tasks.Where(t => t.UserId == userId))
                .FirstOrDefaultAsync();

            if (folder == null)
            {
                return NotFound();
            }
            else
            {
                Folder = folder;
            }
            return Page();
        }

        public async Task<IActionResult> OnPostAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var userId = GetOrSetUserId();

            var folder = await _context.Folders
                .Where(f => f.Id == id && f.UserId == userId)
                .Include(f => f.Tasks.Where(t => t.UserId == userId))
                .FirstOrDefaultAsync();

            if (folder != null)
            {
                // Remove folder from tasks (set FolderId to null)
                foreach (var task in folder.Tasks)
                {
                    task.FolderId = null;
                }
                
                Folder = folder;
                _context.Folders.Remove(Folder);
                await _context.SaveChangesAsync();
            }

            return RedirectToPage("../Tasks/Board");
        }

        private string GetOrSetUserId()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrEmpty(userId)) return userId;

            const string CookieName = "GuestId";
            if (Request.Cookies.TryGetValue(CookieName, out var guestId)) return guestId;

            return Guid.NewGuid().ToString();
        }
    }
}

