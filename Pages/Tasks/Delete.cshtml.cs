using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Tasks
{
    public class DeleteModel : PageModel
    {
        private readonly AppDbContext _context;

        public DeleteModel(AppDbContext context)
        {
            _context = context;
        }

        [BindProperty]
        public TaskItem TaskItem { get; set; } = default!;

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
            else
            {
                TaskItem = taskitem;
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
            var taskitem = await _context.Tasks.FirstOrDefaultAsync(m => m.Id == id && m.UserId == userId);

            if (taskitem != null)
            {
                TaskItem = taskitem;
                _context.Tasks.Remove(TaskItem);
                await _context.SaveChangesAsync();
            }

            return RedirectToPage("./Board");
        }

        private string GetOrSetUserId()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrEmpty(userId)) return userId;

            const string CookieName = "GuestId";
            if (Request.Cookies.TryGetValue(CookieName, out var guestId)) return guestId;

            return Guid.NewGuid().ToString(); // Fallback (shouldn't happen on delete)
        }
    }
}

