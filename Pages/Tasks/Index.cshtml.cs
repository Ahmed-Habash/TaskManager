using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Tasks
{
    public class IndexModel : PageModel
    {
        private readonly AppDbContext _context;

        public IndexModel(AppDbContext context)
        {
            _context = context;
        }

        // Public property to be used in Razor page
        public IList<TaskItem> Tasks { get; set; } = default!;

        // Load tasks from database
        public async Task OnGetAsync()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null)
                return;

            Tasks = await _context.Tasks
                .Where(t => t.UserId == userId)
                .Include(t => t.Folder)
                .OrderByDescending(t => t.CreatedAt)
                .ToListAsync();
        }
    }
}
