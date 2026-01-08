using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Folders
{
    public class CreateModel : PageModel
    {
        private readonly AppDbContext _context;
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<CreateModel> _logger;

        public CreateModel(AppDbContext context, IWebHostEnvironment environment, ILogger<CreateModel> logger)
        {
            _context = context;
            _environment = environment;
            _logger = logger;
        }

        [BindProperty]
        public Folder Folder { get; set; } = new();

        [BindProperty]
        public IFormFile? ImageFile { get; set; }

        public IActionResult OnGet()
        {
            var userId = GetOrSetUserId();
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            // Handle image upload
            if (ImageFile != null && ImageFile.Length > 0)
            {
                try 
                {
                    var uploadsFolder = Path.Combine(_environment.WebRootPath, "uploads", "folders");
                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }

                    var uniqueFileName = Guid.NewGuid().ToString() + Path.GetExtension(ImageFile.FileName);
                    var filePath = Path.Combine(uploadsFolder, uniqueFileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await ImageFile.CopyToAsync(fileStream);
                    }

                    Folder.ImagePath = $"/uploads/folders/{uniqueFileName}";
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error uploading file for folder creation");
                    ModelState.AddModelError("ImageFile", "An error occurred while uploading the file. Please try again.");
                    return Page();
                }
            }

            var userId = GetOrSetUserId();

            Folder.CreatedAt = DateTime.Now;
            Folder.UserId = userId;
            var maxOrder = await _context.Folders
                .Where(f => f.UserId == userId)
                .OrderByDescending(f => f.Order)
                .Select(f => f.Order)
                .FirstOrDefaultAsync();
            Folder.Order = maxOrder + 1;

            _context.Folders.Add(Folder);
            await _context.SaveChangesAsync();

            return RedirectToPage("../Tasks/Board");
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
