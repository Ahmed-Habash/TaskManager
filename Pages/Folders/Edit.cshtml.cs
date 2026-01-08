using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using TaskManager.Models;

namespace TaskManager.Pages.Folders
{
    public class EditModel : PageModel
    {
        private readonly AppDbContext _context;
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<EditModel> _logger;

        public EditModel(AppDbContext context, IWebHostEnvironment environment, ILogger<EditModel> logger)
        {
            _context = context;
            _environment = environment;
            _logger = logger;
        }

        [BindProperty]
        public Folder Folder { get; set; } = default!;

        [BindProperty]
        public IFormFile? ImageFile { get; set; }

        [BindProperty]
        public bool RemoveImage { get; set; }

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var userId = GetOrSetUserId();

            var folder = await _context.Folders.FirstOrDefaultAsync(m => m.Id == id && m.UserId == userId);
            if (folder == null)
            {
                return NotFound();
            }
            Folder = folder;
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            var userId = GetOrSetUserId();

            if (!ModelState.IsValid)
            {
                return Page();
            }

            // Verify folder belongs to user
            var existingFolder = await _context.Folders.FirstOrDefaultAsync(f => f.Id == Folder.Id && f.UserId == userId);
            if (existingFolder == null)
                return NotFound();

            // Handle image upload
            if (ImageFile != null && ImageFile.Length > 0)
            {
                try
                {
                    // Delete old image if it exists
                    if (!string.IsNullOrEmpty(Folder.ImagePath))
                    {
                        var oldImagePath = Path.Combine(_environment.WebRootPath, Folder.ImagePath.TrimStart('/'));
                        if (System.IO.File.Exists(oldImagePath))
                        {
                            System.IO.File.Delete(oldImagePath);
                        }
                    }

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
                    _logger.LogError(ex, "Error uploading file for folder edit");
                    ModelState.AddModelError("ImageFile", "An error occurred while uploading the file. Please try again.");
                    // Re-fetch existing data if needed, or just return page to show error
                    return Page();
                }
            }

            // Preserve UserId and ensure folder belongs to user
            existingFolder.Name = Folder.Name;
            existingFolder.Color = Folder.Color;

            if (RemoveImage)
            {
                // Delete old image if exists
                if (!string.IsNullOrEmpty(existingFolder.ImagePath))
                {
                    var oldImagePath = Path.Combine(_environment.WebRootPath, existingFolder.ImagePath.TrimStart('/'));
                    if (System.IO.File.Exists(oldImagePath))
                    {
                        System.IO.File.Delete(oldImagePath);
                    }
                }
                existingFolder.ImagePath = null;
            }
            else if (!string.IsNullOrEmpty(Folder.ImagePath))
            {
                existingFolder.ImagePath = Folder.ImagePath;
            }

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!FolderExists(Folder.Id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return RedirectToPage("../Tasks/Board");
        }

        private bool FolderExists(int id)
        {
            return _context.Folders.Any(e => e.Id == id);
        }

        private string GetOrSetUserId()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrEmpty(userId)) return userId;

            const string CookieName = "GuestId";
            if (Request.Cookies.TryGetValue(CookieName, out var guestId)) return guestId;

            guestId = Guid.NewGuid().ToString();
            var cookieOptions = new CookieOptions { Expires = DateTime.Now.AddYears(1), HttpOnly = true, IsEssential = true, SameSite = SameSiteMode.Lax };
            Response.Cookies.Append(CookieName, guestId, cookieOptions);
            return guestId;
        }
    }
}
