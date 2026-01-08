using System.ComponentModel.DataAnnotations;

namespace TaskManager.Models
{
    public class Folder
    {
        public int Id { get; set; }

        [Required]
        public string Name { get; set; } = string.Empty;

        public string? Color { get; set; } = "#6366f1"; // Default indigo color

        public string? ImagePath { get; set; } // Optional image path for folder customization

        public int Order { get; set; } = 0;

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        // Navigation property
        public ICollection<TaskItem> Tasks { get; set; } = new List<TaskItem>();

        // User relationship
        public string? UserId { get; set; }
    }
}

