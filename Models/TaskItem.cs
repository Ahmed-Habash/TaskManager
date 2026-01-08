using System;
using System.ComponentModel.DataAnnotations;

namespace TaskManager.Models
{
    public enum TaskCategory
    {
        ThisDay,
        ThisWeek,
        ThisMonth,
        ThisYear
    }

    public enum TaskPriority
    {
        Low,
        Medium,
        High
    }

    public class TaskItem
    {
        public int Id { get; set; }

        [Required]
        public string Title { get; set; } = string.Empty;

        public string? Description { get; set; }

        public bool IsCompleted { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public DateTime? DueDate { get; set; } // Optional completion target date

        // New fields
        public TaskCategory Category { get; set; } = TaskCategory.ThisDay;
        public TaskPriority Priority { get; set; } = TaskPriority.Medium;

        // Optional: Order for drag-and-drop sorting
        public int Order { get; set; } = 0;

        // Optional folder relationship
        public int? FolderId { get; set; }
        public Folder? Folder { get; set; }

        // User relationship
        public string? UserId { get; set; }
    }
}
