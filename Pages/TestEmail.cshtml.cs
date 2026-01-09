using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using System.Net.Sockets;
using System.Text;
using TaskManager.Services;

namespace TaskManager.Pages
{
    public class TestEmailModel : PageModel
    {
        private readonly ITaskExecutionerService _taskService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<TestEmailModel> _logger;
        private readonly IWebHostEnvironment _environment;

        public string ResultLog { get; set; } = "";

        public TestEmailModel(ITaskExecutionerService taskService, IConfiguration configuration, ILogger<TestEmailModel> logger, IWebHostEnvironment environment)
        {
            _taskService = taskService;
            _configuration = configuration;
            _logger = logger;
            _environment = environment;
        }

        public void OnGet()
        {
        }

        public async Task OnPostAsync(string toEmail)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"--- Starting Email Diagnostic at {DateTime.Now} ---");
            
            try
            {
                // 1. Check Configuration
                var smtpServer = _configuration["Email:SmtpServer"] ?? "smtp.gmail.com";
                var portStr = _configuration["Email:Port"] ?? "465";
                int.TryParse(portStr, out int port);
                var user = _configuration["Email:Username"];
                
                sb.AppendLine($"Configuration:");
                sb.AppendLine($"  Server: {smtpServer}");
                sb.AppendLine($"  Configured Port: {port}");
                sb.AppendLine($"  User Configured (appsettings/env): {!string.IsNullOrEmpty(user)}");
                
                // 2. Multi-Port Connectivity Test
                int[] portsToTest = { 465, 587, 2525 };
                sb.AppendLine($"\nTesting Outbound Connectivity to {smtpServer}:");
                foreach (var p in portsToTest)
                {
                    try 
                    {
                        using (var tcpClient = new TcpClient())
                        {
                            var connectTask = tcpClient.ConnectAsync(smtpServer, p);
                            var timeoutTask = Task.Delay(3000); 
                            
                            var completedTask = await Task.WhenAny(connectTask, timeoutTask);
                            if (completedTask == timeoutTask)
                            {
                                sb.AppendLine($"  - Port {p}: ERROR (Timed Out - Likely Blocked)");
                            }
                            else
                            {
                                await connectTask; 
                                sb.AppendLine($"  - Port {p}: SUCCESS (Connection Established)");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"  - Port {p}: ERROR ({ex.Message})");
                    }
                }

                // 3. File System Write Check (for logs/exports)
                sb.AppendLine("\nTesting File System Write Permissions (wwwroot/exports):");
                try
                {
                    var testPath = System.IO.Path.Combine(_environment.WebRootPath, "exports", "write_test.txt");
                    var dir = System.IO.Path.GetDirectoryName(testPath);
                    if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir!);
                    await System.IO.File.WriteAllTextAsync(testPath, $"Test write at {DateTime.Now}");
                    sb.AppendLine("  SUCCESS: Can write to exports directory.");
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"  ERROR: Cannot write to exports. Details: {ex.Message}");
                    sb.AppendLine("  (Note: Many cloud hosts have read-only filesystems by default).");
                }

                // 4. Attempt Full Send
                if (string.IsNullOrEmpty(toEmail))
                {
                    sb.AppendLine("\nSkipping send: No recipient provided.");
                }
                else
                {
                    sb.AppendLine($"\nAttempting Full SendEmailAsync to {toEmail}...");
                    try
                    {
                        await _taskService.SendEmailAsync(toEmail, "Test Email from Diagnostics", "This is a test email to verify deployment connectivity.");
                        sb.AppendLine("  SUCCESS: SendEmailAsync completed without exception.");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"  FAILED: {ex.Message}");
                    }
                }

            }
            catch (Exception ex)
            {
                sb.AppendLine($"\nFATAL ERROR during execution: {ex.Message}");
                sb.AppendLine($"Stack Trace: {ex.StackTrace}");
            }

            sb.AppendLine("--- End Diagnostic ---");
            ResultLog = sb.ToString();
        }
    }
}
