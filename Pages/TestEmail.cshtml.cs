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

        public string ResultLog { get; set; } = "";

        public TestEmailModel(ITaskExecutionerService taskService, IConfiguration configuration, ILogger<TestEmailModel> logger)
        {
            _taskService = taskService;
            _configuration = configuration;
            _logger = logger;
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
                var portStr = _configuration["Email:Port"] ?? "587";
                int.TryParse(portStr, out int port);
                var user = _configuration["Email:Username"];
                
                sb.AppendLine($"Configuration:");
                sb.AppendLine($"  Server: {smtpServer}");
                sb.AppendLine($"  Port: {port}");
                sb.AppendLine($"  User Configured: {!string.IsNullOrEmpty(user)}");
                
                // 2. Network Connectivity Test
                sb.AppendLine($"\nTesting Network Connectivity to {smtpServer}:{port}...");
                try 
                {
                    using (var tcpClient = new TcpClient())
                    {
                        var connectTask = tcpClient.ConnectAsync(smtpServer, port);
                        var timeoutTask = Task.Delay(5000); // 5 second timeout for pure TCP
                        
                        var completedTask = await Task.WhenAny(connectTask, timeoutTask);
                        if (completedTask == timeoutTask)
                        {
                            sb.AppendLine("  ERROR: TCP Connect Timed Out (5s). Firewall likely blocking outbound connection.");
                        }
                        else
                        {
                            await connectTask; // Propagate exceptions
                            sb.AppendLine("  SUCCESS: TCP Connection Established.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"  ERROR: Network Unreachable. Details: {ex.Message}");
                }

                // 3. Attempt Full Send
                if (string.IsNullOrEmpty(toEmail))
                {
                    sb.AppendLine("\nSkipping send: No recipient provided.");
                }
                else
                {
                    sb.AppendLine($"\nAttempting SendEmailAsync to {toEmail}...");
                    await _taskService.SendEmailAsync(toEmail, "Test Email from Diagnostics", "This is a test email to verify deployment connectivity.");
                    sb.AppendLine("  SUCCESS: SendEmailAsync completed without exception.");
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
