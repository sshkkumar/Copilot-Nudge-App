

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Azure.Identity;

using Microsoft.Graph.Models;
using System.Text;


namespace Copilot_Nudge_App
{
    public partial class CopilotNudgeApp : Form
    {

        // Explicit widths (adjust as needed)
        private const int LabelWidth = 200;
        private const int LabelHeight = 48;

        private const int TextWidth = 1200;
        private const int TextHeight = 48;

        private const int BrowseButtonWidth = 200;
        private const int BrowseButtonHeight = 48;

        private const int ActionButtonWidth = 120;
        private const int ActionButtonHeight = 48;

        private const int ProgressWidth = 200;
        private const int ProgressHeight = 48;

        private const int StatusHeight = 280;

        // Task state
        private CancellationTokenSource? _cts;
        private Task? _runningTask;


        // Load configuration
        IConfigurationRoot config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: false).Build();



        string tenantId = String.Empty;
        string clientId = String.Empty; 


        // UI Controls
        private readonly Label lblUsers = new() 
        {
            Text = "Users list:",
            AutoSize = false, // Disable AutoSize so Size takes effect
            Size = new System.Drawing.Size(LabelWidth, LabelHeight),
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        };
        private readonly TextBox txtUsers = new() 
        { 
            ReadOnly = true,
            Size = new System.Drawing.Size(TextWidth, TextHeight),
            Anchor = AnchorStyles.Left

        };
        private readonly Button btnBrowseUsers = new() 
        { 
            Text = "Browse...",
            Size = new System.Drawing.Size(BrowseButtonWidth, BrowseButtonHeight),
            Anchor = AnchorStyles.Right

        };

        private readonly Label lblCard = new() 
        { 
            Text = "Adaptive card:",
            AutoSize = false,
            Size = new System.Drawing.Size(LabelWidth, LabelHeight),
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        };
        private readonly TextBox txtCard = new() 
        { 
            ReadOnly = true,
            Size = new System.Drawing.Size(TextWidth, TextHeight),
            Anchor = AnchorStyles.Left

        };
        private readonly Button btnBrowseCard = new() 
        { 
            Text = "Browse...",
            Size = new System.Drawing.Size(BrowseButtonWidth, BrowseButtonHeight),
            Anchor = AnchorStyles.Right

        };

        private readonly Button btnRun = new() 
        { 
            Text = "Run",
            AutoSize = false,
            Size = new System.Drawing.Size(ActionButtonWidth, ActionButtonHeight),
            Anchor = AnchorStyles.Left

        };
        private readonly Button btnCancel = new() 
        { 
            Text = "Cancel",
            AutoSize = false,
            Enabled = false,
            Size = new System.Drawing.Size(ActionButtonWidth, ActionButtonHeight),
            Anchor = AnchorStyles.Left

        };

        private readonly ProgressBar progressBar = new() 
        {
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            AutoSize = false,
            Size = new System.Drawing.Size(ProgressWidth, ProgressHeight),
            Anchor = AnchorStyles.Right

        };
        private readonly TextBox txtStatus = new() 
        {
            //Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true 

            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Font = new System.Drawing.Font("Consolas", 10),
            AutoSize = false,
            Height = StatusHeight,
            Dock = DockStyle.Fill

        };

        

        public CopilotNudgeApp()
        {
            InitializeComponent();
        }

        private void CopilotNudgeApp_Load(object sender, EventArgs e)
        {
            var config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: false)
            .Build();

            tenantId = config["TenantId"]!;
            clientId = config["ClientId"]!;
           

            StartPosition = FormStartPosition.CenterScreen;
            Width = 1800;
            Height = 822;
            MinimumSize = new System.Drawing.Size(1800, 822);


            // Layout with TableLayoutPanel for responsiveness
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 6,
                Padding = new Padding(50),
            };


            // Fixed column widths to avoid jumbled layout
            // [0] Label column — fixed
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, LabelWidth + 10));
            // [1] Text field column — fixed
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, TextWidth + 10));
            // [2] Button/Progress column — fixed
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Math.Max(BrowseButtonWidth, ProgressWidth) + 16));

            // Row styles: compact header rows + status area
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70)); // Users row
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70)); // Card row
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48)); // Run/Cancel/Progress row
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); // "Status:" header
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Status textbox fill
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10)); // Bottom padding

            SuspendLayout();
            layout.SuspendLayout();



            // Add controls (row by row)
            // Row 0: Users list
            layout.Controls.Add(lblUsers, 0, 0);
            layout.Controls.Add(txtUsers, 1, 0);
            layout.Controls.Add(btnBrowseUsers, 2, 0);

            // Row 1: Adaptive card
            layout.Controls.Add(lblCard, 0, 1);
            layout.Controls.Add(txtCard, 1, 1);
            layout.Controls.Add(btnBrowseCard, 2, 1);

            // Row 2: Run / Cancel / Progress
            layout.Controls.Add(btnRun, 0, 2);
            layout.Controls.Add(btnCancel, 1, 2);
            layout.Controls.Add(progressBar, 2, 2);


            // Row 3: Status header
            var lblStatus = new Label
            {
                Text = "Status:",
                AutoSize = false,
                Size = new System.Drawing.Size(LabelWidth, LabelHeight),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Left
            };

            layout.Controls.Add(lblStatus, 0, 3);




            // Row 4: Status textbox (span all columns)
            layout.SetColumnSpan(txtStatus, 3);
            layout.Controls.Add(txtStatus, 0, 4);

            Controls.Add(layout);

            layout.ResumeLayout(performLayout: true);
            ResumeLayout(performLayout: true);


            // Wire events
            //btnBrowseUsers.Click += (_, __) => BrowseFile(txtUsers, "Select users list", "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*");
            btnBrowseUsers.Click += (_, __) => BrowseFile(txtUsers, "Select users list", "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*");
            btnBrowseCard.Click += (_, __) => BrowseFile(txtCard, "Select adaptive card JSON", "JSON (*.json)|*.json|All files (*.*)|*.*");
            btnRun.Click += async (_, __) => await OnRunClickedAsync();
            btnCancel.Click += (_, __) => OnCancelClicked();

        }

        private void BrowseFile(TextBox target, string title, string filter)
        {
            using var ofd = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                Multiselect = false
            };
            var result = ofd.ShowDialog(this);
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
            {
                target.Text = ofd.FileName;
                //AppendStatus($"Selected: {ofd.FileName}");
            }
        }


        private async Task OnRunClickedAsync()
        {
            // Validation
            if (string.IsNullOrWhiteSpace(txtUsers.Text) || !File.Exists(txtUsers.Text))
            {
                MessageBox.Show(this, "Please select a valid Users list file.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(txtCard.Text) || !File.Exists(txtCard.Text))
            {
                MessageBox.Show(this, "Please select a valid Adaptive Card JSON file.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Reset UI
            btnRun.Enabled = false;
            btnCancel.Enabled = true;
            progressBar.Value = 0;
            txtStatus.Clear();
            AppendStatus("Starting job...");
            LocalLogger.Log($"Starting Nudge app for Tenant: {tenantId} and ClientID used is: {clientId}");

            // Prepare cancellation
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            // Kick off non-blocking work
            _runningTask = RunWorkAsync(txtUsers.Text, txtCard.Text, token);

            try
            {
                await _runningTask;
            }
            catch (OperationCanceledException)
            {
                AppendStatus("Job cancelled.");
                LocalLogger.Log("Job cancelled.");
            }
            catch (Exception ex)
            {
                AppendStatus($"ERROR: {ex.Message}");
                LocalLogger.LogException($"ERROR in running task", ex);
            }
            finally
            {
                btnRun.Enabled = true;
                btnCancel.Enabled = false;
                _cts?.Dispose();
                _cts = null;
                _runningTask = null;
            }
        }

        private void OnCancelClicked()
        {
            if (_cts is not null && !_cts.IsCancellationRequested)
            {
                AppendStatus("Requesting cancellation...");
                _cts.Cancel();
                btnCancel.Enabled = false;
            }
        }

        /// <summary>
        /// Execute the work
        /// It runs on a thread-pool thread via async/await and reports progress and status safely.
        /// </summary>
        private async Task RunWorkAsync(string usersPath, string cardPath, CancellationToken ct)
        {
            // Simulate initial I/O
            //await Task.Delay(300, ct);

            AppendStatus($"Users: {usersPath}");
            AppendStatus($"Card : {cardPath}");
            AppendStatus("Validating inputs...");
            JsonDocument doc;

            //Validate Adaptive Card JSON
            try
            {

                using var fs = File.OpenRead(cardPath);
                doc = await JsonDocument.ParseAsync(fs, cancellationToken: ct);
                // Optional: Validate schema fields you care about (placeholder)
                AppendStatus("Adaptive Card JSON loaded.");
            }
            catch (JsonException jx)
            {
                AppendStatus($"Adaptive Card JSON error: {jx.Message}");
                throw;
            }

            //Read list of users to send the adaptive card to
            AppendStatus("Reading Users list...");
            // Parse emails from CSV
            var userEmails = ParseUserEmailsFromCsv(usersPath)
                .Select(e => e.Trim().ToLowerInvariant())
                .Distinct()
                .ToList();
            //var userEmails = ParseUserEmailsFromCsv(usersPath);

            AppendStatus($"Found {userEmails.Count} unique emails.");
            LocalLogger.Log($"Found {userEmails.Count} unique emails.");
            int totalSteps = userEmails.Count;

            //Authenticate with Graph API
            var graphClient = AuthenticateWithGraph();

            // Resolve signed-in user (service account) to ensure it’s valid and licensed
            var me = await graphClient.Me.GetAsync();
            AppendStatus($"Signed in as: {me?.UserPrincipalName}");
            LocalLogger.Log($"Signed in as: {me?.UserPrincipalName}");


            // Resolve the service account user ID (needed for chat membership)
            // var serviceUserId = await ResolveUserIdByEmailAsync(graphClient, username);

            for (int userCount = 0; userCount < userEmails.Count; userCount++)
            {
                try
                {
                    ct.ThrowIfCancellationRequested();
                    AppendStatus($"Processing {userEmails[userCount]}");
                    LocalLogger.Log($"Processing {userEmails[userCount]}");

                    // Resolve target user ID by email/UPN
                    var targetUserId = await ResolveUserIdByEmailAsync(graphClient, userEmails[userCount]);                    

                    // Chat ID to post the adaptive card to
                    var chatId = await GetOrCreateOneOnOneChatAsync(graphClient, me!.Id, targetUserId);

                    // Pre-generate a stable attachment Id so retries don't produce duplicate cards
                    var attachmentId = Guid.NewGuid().ToString("D");


                    // Send adaptive card into the chat
                    var sent = await SendAdaptiveCardToChatAsync(graphClient, chatId, doc, attachmentId, ct);

                    AppendStatus($"Sent card to {userEmails[userCount]} (messageId: {sent?.Id})");
                    LocalLogger.Log($"Sent card to {userEmails[userCount]} (messageId: {sent?.Id})");

                    //await SendAdaptiveCardToChatAsync(graphClient, chatId, doc,ct);

                    SetProgress((int)Math.Round(userCount+1 / (double)totalSteps * 100));

                }

                catch (ServiceException gex)
                {
                    AppendStatus($"Graph error for {userEmails[userCount]}:  {(int)gex.ResponseStatusCode} {gex.Message}");
                    LocalLogger.LogException($"Graph error for {userEmails[userCount]}", gex);
                }

                catch (Exception ex)
                {
                    AppendStatus($"Failed for {userEmails[userCount]}: {ex.Message}");
                    LocalLogger.LogException($"Failed for {userEmails[userCount]}", ex);                    
                }
            }

            AppendStatus("All steps completed successfully.");
            LocalLogger.Log("All steps completed successfully.");

            SetProgress(100);
        }

        private GraphServiceClient AuthenticateWithGraph()
        {
            // Delegated scopes (the app must have these delegated perms with admin consent)
            var scopes = new[]
            {
                "User.Read",
                "User.Read.All",
                "Chat.ReadWrite",
                "ChatMessage.Send"
            };


            // Opens a system browser; supports MFA/CA
            var credential = new InteractiveBrowserCredential(
                new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId
                });


            var graph = new GraphServiceClient(credential, scopes);
            return graph;
        }



        /// <summary>
        /// Posts an Adaptive Card to a chat with retry (throttling-safe), preserving a stable attachment Id.
        /// </summary>
        public static async Task<ChatMessage> SendAdaptiveCardToChatAsync(
            GraphServiceClient graph,
            string chatId,
            JsonDocument cardDoc,
            string? attachmentId = null,
            CancellationToken ct = default)
        {
            if (graph is null) throw new ArgumentNullException(nameof(graph));
            if (string.IsNullOrWhiteSpace(chatId)) throw new ArgumentException("chatId is required.", nameof(chatId));
            if (cardDoc is null) throw new ArgumentNullException(nameof(cardDoc));

            var root = cardDoc.RootElement;

            // Minimal Adaptive Card validation
            if (!root.TryGetProperty("type", out var typeProp) ||
                !string.Equals(typeProp.GetString(), "AdaptiveCard", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Adaptive Card must have \"type\": \"AdaptiveCard\".");
            }
            if (!root.TryGetProperty("version", out var versionProp) ||
                string.IsNullOrWhiteSpace(versionProp.GetString()))
            {
                throw new InvalidOperationException("Adaptive Card must include a \"version\" (e.g., \"1.4\").");
            }

            var json = root.GetRawText();

            var byteLen = Encoding.UTF8.GetByteCount(json);
            const int MaxBytes = 48 * 1024; // safety cap for Teams card payloads
            if (byteLen > MaxBytes)
                throw new InvalidOperationException($"Adaptive Card payload too large ({byteLen} bytes).");

            var id = string.IsNullOrWhiteSpace(attachmentId) ? Guid.NewGuid().ToString("D") : attachmentId;

            var message = new ChatMessage
            {
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    // IMPORTANT: attachment marker must reference the attachment Id
                    Content = $@"<p></p><attachment id=""{id}""></attachment>"
                },
                Attachments = new System.Collections.Generic.List<ChatMessageAttachment>
                {
                    new ChatMessageAttachment
                    {
                        Id          = id,
                        ContentType = "application/vnd.microsoft.card.adaptive",
                        Content     = json
                    }
                }
            };

            // Wrap the actual POST in the retry helper:
            return await GraphRetry.ExecuteWithGraphRetryAsync(
                () => graph.Chats[chatId].Messages.PostAsync(message, cancellationToken: ct),
                ct: ct,
                maxAttempts: 6,
                baseDelay: TimeSpan.FromMilliseconds(500),
                maxDelay: TimeSpan.FromSeconds(20));
        }


        // --- Graph helpers ---
        static async Task<string> ResolveUserIdByEmailAsync(GraphServiceClient graph, string emailOrUpn)
        {
            // Try matching mail or UPN; require directory read
            var users = await graph.Users.GetAsync(req =>
            {
                req.QueryParameters.Top = 1;
                req.QueryParameters.Select = new[] { "id", "mail", "userPrincipalName", "displayName" };
                req.QueryParameters.Filter = $"(mail eq '{emailOrUpn}') or (userPrincipalName eq '{emailOrUpn}')";
            });

            var user = users?.Value?.FirstOrDefault();
            if (user?.Id is null)
                throw new InvalidOperationException($"User not found: {emailOrUpn}");

            return user.Id;
        }

        static async Task<string> GetOrCreateOneOnOneChatAsync(GraphServiceClient graph, string serviceUserId, string targetUserId)
        {

            if (string.IsNullOrWhiteSpace(serviceUserId)) throw new ArgumentException("serviceUserId is required.");
            if (string.IsNullOrWhiteSpace(targetUserId)) throw new ArgumentException("targetUserId is required.");
            if (serviceUserId == targetUserId) throw new InvalidOperationException("1:1 chat must have two distinct users.");

            var memberA = BuildMemberByUserId(serviceUserId);
            var memberB = BuildMemberByUserId(targetUserId);

            var chatToCreate = new Chat
            {
                ChatType = ChatType.OneOnOne,
                Members = new List<ConversationMember> { memberA, memberB }
            };

            // POST /chats — for OneOnOne, Graph will return the existing chat if it already exists or create a new one.
            var created = await graph.Chats.PostAsync(chatToCreate);
            if (created?.Id is null)
                throw new InvalidOperationException("Failed to create or locate 1:1 chat.");

            return created.Id;

        }


        // <summary>
        /// Build member by user id to create the chat id
        /// </summary>
        static AadUserConversationMember BuildMemberByUserId(string userId)
        {
            // IMPORTANT: Use "user@odata.bind" with the canonical v1.0 users URL.
            return new AadUserConversationMember
            {
                Roles = new List<string> { "owner" },
                AdditionalData = new Dictionary<string, object>
                {
                    // Bind to the user resource by Id; UPN also works in many tenants, but Id is safest.
                    ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{userId}')"
                }
            };
        }

        ///// <summary>
        ///// Sends an Adaptive Card to a chat using a JsonDocument payload.
        ///// Validates minimal schema and enforces content size limits.
        ///// </summary>
        //public static async Task SendAdaptiveCardToChatAsync(GraphServiceClient graph, string chatId, JsonDocument cardDoc, CancellationToken ct = default)
        //{
        //    if (graph is null) throw new ArgumentNullException(nameof(graph));
        //    if (string.IsNullOrWhiteSpace(chatId)) throw new ArgumentException("chatId is required.", nameof(chatId));
        //    if (cardDoc is null) throw new ArgumentNullException(nameof(cardDoc));

        //    // --- Minimal schema checks (optional but recommended) ---
        //    var rootElement = cardDoc.RootElement;

        //    if (!rootElement.TryGetProperty("type", out var typeProp) ||
        //        !string.Equals(typeProp.GetString(), "AdaptiveCard", StringComparison.OrdinalIgnoreCase))
        //    {
        //        throw new InvalidOperationException("Adaptive Card must have \"type\": \"AdaptiveCard\".");
        //    }

        //    if (!rootElement.TryGetProperty("version", out var versionProp) ||
        //        string.IsNullOrWhiteSpace(versionProp.GetString()))
        //    {
        //        throw new InvalidOperationException("Adaptive Card must specify a \"version\" (e.g., \"1.4\").");
        //    }

        //    // Compact JSON string from JsonDocument
        //    var json = rootElement.GetRawText();

        //    // Content length guard (Teams/Graph practical limits ~28–50 KB; keep payloads small)
        //    var contentBytes = Encoding.UTF8.GetByteCount(json);
        //    const int MaxAttachmentBytes = 48 * 1024; // 48 KB safety cap (adjust to your policy)
        //    if (contentBytes > MaxAttachmentBytes)
        //    {
        //        throw new InvalidOperationException(
        //            $"Adaptive Card payload too large ({contentBytes} bytes). Keep under {MaxAttachmentBytes} bytes.");
        //    }


        //    // Generate an attachment Id and reference it in the HTML body marker
        //    var attachmentId = Guid.NewGuid().ToString("D");

        //    var message = new ChatMessage
        //    {
        //        // IMPORTANT: Use HTML and include the attachment marker referencing the Id
        //        Body = new ItemBody
        //        {
        //            ContentType = BodyType.Html,
        //            Content = $@"<p></p><attachment id=""{attachmentId}""></attachment>"
        //            // You can put explanatory text before/after the marker if desired
        //        },
        //        Attachments = new System.Collections.Generic.List<ChatMessageAttachment>
        //        {
        //            new ChatMessageAttachment
        //            {
        //                Id = attachmentId, // REQUIRED so the marker can resolve
        //                ContentType = "application/vnd.microsoft.card.adaptive",
        //                Content = json
        //                // Optional: Name, ThumbnailUrl if you want a label/preview
        //            }
        //        }
        //    };



        //    // POST /chats/{chat-id}/messages
        //    await graph.Chats[chatId].Messages.PostAsync(message, cancellationToken: ct);
        //}




        /// <summary>
        /// Extract emails from the CSV
        /// </summary>
        private List<string> ParseUserEmailsFromCsv(string filePath)
        {
            var emails = new List<string>();

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"CSV file not found: {filePath}");

            foreach (var line in File.ReadLines(filePath))
            {
                var trimmed = line.Trim();

                if (string.IsNullOrWhiteSpace(trimmed))
                    continue;

                // Extract the first column if comma-separated
                var parts = trimmed.Split(',');

                if (parts.Length > 0)
                {
                    var email = parts[0].Trim();

                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                        emails.Add(email);
                }
            }

            return emails;
        }


        // Thread-safe UI helpers
        private void AppendStatus(string message)
        {
            if (txtStatus.InvokeRequired)
            {
                txtStatus.Invoke(new Action(() => txtStatus.AppendText(message + Environment.NewLine)));
            }
            else
            {
                txtStatus.AppendText(message + Environment.NewLine);
            }
        }

        private void SetProgress(int value)
        {
            var v = Math.Clamp(value, progressBar.Minimum, progressBar.Maximum);
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => progressBar.Value = v));
            }
            else
            {
                progressBar.Value = v;
            }
        }


    }
}
