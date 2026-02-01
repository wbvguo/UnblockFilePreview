using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Ookii.Dialogs.WinForms;

namespace UnblockFilePreview
{
    public class MainForm : Form
    {
        private TextBox txtFolder = null!;
        private Button btnSelectFolder = null!;

        private CheckBox chkRecurse = null!;
        private CheckBox chkDryRun = null!;
        private CheckBox chkOfficeDocs = null!;

        private CheckedListBox clbExts = null!;
        private Button btnScan = null!;
        private Button btnUnblockSelected = null!;

        private ListView lvFiles = null!;
        private TextBox txtLog = null!;

        private readonly ToolTip toolTip = new ToolTip();

        // Safer default allowlist: common preview-friendly formats (NO Office by default).
        private readonly string[] defaultExtAllowlist = new[]
        {
            ".pdf", ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".log",
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"
        };

        private static readonly string[] officeExts = new[] { ".docx", ".xlsx", ".pptx" };

        // Layout constants (easy to tweak)
        private const int MarginX = 12;
        private const int TopY = 12;

        private const int LeftColWidth = 240;   // allowlist column
        private const int Gap = 18;             // space between columns
        private const int RightColWidth = 780;  // file list column (will be set by form width)
        private const int RowH = 30;

        public MainForm()
        {
                Text = "Unblock File Preview (Trusted Files)";
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "UnblockFilePreview.assets.UnblockFilePreview.ico";
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    Icon = new Icon(stream);
                }
                else
                {
                    Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                }
            }
            catch
            {
                // Fallback to executable icon if embedded resource is not found
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            Width = 1060;
            Height = 740;
            
            BuildUi();
        }

        private void BuildUi()
        {
            // ---- Row 1: folder selector (taller textbox) ----
            var lblFolder = new Label { Left = MarginX, Top = TopY + 6, Width = 60, Text = "Folder:" };

            txtFolder = new TextBox
            {
                Left = MarginX + 65,
                Top = TopY,
                Width = 770,
                Multiline = true,
                Height = 30,
                ReadOnly = true,
                ScrollBars = ScrollBars.Horizontal
            };

            btnSelectFolder = new Button
            {
                Left = MarginX + 65 + 770 + 15,
                Top = TopY,
                Width = 170,
                Height = 30,
                Text = "Select folder..."
            };
            btnSelectFolder.Click += (_, __) => SelectFolder();

            // ---- Row 2: options ----
            int optionsTop = TopY + 40;
            chkRecurse = new CheckBox { Left = MarginX + 65, Top = optionsTop, Width = 160, Text = "Include subfolders", Checked = false };
            chkDryRun = new CheckBox { Left = MarginX + 240, Top = optionsTop, Width = 220, Text = "Dry run (WhatIf)", Checked = true };

            chkOfficeDocs = new CheckBox
            {
                Left = MarginX + 470,
                Top = optionsTop,
                Width = 420,
                Text = "Include Office documents (Word / Excel / PowerPoint)",
                Checked = false
            };
            toolTip.SetToolTip(chkOfficeDocs,
                "Adds .docx/.xlsx/.pptx to the allowed extension list.\r\n" +
                "Only enable this if you trust the file source.");
            chkOfficeDocs.CheckedChanged += (_, __) => SyncOfficeExtensionsToAllowedList();

            // ---- Two-column area ----
            // Left: allowlist + buttons underneath
            int twoColTop = optionsTop + 40;

            var lblExt = new Label
            {
                Left = MarginX,
                Top = twoColTop,
                Width = LeftColWidth,
                Text = "Allowed extensions:"
            };

            int clbTop = twoColTop + 24;
            int clbHeight = 380; // taller allowlist window
            clbExts = new CheckedListBox
            {
                Left = MarginX,
                Top = clbTop,
                Width = LeftColWidth,
                Height = clbHeight,
                CheckOnClick = true
            };

            foreach (var ext in defaultExtAllowlist.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x))
                clbExts.Items.Add(ext, true);

            toolTip.SetToolTip(clbExts, "Only blocked files with these extensions will be listed/unblocked.");

            // Buttons below extension window
            int btnTop = clbTop + clbHeight + 12;
            btnScan = new Button
            {
                Left = MarginX,
                Top = btnTop,
                Width = LeftColWidth,
                Height = 34,
                Text = "Scan (blocked only)"
            };
            btnScan.Click += async (_, __) => await ScanAsync();

            btnUnblockSelected = new Button
            {
                Left = MarginX,
                Top = btnTop + 44,
                Width = LeftColWidth,
                Height = 34,
                Text = "Unblock selected",
                Enabled = false
            };
            btnUnblockSelected.Click += async (_, __) => await UnblockSelectedAsync();

            toolTip.SetToolTip(btnScan, "Find files that still have Mark of the Web (Zone.Identifier).");
            toolTip.SetToolTip(btnUnblockSelected, "Remove Mark of the Web for the checked files.");

            // Right: file list aligned to top of allowlist (top aligns with clbExts top)
            int rightLeft = MarginX + LeftColWidth + Gap;
            int rightWidth = ClientSize.Width - rightLeft - MarginX; // respect current window width
            int lvTop = clbTop; // ALIGN TOP with allowlist listbox
            int lvHeight = (btnTop + 44 + 34) - lvTop; // roughly align bottom with buttons area
            // But we also want it to feel big, so give it a minimum height and let log take the rest.
            lvHeight = Math.Max(lvHeight, 430);

            lvFiles = new ListView
            {
                Left = rightLeft,
                Top = lvTop,
                Width = rightWidth,
                Height = lvHeight,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true
            };
            lvFiles.Columns.Add("File", 500);
            lvFiles.Columns.Add("Ext", 70);
            lvFiles.Columns.Add("Size", 90);
            lvFiles.Columns.Add("Last Modified", 160);

            lvFiles.ItemChecked += (_, __) =>
            {
                btnUnblockSelected.Enabled = lvFiles.CheckedItems.Count > 0;
            };

            // Log at bottom, spanning full width
            int logTop = lvFiles.Bottom + 12;
            int logHeight = ClientSize.Height - logTop - 12;
            logHeight = Math.Max(logHeight, 90);

            txtLog = new TextBox
            {
                Left = MarginX,
                Top = logTop,
                Width = ClientSize.Width - 2 * MarginX,
                Height = logHeight,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };

            // Handle resizing so right column & log expand nicely
            Resize += (_, __) => ReflowLayout();

            Controls.Add(lblFolder);
            Controls.Add(txtFolder);
            Controls.Add(btnSelectFolder);

            Controls.Add(chkRecurse);
            Controls.Add(chkDryRun);
            Controls.Add(chkOfficeDocs);

            Controls.Add(lblExt);
            Controls.Add(clbExts);
            Controls.Add(btnScan);
            Controls.Add(btnUnblockSelected);

            Controls.Add(lvFiles);
            Controls.Add(txtLog);

            // Initial layout reflow for correct widths/heights
            ReflowLayout();
        }

        private void ReflowLayout()
        {
            // Keep right side list and log responsive on resize
            int rightLeft = MarginX + LeftColWidth + Gap;
            int rightWidth = ClientSize.Width - rightLeft - MarginX;
            if (rightWidth < 300) rightWidth = 300;

            lvFiles.Left = rightLeft;
            lvFiles.Width = rightWidth;

            // Make lvFiles height take most vertical space, keeping log visible
            int logMinHeight = 90;
            int spacing = 12;

            // lvFiles top is aligned with clbExts top; keep that
            int lvTop = lvFiles.Top;
            int logTop = ClientSize.Height - logMinHeight - spacing;
            if (logTop < lvTop + 220) logTop = lvTop + 220;

            lvFiles.Height = Math.Max(260, logTop - spacing - lvTop);

            txtLog.Left = MarginX;
            txtLog.Width = ClientSize.Width - 2 * MarginX;
            txtLog.Top = lvFiles.Bottom + spacing;
            txtLog.Height = ClientSize.Height - txtLog.Top - spacing;
            if (txtLog.Height < logMinHeight) txtLog.Height = logMinHeight;

            // Resize folder row widgets to match width changes
            txtFolder.Width = Math.Max(400, ClientSize.Width - (MarginX + 65) - 170 - 15 - MarginX);
            btnSelectFolder.Left = txtFolder.Right + 15;

            // Keep options row from overflowing too badly (leave as-is; WinForms is limited)
        }

        private void Log(string msg)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
        }

        private void SelectFolder()
        {
            var dlg = new VistaFolderBrowserDialog
            {
                Description = "Select a folder to enable Explorer preview (removes Mark of the Web for trusted files).",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = false
            };

            if (!string.IsNullOrWhiteSpace(txtFolder.Text) && Directory.Exists(txtFolder.Text))
                dlg.SelectedPath = txtFolder.Text;

            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                txtFolder.Text = dlg.SelectedPath;
                Log($"Selected folder: {dlg.SelectedPath}");
            }
        }

        private string[] GetSelectedExtensions()
        {
            return clbExts.CheckedItems
                          .Cast<string>()
                          .Select(e => e.Trim().ToLowerInvariant())
                          .Where(e => e.StartsWith("."))
                          .Distinct()
                          .ToArray();
        }

        // Toggle: add/remove Office extensions in the allowlist UI (checked by default when added).
        private void SyncOfficeExtensionsToAllowedList()
        {
            if (chkOfficeDocs.Checked)
            {
                foreach (var ext in officeExts)
                {
                    int idx = IndexOfItem(clbExts, ext);
                    if (idx < 0)
                        clbExts.Items.Add(ext, true);
                    else
                        clbExts.SetItemChecked(idx, true);
                }
                SortCheckedListBoxItems(clbExts);
            }
            else
            {
                for (int i = clbExts.Items.Count - 1; i >= 0; i--)
                {
                    var item = (clbExts.Items[i]?.ToString() ?? "").Trim();
                    if (officeExts.Contains(item, StringComparer.OrdinalIgnoreCase))
                        clbExts.Items.RemoveAt(i);
                }
            }
        }

        private static int IndexOfItem(CheckedListBox clb, string value)
        {
            for (int i = 0; i < clb.Items.Count; i++)
            {
                var item = clb.Items[i]?.ToString();
                if (string.Equals(item, value, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private static void SortCheckedListBoxItems(CheckedListBox clb)
        {
            var items = new List<(string ext, bool isChecked)>();
            for (int i = 0; i < clb.Items.Count; i++)
            {
                string ext = (clb.Items[i]?.ToString() ?? "").Trim();
                bool isChecked = clb.GetItemChecked(i);
                if (!string.IsNullOrWhiteSpace(ext))
                    items.Add((ext, isChecked));
            }

            items = items.OrderBy(x => x.ext, StringComparer.OrdinalIgnoreCase).ToList();

            clb.Items.Clear();
            foreach (var (ext, isChecked) in items)
                clb.Items.Add(ext, isChecked);
        }

        // ---------------- PowerShell runner ----------------

        private async System.Threading.Tasks.Task<(int exitCode, string stdout, string stderr)> RunPowerShellAsync(string psScript)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command " + QuoteForPowerShellCommand(psScript),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using var p = new Process { StartInfo = psi };
            p.Start();
            string stdout = await p.StandardOutput.ReadToEndAsync();
            string stderr = await p.StandardError.ReadToEndAsync();
            await p.WaitForExitAsync();
            return (p.ExitCode, stdout, stderr);
        }

        private static string QuoteForPowerShellCommand(string s)
        {
            return "\"" + s.Replace("\"", "`\"") + "\"";
        }

        private static string PsSingleQuote(string s) => s.Replace("'", "''");

        // ---------------- Scan logic ----------------

        private async System.Threading.Tasks.Task ScanAsync()
        {
            var folder = txtFolder.Text.Trim();
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("Please select a valid folder first.", "No folder selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var exts = GetSelectedExtensions();
            if (exts.Length == 0)
            {
                MessageBox.Show("Please select at least one allowed extension.", "No extensions selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnScan.Enabled = false;
            btnUnblockSelected.Enabled = false;
            btnSelectFolder.Enabled = false;

            lvFiles.Items.Clear();
            Log("Scanning for blocked files (Mark of the Web / Zone.Identifier) ...");

            string psFolder = PsSingleQuote(folder);
            string recurse = chkRecurse.Checked ? "$true" : "$false";
            string psExtArray = "@(" + string.Join(",", exts.Select(e => "'" + PsSingleQuote(e) + "'")) + ")";

            // IMPORTANT: output LastWriteTimeStr as a string to avoid locale-specific JSON datetime parsing
            string ps = $@"
$folder = '{psFolder}'
$recurse = {recurse}
$exts = {psExtArray}

$items = Get-ChildItem -LiteralPath $folder -File -Force -Recurse:$recurse |
  Where-Object {{ $exts -contains $_.Extension.ToLower() }} |
  ForEach-Object {{
    $zi = Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue
    if ($zi) {{
      [pscustomobject]@{{
        FullName = $_.FullName
        Name = $_.Name
        Ext = $_.Extension
        Length = $_.Length
        LastWriteTimeStr = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
      }}
    }}
  }}

$items | ConvertTo-Json -Compress
";

            var (code, stdout, stderr) = await RunPowerShellAsync(ps);

            if (!string.IsNullOrWhiteSpace(stderr))
                Log("PS ERR: " + stderr.Trim());

            if (code != 0)
            {
                Log($"Scan failed. Exit code: {code}");
                btnScan.Enabled = true;
                btnSelectFolder.Enabled = true;
                return;
            }

            var json = stdout.Trim();
            if (string.IsNullOrWhiteSpace(json) || json == "null")
            {
                Log("No blocked files found (within allowlist).");
                btnScan.Enabled = true;
                btnSelectFolder.Enabled = true;
                return;
            }

            List<FileRow> rows;
            try
            {
                rows = ParseJsonRows(json);
            }
            catch (Exception ex)
            {
                Log("Failed to parse scan results: " + ex.Message);
                btnScan.Enabled = true;
                btnSelectFolder.Enabled = true;
                return;
            }

            foreach (var r in rows.OrderBy(r => r.FullName, StringComparer.OrdinalIgnoreCase))
            {
                var item = new ListViewItem(r.Name)
                {
                    Tag = r.FullName,
                    Checked = true
                };
                item.SubItems.Add(r.Ext);
                item.SubItems.Add(FormatBytes(r.Length));
                item.SubItems.Add(r.LastWriteTimeStr);
                lvFiles.Items.Add(item);
            }

            Log($"Scan complete. Blocked files found: {rows.Count}");

            btnScan.Enabled = true;
            btnSelectFolder.Enabled = true;
            btnUnblockSelected.Enabled = lvFiles.CheckedItems.Count > 0;
        }

        // ---------------- Unblock logic ----------------

        private async System.Threading.Tasks.Task UnblockSelectedAsync()
        {
            if (lvFiles.CheckedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one file to unblock.", "No selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!chkDryRun.Checked)
            {
                var res = MessageBox.Show(
                    "This will remove 'downloaded from the Internet' mark (Mark of the Web) for the selected files.\n\nProceed only if you trust the source.",
                    "Confirm unblock",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (res != DialogResult.Yes) return;
            }

            btnUnblockSelected.Enabled = false;
            btnScan.Enabled = false;
            btnSelectFolder.Enabled = false;

            var paths = lvFiles.CheckedItems.Cast<ListViewItem>()
                           .Select(i => (string)i.Tag!)
                           .ToArray();

            Log($"{(chkDryRun.Checked ? "Dry run" : "Unblocking")} {paths.Length} file(s) ...");

            string whatIf = chkDryRun.Checked ? " -WhatIf" : "";
            string psPaths = "@(" + string.Join(",", paths.Select(p => "'" + PsSingleQuote(p) + "'")) + ")";

            string ps = $@"
$paths = {psPaths}
foreach ($p in $paths) {{
  try {{
    Unblock-File -LiteralPath $p -ErrorAction Continue{whatIf}
    Write-Output $p
  }} catch {{
    Write-Error ""Failed: $p -> $($_.Exception.Message)""
  }}
}}
";

            var (code, _, stderr) = await RunPowerShellAsync(ps);

            if (!string.IsNullOrWhiteSpace(stderr))
                Log("PS ERR: " + stderr.Trim());

            if (code != 0)
            {
                Log($"Unblock step failed. Exit code: {code}");
            }
            else
            {
                if (chkDryRun.Checked)
                {
                    Log("Dry run complete (no changes made).");
                }
                else
                {
                    Log("Unblock complete. Refreshing list ...");
                    await ScanAsync();
                }
            }

            btnScan.Enabled = true;
            btnSelectFolder.Enabled = true;
            btnUnblockSelected.Enabled = lvFiles.CheckedItems.Count > 0;
        }

        // ---------------- Helpers ----------------

        private static string FormatBytes(long bytes)
        {
            string[] suffix = { "B", "KB", "MB", "GB", "TB" };
            double v = bytes;
            int i = 0;
            while (v >= 1024 && i < suffix.Length - 1) { v /= 1024; i++; }
            return $"{v:0.##} {suffix[i]}";
        }

        private sealed class FileRow
        {
            public string FullName { get; set; } = "";
            public string Name { get; set; } = "";
            public string Ext { get; set; } = "";
            public long Length { get; set; }
            public string LastWriteTimeStr { get; set; } = "";
        }

        private static List<FileRow> ParseJsonRows(string json)
        {
            var rows = new List<FileRow>();
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var el in doc.RootElement.EnumerateArray())
                    rows.Add(ParseRow(el));
            }
            else if (doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                rows.Add(ParseRow(doc.RootElement));
            }
            else
            {
                throw new InvalidOperationException("Unexpected JSON output from PowerShell scan.");
            }

            return rows;

            static FileRow ParseRow(JsonElement el)
            {
                return new FileRow
                {
                    FullName = el.GetProperty("FullName").GetString() ?? "",
                    Name = el.GetProperty("Name").GetString() ?? "",
                    Ext = el.GetProperty("Ext").GetString() ?? "",
                    Length = el.TryGetProperty("Length", out var len) ? len.GetInt64() : 0,
                    LastWriteTimeStr = el.TryGetProperty("LastWriteTimeStr", out var t)
                        ? (t.GetString() ?? "")
                        : ""
                };
            }
        }
    }
}
