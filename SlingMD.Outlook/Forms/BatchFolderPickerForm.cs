using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Folder picker used when slinging emails — for a batch, and for a single email when
    /// <see cref="Models.ObsidianSettings.PromptForFolderOnSling"/> is enabled. Lets the user choose an existing
    /// subfolder under the configured Inbox path, type a new one to be created, or skip and use the
    /// default Inbox path.
    /// </summary>
    public partial class BatchFolderPickerForm : Form
    {
        public enum PickerResult
        {
            UseSubfolder,
            UseDefault,
            Cancel
        }

        private readonly string _inboxPath;
        private readonly bool _inboxPathUsable;
        private ListBox _lstFolders;
        private TextBox _txtNewFolder;
        private Label _lblHeader;
        private Label _lblExisting;
        private Label _lblOr;
        private Button _btnOk;
        private Button _btnSkip;
        private Button _btnCancel;
        private Font _headerFont;

        public PickerResult Result { get; private set; } = PickerResult.Cancel;

        /// <summary>
        /// Full path to the selected or created subfolder. Empty when <see cref="Result"/> is not
        /// <see cref="PickerResult.UseSubfolder"/>.
        /// </summary>
        public string SelectedFolderPath { get; private set; } = string.Empty;

        /// <summary>
        /// True when <see cref="SelectedFolderPath"/> did not exist and this dialog created it, as
        /// opposed to the user picking a folder that was already there. Callers use this to clean up
        /// an empty folder when the sling that requested it ends up writing nothing.
        /// </summary>
        public bool CreatedNewFolder { get; private set; }

        public BatchFolderPickerForm(int emailCount, string inboxPath)
        {
            _inboxPath = inboxPath ?? string.Empty;
            _inboxPathUsable = IsUsableInboxPath(_inboxPath);
            InitializeComponents(emailCount);
            PopulateExistingFolders();

            if (!_inboxPathUsable)
            {
                // Creating a subfolder is meaningless (and potentially destructive) when we cannot
                // tell where the Inbox actually is. Leave Skip/Cancel usable so the user can still
                // sling to whatever the add-in has configured, or back out.
                _lblExisting.Text = "Vault path is not configured — set it in Settings to pick a folder.";
                _lstFolders.Enabled = false;
                _txtNewFolder.Enabled = false;
                _btnOk.Enabled = false;
            }
        }

        /// <summary>
        /// True when the configured Inbox path is absolute and well-formed, so subfolders can be
        /// created under it safely. A blank or relative path — which <see cref="Models.ObsidianSettings.GetInboxPath"/>
        /// happily returns when VaultBasePath is unset — would otherwise resolve against Outlook's
        /// working directory and silently create vault folders next to OUTLOOK.EXE. An empty path is
        /// worse still: Path.GetFullPath("") throws straight out of the OK handler.
        /// </summary>
        private static bool IsUsableInboxPath(string inboxPath)
        {
            if (string.IsNullOrWhiteSpace(inboxPath))
            {
                return false;
            }

            // Root-check the raw input, never the resolved result: Path.GetFullPath resolves a
            // relative path against the working directory, so its output is always rooted and
            // would wave through exactly the case this guard exists to catch.
            if (!Path.IsPathRooted(inboxPath))
            {
                return false;
            }

            try
            {
                Path.GetFullPath(inboxPath);
                return true;
            }
            catch (System.Exception)
            {
                // Invalid characters, path too long, unsupported format — all mean "not usable".
                return false;
            }
        }

        private void InitializeComponents(int emailCount)
        {
            this.SuspendLayout();

            _headerFont = new Font(this.Font, FontStyle.Bold);
            _lblHeader = new Label
            {
                Text = string.Format("Sling {0} {1} to:", emailCount, emailCount == 1 ? "email" : "emails"),
                AutoSize = true,
                Location = new Point(12, 12),
                Font = _headerFont
            };

            _lblExisting = new Label
            {
                Text = "Existing folders:",
                AutoSize = true,
                Location = new Point(12, _lblHeader.Bottom + 10)
            };

            _lstFolders = new ListBox
            {
                Location = new Point(12, _lblExisting.Bottom + 4),
                Size = new Size(440, 220),
                IntegralHeight = false
            };
            _lstFolders.SelectedIndexChanged += LstFolders_SelectedIndexChanged;

            _lblOr = new Label
            {
                Text = "Or create new folder:",
                AutoSize = true,
                Location = new Point(12, _lstFolders.Bottom + 10)
            };

            _txtNewFolder = new TextBox
            {
                Location = new Point(12, _lblOr.Bottom + 4),
                Width = 440
            };
            _txtNewFolder.TextChanged += TxtNewFolder_TextChanged;

            _btnOk = new Button
            {
                Text = "OK",
                Size = new Size(90, 30),
                Location = new Point(_txtNewFolder.Right - 290, _txtNewFolder.Bottom + 12)
            };
            _btnOk.Click += BtnOk_Click;

            _btnSkip = new Button
            {
                Text = "Skip (Inbox)",
                Size = new Size(100, 30),
                Location = new Point(_btnOk.Right + 10, _btnOk.Top)
            };
            _btnSkip.Click += BtnSkip_Click;

            _btnCancel = new Button
            {
                Text = "Cancel",
                Size = new Size(90, 30),
                Location = new Point(_btnSkip.Right + 10, _btnOk.Top),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[]
            {
                _lblHeader, _lblExisting, _lstFolders,
                _lblOr, _txtNewFolder,
                _btnOk, _btnSkip, _btnCancel
            });

            this.Text = emailCount == 1 ? "Sling Email" : "Sling Multiple Emails";
            this.ClientSize = new Size(464, _btnOk.Bottom + 12);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.AcceptButton = _btnOk;
            this.CancelButton = _btnCancel;

            this.ResumeLayout(false);
        }

        private void PopulateExistingFolders()
        {
            _lstFolders.Items.Clear();
            try
            {
                if (Directory.Exists(_inboxPath))
                {
                    IEnumerable<string> dirs = Directory
                        .GetDirectories(_inboxPath)
                        .Select(Path.GetFileName)
                        .OrderBy(n => n, StringComparer.OrdinalIgnoreCase);
                    foreach (string name in dirs)
                    {
                        _lstFolders.Items.Add(name);
                    }
                }
            }
            catch (System.Exception)
            {
                // Listing the inbox is best-effort; the user can still type a new folder name.
            }
        }

        private void LstFolders_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lstFolders.SelectedIndex >= 0)
            {
                _txtNewFolder.Text = string.Empty;
            }
        }

        private void TxtNewFolder_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_txtNewFolder.Text))
            {
                _lstFolders.ClearSelected();
            }
        }

        private static readonly string[] ReservedDeviceNames =
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        /// <summary>
        /// A single, safe folder-name segment: non-empty, no invalid/path chars, not '.'/'..' or
        /// all-dots/spaces, no trailing dot/space, and not a Windows reserved device name.
        /// </summary>
        private static bool IsValidFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;
            if (name != name.Trim()) return false;
            if (name.EndsWith(".")) return false;
            if (name.Trim('.', ' ').Length == 0) return false; // ".", "..", "...", "  " etc.

            string bare = name;
            int dot = bare.IndexOf('.');
            if (dot >= 0) bare = bare.Substring(0, dot);
            foreach (string reserved in ReservedDeviceNames)
            {
                if (string.Equals(bare, reserved, StringComparison.OrdinalIgnoreCase)) return false;
            }
            return true;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            // Belt and braces: the button is already disabled in this state, but an AcceptButton
            // Enter-press must never reach the path arithmetic below with an unusable inbox.
            if (!_inboxPathUsable)
            {
                return;
            }

            string typed = (_txtNewFolder.Text ?? string.Empty).Trim();
            string selected = _lstFolders.SelectedItem as string;

            string folderName = !string.IsNullOrEmpty(typed) ? typed : selected;
            if (string.IsNullOrEmpty(folderName))
            {
                MessageBox.Show(
                    "Pick an existing folder or type a name to create one.\nClick Skip (Inbox) to use the default inbox folder.",
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (!IsValidFolderName(folderName))
            {
                MessageBox.Show(
                    "Enter a single valid folder name — no invalid characters, no path separators, "
                        + "not '.' or '..', and not a reserved name.",
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string fullPath = Path.Combine(_inboxPath, folderName);

            // Defense in depth: ensure the resolved target really is directly under the inbox and did
            // not escape via a crafted name (e.g. trailing-dot/space normalization).
            string resolvedInbox;
            string resolvedTarget;
            try
            {
                resolvedInbox = Path.GetFullPath(_inboxPath);
                resolvedTarget = Path.GetFullPath(fullPath);
            }
            catch (System.Exception ex)
            {
                // A name that survives IsValidFolderName can still blow the path-length limit once
                // combined with the inbox path. Report it instead of crashing the click handler.
                MessageBox.Show(
                    string.Format("That folder name cannot be resolved to a valid path:\n{0}", ex.Message),
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string inboxPrefix = resolvedInbox.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!resolvedTarget.StartsWith(inboxPrefix, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    "That folder name would create a folder outside the Inbox. Pick a simple name.",
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            bool preExisting = Directory.Exists(fullPath);
            try
            {
                Directory.CreateDirectory(fullPath);
                CreatedNewFolder = !preExisting;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Could not create folder \"{0}\":\n{1}", fullPath, ex.Message),
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            SelectedFolderPath = fullPath;
            Result = PickerResult.UseSubfolder;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnSkip_Click(object sender, EventArgs e)
        {
            SelectedFolderPath = string.Empty;
            Result = PickerResult.UseDefault;
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// This form is built in code with no components container, so the bold header Font (a GDI
        /// handle the control does not own) must be disposed explicitly.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _headerFont?.Dispose();
                _headerFont = null;
            }
            base.Dispose(disposing);
        }
    }
}
