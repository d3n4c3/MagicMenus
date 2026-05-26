using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using MagicMenus.Properties;
using System.Configuration;
using MagicMenus;
using MagicMenus.Settings;

namespace QuickMenu
{

    public partial class MagicMenus : Form
    {
        KeyboardHook hook = new KeyboardHook();
        KeyboardHook hook2 = new KeyboardHook();

        public bool clipboardSettings = false;
        public bool actionSettings = false;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
            (

            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
            );

        public object lb_item = null;
        List<string> ActionList = new List<string>();

        public MagicMenus()
        {
            InitializeComponent();

            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 5, 5));

            // Diagnostic: when launched with MM_SCREENSHOT=1 in the
            // environment we force-open the Settings dialog in dark mode at
            // the Action Menu tab. This is used by capture-screenshot.ps1 to
            // headlessly verify theme rendering during development; it is a
            // no-op in normal operation.
            bool screenshotMode = Environment.GetEnvironmentVariable("MM_SCREENSHOT") == "1";
            string screenshotTab = Environment.GetEnvironmentVariable("MM_SCREENSHOT_TAB");
            // When capturing a modal sub-dialog (Tutorial / About) the
            // main Settings window should stay hidden so the captured
            // image only contains the modal. We compute this here and
            // also use it later to suppress the "no hotkey set => auto
            // open Settings" code path further down in this constructor.
            bool capturingModalDialog =
                screenshotMode && screenshotTab != null &&
                (screenshotTab.StartsWith("tutorial", StringComparison.OrdinalIgnoreCase) ||
                 screenshotTab.StartsWith("about", StringComparison.OrdinalIgnoreCase));
            if (screenshotMode)
            {
                // Default to dark mode for screenshots unless the harness
                // explicitly asked for a light-mode capture via the special
                // "-light" suffix on a tutorial step.
                ThemeManager.IsDark =
                    !(screenshotTab != null &&
                      screenshotTab.IndexOf("light", StringComparison.OrdinalIgnoreCase) >= 0);
                // actionSettings makes QuickMenu_Load resize the form to
                // 446x365 to show the Settings dialog. Skip this when the
                // harness is capturing a modal sub-dialog - the modal
                // draws its own window and the settings dialog behind it
                // would just clutter the screenshot.
                if (!capturingModalDialog) actionSettings = true;
            }

            // Apply the persisted theme before the form is first painted so
            // there is no flash of the default light palette in dark mode.
            chkDM.Checked = ThemeManager.IsDark;
            ApplyCurrentTheme();

            if (screenshotMode)
            {
                this.Shown += delegate
                {
                    if (!string.IsNullOrEmpty(screenshotTab))
                    {
                        try
                        {
                            if (screenshotTab.Equals("general", StringComparison.OrdinalIgnoreCase))
                                tabControl2.SelectedTab = tabGeneral;
                            else if (screenshotTab.Equals("general-settings", StringComparison.OrdinalIgnoreCase))
                            {
                                tabControl2.SelectedTab = tabGeneral;
                                if (lbWildcards.Items.Count > 1) lbWildcards.SelectedIndex = 1;
                            }
                            else if (screenshotTab.Equals("general-custom", StringComparison.OrdinalIgnoreCase))
                            {
                                tabControl2.SelectedTab = tabGeneral;
                                txtWCName.Text = "EMAIL";
                                txtWCString.Text = "katco@example.com";
                                btnWCString_Click(this, EventArgs.Empty);
                            }
                            else if (screenshotTab.Equals("cleanup", StringComparison.OrdinalIgnoreCase))
                            {
                                Settings.Default.settingsWildcards = new System.Collections.Specialized.StringCollection();
                                Settings.Default.Save();
                            }
                            else if (screenshotTab.StartsWith("about", StringComparison.OrdinalIgnoreCase))
                            {
                                BeginInvoke(new Action(delegate
                                {
                                    AboutForm dlg = new AboutForm();
                                    dlg.StartPosition = FormStartPosition.CenterScreen;
                                    dlg.Show();
                                }));
                            }
                            else if (screenshotTab.StartsWith("tutorial", StringComparison.OrdinalIgnoreCase))
                            {
                                // tutorial / tutorial-1 / tutorial-2 ... open the
                                // wizard, optionally pre-advanced to a specific
                                // step so the screenshot harness can capture each
                                // page.
                                int targetStep = 0;
                                int dash = screenshotTab.LastIndexOf('-');
                                if (dash > 0 && dash < screenshotTab.Length - 1)
                                    int.TryParse(screenshotTab.Substring(dash + 1), out targetStep);
                                // Show the tutorial non-modally so this lambda
                                // doesn't block.
                                BeginInvoke(new Action(delegate
                                {
                                    TutorialForm dlg = new TutorialForm(this);
                                    dlg.StartPosition = FormStartPosition.CenterScreen;
                                    dlg.Show();
                                    for (int i = 0; i < targetStep; i++)
                                        dlg.AdvanceForScreenshot();
                                }));
                            }
                            else if (screenshotTab.Equals("action", StringComparison.OrdinalIgnoreCase))
                                tabControl2.SelectedTab = tabAction;
                            else if (screenshotTab.Equals("action-link", StringComparison.OrdinalIgnoreCase))
                            {
                                tabControl2.SelectedTab = tabAction;
                                if (tabControl1.TabCount > 1) tabControl1.SelectedIndex = 1;
                            }
                            else if (screenshotTab.Equals("clipboard", StringComparison.OrdinalIgnoreCase))
                                tabControl2.SelectedTab = tabClipboard;
                            else if (screenshotTab.Equals("clipboard-string", StringComparison.OrdinalIgnoreCase))
                            {
                                tabControl2.SelectedTab = tabClipboard;
                                if (tabControl3.TabCount > 1) tabControl3.SelectedIndex = 1;
                            }
                        }
                        catch { }
                    }
                    this.TopMost = true;
                    this.BringToFront();
                    this.Activate();
                    this.Refresh();
                };
            }

            //Action Hook
            hook.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);

            if (Settings.Default.settingsActionHotkey == "")
            {
                // Don't yank Settings open if we're just capturing a
                // screenshot of a modal sub-dialog (Tutorial / About).
                if (!capturingModalDialog) actionSettings = true;
            }
            else
            {
                textBox4.Text = Settings.Default.settingsActionHotkey;
                string[] Splitter = textBox4.Text.Split(char.Parse(","));
                string Modkey1 = Splitter[0].Trim();
                string Modkey2 = Splitter[1].Trim();
                string Keypress1 = Splitter[2].Trim();
                ModKeys modKeys1 = ModKeys.Shift;
                ModKeys modKeys2 = ModKeys.Control;
                if (Modkey1.Contains("Shift")) modKeys1 = ModKeys.Shift;
                if (Modkey1.Contains("Alt")) modKeys1 = ModKeys.Alt;
                if (Modkey1.Contains("Control")) modKeys1 = ModKeys.Control;
                if (Modkey2.Contains("Shift")) modKeys2 = ModKeys.Shift;
                if (Modkey2.Contains("Alt")) modKeys2 = ModKeys.Alt;
                if (Modkey2.Contains("Control")) modKeys2 = ModKeys.Control;
                Keys K1 = (Keys)Enum.Parse(typeof(Keys), Keypress1);
                hook.RegisterHotKey(modKeys1 | modKeys2, K1);
                textBox4.Text = textBox4.Text.Replace(",", "+");
                textBox4.Text = textBox4.Text.ToUpper();

            }
            //End Action Hook

            //Clipboard Hook
            hook2.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(hook2_KeyPressed);

            if (Settings.Default.settingsClipboardHotkey == "")
            {
                //MessageBox.Show("You need to set a hotkey for the clipboard menu.", "Set Clipboard Hotkey", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (!capturingModalDialog) clipboardSettings = true;
            }
            else
            {
                textBox1.Text = Settings.Default.settingsClipboardHotkey;
                string[] Splitter = textBox1.Text.Split(char.Parse(","));
                string Modkey1 = Splitter[0].Trim();
                string Modkey2 = Splitter[1].Trim();
                string Keypress1 = Splitter[2].Trim();
                ModKeys modKeys1 = ModKeys.Shift;
                ModKeys modKeys2 = ModKeys.Control;
                if (Modkey1.Contains("Shift")) modKeys1 = ModKeys.Shift;
                if (Modkey1.Contains("Alt")) modKeys1 = ModKeys.Alt;
                if (Modkey1.Contains("Control")) modKeys1 = ModKeys.Control;
                if (Modkey2.Contains("Shift")) modKeys2 = ModKeys.Shift;
                if (Modkey2.Contains("Alt")) modKeys2 = ModKeys.Alt;
                if (Modkey2.Contains("Control")) modKeys2 = ModKeys.Control;
                Keys K1 = (Keys)Enum.Parse(typeof(Keys), Keypress1);
                hook2.RegisterHotKey(modKeys1 | modKeys2, K1);
                textBox1.Text = textBox1.Text.Replace(",", "+");
                textBox1.Text = textBox1.Text.ToUpper();

            }
            //End Clipboard Hook

            //String Wildcards
            // The "Add String Wildcard" GroupBox is disabled in the
            // designer - enable it now that we wire up Add/Delete and
            // populate the listbox with any persisted user wildcards.
            groupBox6.Enabled = true;
            btnWCString.Click += new EventHandler(btnWCString_Click);
            txtWCName.KeyDown += new KeyEventHandler(WildcardInput_KeyDown);
            txtWCString.KeyDown += new KeyEventHandler(WildcardInput_KeyDown);
            lbWildcards.KeyDown += new KeyEventHandler(lbWildcards_KeyDown);

            // Context menu for deleting custom wildcards. Built-in wildcards
            // (%INPUT%, %SETTINGS%) are protected in the click handler.
            ContextMenuStrip wildcardMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteItem = new ToolStripMenuItem("Delete wildcard");
            deleteItem.Click += new EventHandler(WildcardMenuDelete_Click);
            wildcardMenu.Items.Add(deleteItem);
            wildcardMenu.Opening += new CancelEventHandler(WildcardMenu_Opening);
            ThemeManager.ApplyMenu(wildcardMenu);
            lbWildcards.ContextMenuStrip = wildcardMenu;

            LoadCustomWildcards();
            //End String Wildcards

            // Tutorial entry point on the General tab. Adding the button in
            // code (vs. the designer) keeps the change localized to the
            // tutorial feature and avoids regenerating the form layout.
            InstallTutorialButton();

            // "About" entry on the tray-icon right-click menu, inserted
            // between Settings and Exit so users have a one-click path to
            // version / author info.
            InstallAboutMenuItem();

            // Auto-open the onboarding tutorial whenever the user hasn't
            // already completed (or explicitly skipped) it. Existing
            // installs upgrading to this version will see it once; if they
            // dismiss it, we set tutorialCompleted so it never auto-opens
            // again. The "Show tutorial" button on the General tab is the
            // explicit re-run path. The screenshot harness has its own
            // scripted flow, so opt out.
            if (!screenshotMode && !Settings.Default.tutorialCompleted)
            {
                // Defer until the main form is shown so the modal centers on
                // it correctly.
                this.Shown += FirstRunShowTutorial;
            }
        }

        /// <summary>
        /// Add a "Show tutorial" button to the Personalization group on the
        /// General tab. The button mirrors the styling of the dark-mode
        /// checkbox row and re-opens the onboarding wizard on demand.
        /// </summary>
        private void InstallTutorialButton()
        {
            if (groupBox2 == null) return;
            Button btn = new Button();
            btn.Name = "btnShowTutorial";
            btn.Text = "Show tutorial";
            btn.Size = new Size(110, 24);
            // Right-align inside the Personalization group, vertically
            // centered on the dark mode checkbox row.
            btn.Location = new Point(groupBox2.ClientSize.Width - btn.Width - 12, 16);
            btn.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btn.Cursor = Cursors.Hand;
            btn.Tag = ThemeManager.TagPrimary;
            btn.Click += new EventHandler(ShowTutorialButton_Click);
            groupBox2.Controls.Add(btn);
            ThemeManager.RestyleControl(btn);
        }

        private void ShowTutorialButton_Click(object sender, EventArgs e)
        {
            ShowTutorial();
        }

        /// <summary>
        /// Insert an "About" item into the tray icon's right-click menu,
        /// placed just before "Exit" so it sits in the natural reading
        /// order (Settings, About, Exit). Wired to open the AboutForm
        /// modally.
        /// </summary>
        private void InstallAboutMenuItem()
        {
            if (icoStrip == null) return;
            ToolStripMenuItem about = new ToolStripMenuItem();
            about.Name = "aboutToolStripMenuItem";
            about.Text = "&About";
            about.Click += AboutMenuItem_Click;
            // Insert above Exit; fall back to Add if the Exit item isn't
            // where we expect (defensive in case the designer changes).
            int exitIndex = icoStrip.Items.IndexOf(exitToolStripMenuItem);
            if (exitIndex >= 0) icoStrip.Items.Insert(exitIndex, about);
            else icoStrip.Items.Add(about);
            ThemeManager.ApplyMenu(icoStrip);
        }

        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            using (AboutForm dlg = new AboutForm())
            {
                dlg.ShowDialog(this);
            }
        }

        private void FirstRunShowTutorial(object sender, EventArgs e)
        {
            // Detach so this only runs once per launch even if the form is
            // re-shown later.
            this.Shown -= FirstRunShowTutorial;
            // Make sure Settings is visible behind the modal so the user has
            // context for what the tutorial is talking about.
            this.Width = 446;
            this.Height = 365;
            this.CenterToScreen();
            this.TopMost = false;
            tabControl2.SelectedTab = tabGeneral;
            ShowTutorial();
        }

        private void ShowTutorial()
        {
            using (TutorialForm dlg = new TutorialForm(this))
            {
                dlg.ShowDialog(this);
            }
        }

        /// <summary>
        /// Re-load the action menu listbox from settings and regenerate the
        /// tray menu. Called from the tutorial after it adds menu items so
        /// the change is visible immediately, without restarting.
        /// </summary>
        public void RefreshActionMenuUI()
        {
            lbActionMenu.Items.Clear();
            var list = Settings.Default.settingsActionList.Cast<string>().ToList();
            listLoader(lbActionMenu, list);
            GenerateActionMenu();
        }

        /// <summary>
        /// Same as RefreshActionMenuUI but for the clipboard side.
        /// </summary>
        public void RefreshClipboardMenuUI()
        {
            lbClipboardMenu.Items.Clear();
            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();
            listLoader(lbClipboardMenu, list);
            GenerateClipboardMenu();
        }

        /// <summary>
        /// Re-apply the active palette to this form. Public wrapper around
        /// the private ApplyCurrentTheme so the tutorial can flip between
        /// light and dark mode and have the change reflected on the
        /// settings dialog behind it as well as the system tray menus.
        /// </summary>
        public void RefreshTheme()
        {
            chkDM.Checked = ThemeManager.IsDark;
            ApplyCurrentTheme();
        }

        /// <summary>
        /// Pop the SetHotKey dialog, capture the chosen combo, and register
        /// it as the action menu hotkey without restarting the app. Returns
        /// true if a hotkey was set, false if the user cancelled. Used by
        /// the tutorial's final step.
        /// </summary>
        public bool SetActionHotkeyInteractively()
        {
            return SetHotkeyInteractivelyCore(true);
        }

        /// <summary>
        /// Same as SetActionHotkeyInteractively but for the clipboard
        /// menu hotkey.
        /// </summary>
        public bool SetClipboardHotkeyInteractively()
        {
            return SetHotkeyInteractivelyCore(false);
        }

        private bool SetHotkeyInteractivelyCore(bool action)
        {
            SetHotKey dlg = new SetHotKey();
            dlg.ShowDialog(this);
            if (string.IsNullOrEmpty(dlg.hkModkeys)) return false;

            string[] parts = dlg.hkModkeys.Split(',');
            ModKeys mod1 = ParseModKey(parts.Length > 0 ? parts[0] : "");
            ModKeys mod2 = ParseModKey(parts.Length > 1 ? parts[1] : "");

            if (action)
            {
                try { hook.Dispose(); } catch { }
                hook = new KeyboardHook();
                hook.KeyPressed += new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);
                hook.RegisterHotKey(mod1 | mod2, dlg.hkKeypress);
                Settings.Default.settingsActionHotkey = dlg.hkModkeys + ", " + dlg.hkKeypress;
                textBox4.Text = Settings.Default.settingsActionHotkey
                    .Replace(",", "+").ToUpperInvariant();
            }
            else
            {
                try { hook2.Dispose(); } catch { }
                hook2 = new KeyboardHook();
                hook2.KeyPressed += new EventHandler<KeyPressedEventArgs>(hook2_KeyPressed);
                hook2.RegisterHotKey(mod1 | mod2, dlg.hkKeypress);
                Settings.Default.settingsClipboardHotkey = dlg.hkModkeys + ", " + dlg.hkKeypress;
                textBox1.Text = Settings.Default.settingsClipboardHotkey
                    .Replace(",", "+").ToUpperInvariant();
            }
            Settings.Default.Save();
            return true;
        }

        private static ModKeys ParseModKey(string s)
        {
            if (string.IsNullOrEmpty(s)) return ModKeys.Shift;
            if (s.Contains("Alt")) return ModKeys.Alt;
            if (s.Contains("Control")) return ModKeys.Control;
            return ModKeys.Shift;
        }

        // Tokens reserved by the application - users cannot create wildcards
        // with these names because Process_Wildcard handles them specially.
        private static readonly string[] _reservedWildcards = new[] { "%INPUT%", "%SETTINGS%" };

        // The first two listbox rows are always built-in wildcards. Custom
        // wildcards live at index 2 and beyond.
        private const int _builtInWildcardCount = 2;

        /// <summary>
        /// Read persisted custom wildcards out of settings and re-populate
        /// the listbox below the built-in entries. Safe to call repeatedly.
        /// </summary>
        private void LoadCustomWildcards()
        {
            // Clear out any custom rows from a previous load while leaving
            // the two built-in entries (%INPUT%, %SETTINGS%) in place.
            while (lbWildcards.Items.Count > _builtInWildcardCount)
                lbWildcards.Items.RemoveAt(_builtInWildcardCount);

            System.Collections.Specialized.StringCollection saved = Settings.Default.settingsWildcards;
            if (saved == null) return;
            foreach (string entry in saved)
            {
                if (string.IsNullOrEmpty(entry)) continue;
                int pipe = entry.IndexOf('|');
                if (pipe < 1) continue;
                string token = entry.Substring(0, pipe);
                lbWildcards.Items.Add(token);
            }
        }

        /// <summary>
        /// Return the saved value for a given wildcard token (e.g.
        /// "%FOO%"), or null if no such entry exists.
        /// </summary>
        private string GetCustomWildcardValue(string token)
        {
            System.Collections.Specialized.StringCollection saved = Settings.Default.settingsWildcards;
            if (saved == null) return null;
            foreach (string entry in saved)
            {
                if (string.IsNullOrEmpty(entry)) continue;
                int pipe = entry.IndexOf('|');
                if (pipe < 1) continue;
                if (string.Equals(entry.Substring(0, pipe), token, StringComparison.OrdinalIgnoreCase))
                    return entry.Substring(pipe + 1);
            }
            return null;
        }

        /// <summary>
        /// Persist a new or updated wildcard. Returns true on success and
        /// false (with an error message already displayed) if the input is
        /// invalid.
        /// </summary>
        private bool SaveCustomWildcard(string rawName, string value)
        {
            // Sanitize: strip surrounding %, whitespace, and any character
            // that isn't a letter or digit so the resulting token is
            // unambiguous to Process_Wildcard's String.Replace.
            if (rawName == null) rawName = string.Empty;
            string name = rawName.Trim().Trim('%');
            System.Text.StringBuilder cleaned = new System.Text.StringBuilder(name.Length);
            foreach (char ch in name)
                if (char.IsLetterOrDigit(ch) || ch == '_') cleaned.Append(char.ToUpperInvariant(ch));
            name = cleaned.ToString();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(this,
                    "Please enter a wildcard name. Use letters, digits, or underscores - no spaces.",
                    "Invalid wildcard name",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string token = "%" + name + "%";
            foreach (string reserved in _reservedWildcards)
            {
                if (string.Equals(reserved, token, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this,
                        token + " is reserved by Magic Menus and can't be used as a custom wildcard.",
                        "Reserved wildcard",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            if (value == null) value = string.Empty;
            if (value.Contains("|"))
            {
                MessageBox.Show(this,
                    "Wildcard values can't contain the '|' character.",
                    "Invalid wildcard value",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            System.Collections.Specialized.StringCollection saved = Settings.Default.settingsWildcards
                ?? new System.Collections.Specialized.StringCollection();

            // Remove an existing entry with the same name (case-insensitive)
            // so a second Add for the same token acts as an update.
            for (int i = saved.Count - 1; i >= 0; i--)
            {
                string entry = saved[i];
                if (string.IsNullOrEmpty(entry)) continue;
                int pipe = entry.IndexOf('|');
                if (pipe < 1) continue;
                if (string.Equals(entry.Substring(0, pipe), token, StringComparison.OrdinalIgnoreCase))
                    saved.RemoveAt(i);
            }

            saved.Add(token + "|" + value);
            Settings.Default.settingsWildcards = saved;
            Settings.Default.Save();

            LoadCustomWildcards();

            // Select the just-added entry so the user sees the description
            // panel update with their new wildcard.
            int newIndex = lbWildcards.Items.IndexOf(token);
            if (newIndex >= 0) lbWildcards.SelectedIndex = newIndex;

            return true;
        }

        /// <summary>
        /// Remove a custom wildcard from settings. Built-in wildcards
        /// cannot be deleted.
        /// </summary>
        private void DeleteCustomWildcard(string token)
        {
            if (string.IsNullOrEmpty(token)) return;
            foreach (string reserved in _reservedWildcards)
                if (string.Equals(reserved, token, StringComparison.OrdinalIgnoreCase)) return;

            System.Collections.Specialized.StringCollection saved = Settings.Default.settingsWildcards;
            if (saved == null) return;

            bool changed = false;
            for (int i = saved.Count - 1; i >= 0; i--)
            {
                string entry = saved[i];
                if (string.IsNullOrEmpty(entry)) continue;
                int pipe = entry.IndexOf('|');
                if (pipe < 1) continue;
                if (string.Equals(entry.Substring(0, pipe), token, StringComparison.OrdinalIgnoreCase))
                {
                    saved.RemoveAt(i);
                    changed = true;
                }
            }
            if (!changed) return;

            Settings.Default.settingsWildcards = saved;
            Settings.Default.Save();
            LoadCustomWildcards();
            lbWildcards.SelectedIndex = 0;
        }

        private void btnWCString_Click(object sender, EventArgs e)
        {
            if (SaveCustomWildcard(txtWCName.Text, txtWCString.Text))
            {
                txtWCName.Clear();
                txtWCString.Clear();
                txtWCName.Focus();
            }
        }

        private void WildcardInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnWCString_Click(sender, EventArgs.Empty);
            }
        }

        private void lbWildcards_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && lbWildcards.SelectedIndex >= _builtInWildcardCount)
            {
                e.SuppressKeyPress = true;
                ConfirmAndDeleteSelectedWildcard();
            }
        }

        private void WildcardMenu_Opening(object sender, CancelEventArgs e)
        {
            // Hide the "Delete" item for the two built-in wildcards.
            e.Cancel = lbWildcards.SelectedIndex < _builtInWildcardCount;
        }

        private void WildcardMenuDelete_Click(object sender, EventArgs e)
        {
            ConfirmAndDeleteSelectedWildcard();
        }

        private void ConfirmAndDeleteSelectedWildcard()
        {
            if (lbWildcards.SelectedIndex < _builtInWildcardCount) return;
            string token = lbWildcards.SelectedItem as string;
            if (string.IsNullOrEmpty(token)) return;
            DialogResult dr = MessageBox.Show(this,
                "Delete the wildcard " + token + "?\n\nMenu items that reference it will stop expanding the token.",
                "Delete wildcard",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes) DeleteCustomWildcard(token);
        }

        private void DragAndDrop(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void QuickMenu_Load(object sender, EventArgs e)
        {
            
            lbActionMenu.Items.Clear();
            var list = Settings.Default.settingsActionList.Cast<string>().ToList();
            listLoader(lbActionMenu, list);

            lbClipboardMenu.Items.Clear();
            var list2 = Settings.Default.settingClipboardList.Cast<string>().ToList();
            listLoader(lbClipboardMenu, list2);

            btnUp.Text = char.ConvertFromUtf32(0x2191);
            btnDown.Text = char.ConvertFromUtf32(0x2193);

            btnCUp.Text = char.ConvertFromUtf32(0x2191);
            btnCDown.Text = char.ConvertFromUtf32(0x2193);
            gbGeneral.BringToFront();
            GenerateActionMenu();
            GenerateClipboardMenu();

            // Hide leftover developer/todo controls that aren't part of the
            // shipped UI and would otherwise show when the form is widened.
            if (listBox1 != null) listBox1.Visible = false;
            if (label28 != null) label28.Visible = false;

            ApplyCurrentTheme();

            Point p = new Point();
            this.Height = 0;
            this.Width = 0;
            p.X = Screen.PrimaryScreen.WorkingArea.Left;
            p.Y = Screen.PrimaryScreen.WorkingArea.Top;
            this.Location = p;
            this.TopMost = true;
            SetForegroundWindow(this.Handle);

            if (clipboardSettings)
            {
                this.Height = 365;
                this.Width = 446;
                this.CenterToScreen();
                tabControl2.SelectedTab = tabClipboard;
            }
            if (actionSettings)
            {
                this.Height = 365;
                this.Width = 446;
                this.CenterToScreen();
                tabControl2.SelectedTab = tabAction;
            }
        }

        private void QuickMenu_DragDrop(object sender, DragEventArgs e)
        {
            lb_item = null;
        }

        void hook_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - qMenuStrip.Height;
            qMenuStrip.Show(0, y);
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }
        void hook2_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - cMenuStrip.Height;
            cMenuStrip.Show(0, y);
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }

        private void QuickMenu_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        public void qMenuStrip_Opening(object sender, CancelEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            qMenuStrip.AutoClose = true;
            qMenuStrip.Close();
            label1.BackColor = ThemeManager.CloseHoverColor;
            Point p = new Point();
            this.Height = 0;
            this.Width = 0;
            p.X = Screen.PrimaryScreen.WorkingArea.Left;
            p.Y = Screen.PrimaryScreen.WorkingArea.Top;
            this.Location = p;
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
            label1.BackColor = ThemeManager.TitleBarBackground;
        }

        private void label1_MouseEnter(object sender, EventArgs e)
        {
            label1.BackColor = ThemeManager.CloseHoverColor;
            label1.ForeColor = Color.White;
        }
        private void label1_MouseLeave(object sender, EventArgs e)
        {
            label1.BackColor = ThemeManager.TitleBarBackground;
            label1.ForeColor = ThemeManager.TitleBarForeground;
        }

        private void chkDM_CheckedChanged(object sender, EventArgs e)
        {
            ThemeManager.IsDark = chkDM.Checked;
            ApplyCurrentTheme();
        }

        private void ApplyCurrentTheme()
        {
            ThemeManager.Apply(this);
            ThemeManager.ApplyTitleBar(pictureBox1, label3, label1);
            ThemeManager.ApplyMenu(qMenuStrip);
            ThemeManager.ApplyMenu(cMenuStrip);
            ThemeManager.ApplyMenu(icoStrip);
            this.Invalidate(true);
        }

        private void label2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void label3_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void createMenuItem(string itemName, Keys itemHotkey, string itemText, string itemType, string itemPath, string itemArg)
        {
            string toolTipData = itemType + "ç" + itemPath + "ç" + itemArg;

            if (itemType == "Seperator")
            {
                ToolStripSeparator qSeperator = new ToolStripSeparator();
                qMenuStrip.Items.Add(qSeperator);
                return;
            }

            ToolStripMenuItem qMenuItem = new ToolStripMenuItem(itemName)
            {

                Name = itemName,
                Text = "[&" + itemHotkey + "] " + itemText,
                TextAlign = ContentAlignment.BottomRight,
                ToolTipText = toolTipData,
            };

            qMenuStrip.Items.Add(qMenuItem);

            qMenuItem.Click += new EventHandler(this.qMenuItems_Click);
        
        }

        private void createCMenuItem(string itemName, Keys itemHotkey, string itemText, string itemType, string itemPath)
        {
            string toolTipData = itemType + "ç" + itemPath + "ç";

            if (itemType == "Seperator")
            {
                ToolStripSeparator cSeperator = new ToolStripSeparator();
                cMenuStrip.Items.Add(cSeperator);
                return;
            }

            ToolStripMenuItem cMenuItem = new ToolStripMenuItem(itemName)
            {

                Name = itemName,
                Text = "[&" + itemHotkey + "] " + itemText,
                TextAlign = ContentAlignment.BottomRight,
                ToolTipText = toolTipData,
            };

            cMenuStrip.Items.Add(cMenuItem);

            cMenuItem.Click += new EventHandler(this.cMenuItems_Click);

        }

        private void qMenuItems_Click(object sender, EventArgs e)
        {
            try
            {
            ToolStripMenuItem SenderStrip = (ToolStripMenuItem)sender;
            string tData = SenderStrip.ToolTipText;
            string[] Splitter = tData.Split(char.Parse("ç"));
            string aType = Splitter[0];
            string aPath = Splitter[1];
            string aArg = Splitter[2];
           
            if (aType == "Link") {
               aPath = Process_Wildcard(aPath);
                if (!string.IsNullOrEmpty(aPath)) Process.Start(aPath);
            }

            if (aType == "Shell")
            {
                aArg = Process_Wildcard(aArg);
                aPath = Process_Wildcard(aPath);
                Process p = new Process();
                p.StartInfo.FileName = aPath;
                p.StartInfo.Arguments = aArg;
                if (!string.IsNullOrEmpty(aPath)) p.Start();
            }

            if (aType == "Settings")
            {
                this.Height = 365;
                this.Width = 446;

                this.CenterToScreen();
            }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void cMenuItems_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem SenderStrip = (ToolStripMenuItem)sender;
            string tData = SenderStrip.ToolTipText;
            string[] Splitter = tData.Split(char.Parse("ç"));
            string aType = Splitter[0];
            string aPathorString = Splitter[1];

            if (aType == "RTF")
            {
                using (var rtf = new RichTextBox())
                {
          
                    rtf.Rtf = File.ReadAllText(aPathorString);
         
                    Clipboard.Clear();

                    DataObject data = new DataObject();
                    data.SetData(DataFormats.Text, rtf.Text);
                    data.SetData(DataFormats.Rtf, rtf.Rtf);

                    Clipboard.SetDataObject(data);
                }

            }

            if (aType == "String")
            {
                Clipboard.Clear();
                DataObject data = new DataObject();
                data.SetData(DataFormats.Text, aPathorString);
                data.SetData(DataFormats.Rtf, aPathorString);

                Clipboard.SetDataObject(data);
            }

        }
        private string Process_Wildcard(string text)
        {
            if (text == null) return text;

            // User-defined string wildcards are expanded first so that a
            // custom wildcard can itself reference %INPUT% (which is then
            // resolved on the second pass below).
            System.Collections.Specialized.StringCollection saved = Settings.Default.settingsWildcards;
            if (saved != null && saved.Count > 0)
            {
                foreach (string entry in saved)
                {
                    if (string.IsNullOrEmpty(entry)) continue;
                    int pipe = entry.IndexOf('|');
                    if (pipe < 1) continue;
                    string token = entry.Substring(0, pipe);
                    string value = entry.Substring(pipe + 1);
                    if (text.Contains(token)) text = text.Replace(token, value);
                }
            }

            if (text.Contains("%INPUT%"))
            {
                var input = new UserInput();
                input.ShowDialog();
                text = text.Replace("%INPUT%", input.inputReturn);

            }
            if (text.Contains("%SETTINGS%"))
            {
                this.Height = 365;
                this.Width = 446;
                this.CenterToScreen();
                text = "";
                return text;
            }

            return text;
        }

        private void QuickMenu_FormClosed(object sender, FormClosedEventArgs e)
        {
            notifyIcon1.Dispose();
        }

        private void icoStrip_Opening(object sender, CancelEventArgs e)
        {

        }

        private void quickMenuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - qMenuStrip.Height;
            qMenuStrip.Show(0, y);
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Height = 365;
            this.Width = 446;
            this.CenterToScreen();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void gbActionMenu_Enter(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tabControl2.SelectedTab = tabGeneral;
        }

        private void btnShowActionMenu_Click(object sender, EventArgs e)
        {
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - qMenuStrip.Height;
            qMenuStrip.Show(0, y);
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }

        private void GenerateActionMenu()
        {
            var actionMenu = Settings.Default.settingsActionList.Cast<string>().ToList();
            qMenuStrip.Items.Clear();

            foreach (var lstItem in actionMenu)
            {
                try
                {
                    string lData = lstItem;
                    string[] Splitter = lData.Split(char.Parse("ç"));
                    string aLabel = Splitter[0];
                    string aURL = Splitter[1];
                    Keys aKey = Keys.None;
                    if (Splitter[2] == "-")
                    {
                    }
                    else
                    {
                        aKey = (Keys)Enum.Parse(typeof(Keys), Splitter[2]);
                    }
                    string aType = Splitter[4];
                    string aArg = Splitter[3];

                    if (aURL == "%SEPARATOR%")
                    {
                        createMenuItem("Seperator", Keys.None, "Seperator", "Seperator", "", "");
                    }
                    else 
                    {
                        createMenuItem(aLabel, aKey, aLabel, aType, aURL, aArg);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            ThemeManager.ApplyMenu(qMenuStrip);
        }

        private void GenerateClipboardMenu()
        {
            var actionMenu = Settings.Default.settingClipboardList.Cast<string>().ToList();

            cMenuStrip.Items.Clear();

            foreach (var lstItem in actionMenu)
            {
                try
                {
                    string lData = lstItem;
                    string[] Splitter = lData.Split(char.Parse("ç"));
                    string aLabel = Splitter[0];
                    string aStringOrPath = Splitter[1];
                    Keys aKey = Keys.None;
                    if (Splitter[2] == "-")
                    {
                    }
                    else
                    {
                        aKey = (Keys)Enum.Parse(typeof(Keys), Splitter[2]);
                    }
                    string aType = Splitter[3];

                    if (aStringOrPath == "%SEPARATOR%")
                    {
                        createCMenuItem("Seperator", Keys.None, "Seperator", "Seperator", "");
                    }
                    else
                    {
                        createCMenuItem(aLabel, aKey, aLabel, aType, aStringOrPath);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            ThemeManager.ApplyMenu(cMenuStrip);
        }

        private void btnUp_Click(object sender, EventArgs e)
        {

            if (lbActionMenu.SelectedItem == null || lbActionMenu.SelectedIndex < 0)
                return;

            int newIndex = lbActionMenu.SelectedIndex + -1;

            if (newIndex < 0 || newIndex >= lbActionMenu.Items.Count)
                return;

            var list = Settings.Default.settingsActionList.Cast<string>().ToList();

            int selected = lbActionMenu.SelectedIndex;
            var selData = list[selected];

            

            Settings.Default.settingsActionList.RemoveAt(selected);
            Settings.Default.settingsActionList.Insert(newIndex, selData);

           


            Settings.Default.Save();

            var list2 = Settings.Default.settingsActionList.Cast<string>().ToList();
            lbActionMenu.Items.Clear();
            listLoader(lbActionMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();

            lbActionMenu.SetSelected(newIndex, true);
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (lbActionMenu.SelectedItem == null || lbActionMenu.SelectedIndex < 0)
                return;

            int newIndex = lbActionMenu.SelectedIndex + 1;

            if (newIndex < 0 || newIndex >= lbActionMenu.Items.Count)
                return;

            var list = Settings.Default.settingsActionList.Cast<string>().ToList();

            int selected = lbActionMenu.SelectedIndex;
            var selData = list[selected];



            Settings.Default.settingsActionList.RemoveAt(selected);
            Settings.Default.settingsActionList.Insert(newIndex, selData);
            

            Settings.Default.Save();

            var list2 = Settings.Default.settingsActionList.Cast<string>().ToList();
            lbActionMenu.Items.Clear();
            listLoader(lbActionMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();
            lbActionMenu.SetSelected(newIndex, true);
        }

        private void txtShellKey_Keypress(object sender, KeyPressEventArgs e)
        {
            txtShellKey.Text = "";
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void txtURLKey_Keypress(object sender, KeyPressEventArgs e)
        {

            txtURLKey.Text = "";
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void btnActionURL_Click(object sender, EventArgs e)
        {

            Settings.Default.settingsActionList.Add(txtURLLabel.Text + "ç" + txtURL.Text + "ç" + txtURLKey.Text + "ç" + "" + "ç" + "Link");
            Settings.Default.Save();
            lbActionMenu.Items.Clear();
            var list = Settings.Default.settingsActionList.Cast<string>().ToList();
            listLoader(lbActionMenu, list);
            GenerateActionMenu();
            GenerateClipboardMenu();

            //IF It's selected be sure to ask if they want to edit or make a new item. 
            //ASK For the menu text here
            //IF Hotkey isnt defined, ask for that too. 
            //if url isnt defined, ask for that as well. 
        }

        private void btnActionShell_Click(object sender, EventArgs e)
        {
            Settings.Default.settingsActionList.Add(txtShellLabel.Text + "ç" + txtShell.Text + "ç" + txtShellKey.Text + "ç" + txtShellArg.Text + "ç" + "Shell");
            Settings.Default.Save();
            lbActionMenu.Items.Clear();
            var list = Settings.Default.settingsActionList.Cast<string>().ToList();
            listLoader(lbActionMenu, list);
            GenerateActionMenu();
            GenerateClipboardMenu();
        }

        private void btnSetActionHotkey_Click(object sender, EventArgs e)
        {
            var hotkey1 = new SetHotKey();
            hotkey1.ShowDialog();
            if (hotkey1.hkModkeys == string.Empty)
            {

            }
            else
            {
                hook.Dispose();
                KeyboardHook reDefineHook = new KeyboardHook();
                reDefineHook.KeyPressed +=
                 new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);

                string[] Splitter = hotkey1.hkModkeys.Split(char.Parse(","));
                string Modkey1 = Splitter[0];
                string Modkey2 = Splitter[1];

                ModKeys modKeys1 = ModKeys.Shift;
                ModKeys modKeys2 = ModKeys.Control;
                if (Modkey1.Contains("Shift")) modKeys1 = ModKeys.Shift;
                if (Modkey1.Contains("Alt")) modKeys1 = ModKeys.Alt;
                if (Modkey1.Contains("Control")) modKeys1 = ModKeys.Control;
                if (Modkey2.Contains("Shift")) modKeys2 = ModKeys.Shift;
                if (Modkey2.Contains("Alt")) modKeys2 = ModKeys.Alt;
                if (Modkey2.Contains("Control")) modKeys2 = ModKeys.Control;

                reDefineHook.RegisterHotKey(modKeys1 | modKeys2, hotkey1.hkKeypress);

                Settings.Default.settingsActionHotkey = hotkey1.hkModkeys + ", " + hotkey1.hkKeypress;
                textBox4.Text = Settings.Default.settingsActionHotkey;
                textBox4.Text = textBox4.Text.Replace(",", "+");
                textBox4.Text = textBox4.Text.ToUpper();
                Settings.Default.Save();
                Process.Start(Application.ExecutablePath);
                this.Close();
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select a script or Executable";
            ofd.Filter = "All Files(*.*) | *.*";
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                txtShell.Text = ofd.FileName;
            }
        }

        private void btnActionSeperator_Click(object sender, EventArgs e)
        {
            Settings.Default.settingsActionList.Add("---------------ç%SEPARATOR%ç-ç-ç-");
            Settings.Default.Save();
            lbActionMenu.Items.Clear();
            var list = Settings.Default.settingsActionList.Cast<string>().ToList();
            listLoader(lbActionMenu, list);
            //---------------|%SEPARATOR%|-
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void lbActionMenu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbActionMenu.SelectedItem == null || lbActionMenu.SelectedIndex < 0)
                return;

            var list = Settings.Default.settingsActionList.Cast<string>().ToList();

            int selected = lbActionMenu.SelectedIndex;
            

            try
            {
                var selData = list[selected];
                string[] Splitter = selData.Split(char.Parse("ç"));
                string aLabel = Splitter[0];
                string aURL = Splitter[1];
                string aKey = Splitter[2];
                string aArg = Splitter[3];
                string aType= Splitter[4];

                if (aType == "Link")
                {
                    tabControl1.SelectedTab = tabPage2;
                    gbLinks.BringToFront();
                    txtURL.Text = aURL;
                    txtURLKey.Text = aKey;
                    txtURLLabel.Text = aLabel;
                    txtURL.SelectionStart = txtURL.Text.Length;
                    txtURL.SelectionLength = 0;
                }
                if (aType == "Shell")
                {
                    tabControl1.SelectedTab = tabPage1;
                    gbShell.BringToFront();
                    txtShell.Text = aURL;
                    txtShellKey.Text = aKey;
                    txtShellLabel.Text = aLabel;
                    txtShellArg.Text = aArg;
                    txtShell.SelectionStart = txtShell.Text.Length;
                    txtShell.SelectionLength = 0;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void listLoader(ListBox listBox, List<string> settingList)
        {
            foreach(var lstItem in settingList)
{
                try
                {
                string lData = lstItem;
                string[] Splitter = lData.Split(char.Parse("ç"));
                string aLabel = Splitter[0];
                string aURL = Splitter[1];
                string aKey = Splitter[2];

                if (aURL == "%SEPARATOR%")
                    {
                        listBox.Items.Add(aLabel);
                    }
                else
                    {
                        listBox.Items.Add("[" + aKey + "] " + aLabel);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lbActionMenu.SelectedItem == null || lbActionMenu.SelectedIndex < 0)
                return;

            var list = Settings.Default.settingsActionList.Cast<string>().ToList();

            int selected = lbActionMenu.SelectedIndex;
            var selData = list[selected];

            Settings.Default.settingsActionList.RemoveAt(selected);

            Settings.Default.Save();

            var list2 = Settings.Default.settingsActionList.Cast<string>().ToList();
            lbActionMenu.Items.Clear();
            listLoader(lbActionMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();
        }

        private void gbGeneral_Enter(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void lbWildcards_SelectedIndexChanged(object sender, EventArgs e)
        {
            int idx = lbWildcards.SelectedIndex;
            if (idx == 0)
            {
                gbInput.BringToFront();
                return;
            }
            if (idx == 1)
            {
                gbSettings.BringToFront();
                return;
            }
            if (idx >= _builtInWildcardCount)
            {
                string token = lbWildcards.Items[idx] as string;
                string value = GetCustomWildcardValue(token) ?? string.Empty;
                ThemeManager.ShowCustomStringWildcard(gbString, token, value);
                gbString.BringToFront();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select a Rich Text Format File";
            ofd.Filter = "Rich Text Format (*.rtf)|*.rtf";// | All Files(*.*)|*.*";
            //openFileDialog1.Filter = "public key  (*.publ)|*.publ";
            //"txt files (*.txt)|*.txt|All files (*.*)|*.*"
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                txtRTFPath.Text = ofd.FileName;
            }
        }

        private void txtRTFKey_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtRTFKey.Text = "";
            e.KeyChar = char.ToUpper(e.KeyChar);
        }
        private void txtStringKey_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtStringKey.Text = "";
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void btnCUp_Click(object sender, EventArgs e)
        {
            if (lbClipboardMenu.SelectedItem == null || lbClipboardMenu.SelectedIndex < 0)
                return;

            int newIndex = lbClipboardMenu.SelectedIndex + -1;

            if (newIndex < 0 || newIndex >= lbClipboardMenu.Items.Count)
                return;

            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();

            int selected = lbClipboardMenu.SelectedIndex;
            var selData = list[selected];



            Settings.Default.settingClipboardList.RemoveAt(selected);
            Settings.Default.settingClipboardList.Insert(newIndex, selData);




            Settings.Default.Save();

            var list2 = Settings.Default.settingClipboardList.Cast<string>().ToList();
            lbClipboardMenu.Items.Clear();
            listLoader(lbClipboardMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();

            lbClipboardMenu.SetSelected(newIndex, true);
        }

        private void btnCDown_Click(object sender, EventArgs e)
        {
            if (lbClipboardMenu.SelectedItem == null || lbClipboardMenu.SelectedIndex < 0)
                return;

            int newIndex = lbClipboardMenu.SelectedIndex + 1;

            if (newIndex < 0 || newIndex >= lbClipboardMenu.Items.Count)
                return;

            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();

            int selected = lbClipboardMenu.SelectedIndex;
            var selData = list[selected];



            Settings.Default.settingClipboardList.RemoveAt(selected);
            Settings.Default.settingClipboardList.Insert(newIndex, selData);


            Settings.Default.Save();

            var list2 = Settings.Default.settingClipboardList.Cast<string>().ToList();
            lbClipboardMenu.Items.Clear();
            listLoader(lbClipboardMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();
            lbClipboardMenu.SetSelected(newIndex, true);
        }

        private void btnCDelete_Click(object sender, EventArgs e)
        {
            if (lbClipboardMenu.SelectedItem == null || lbClipboardMenu.SelectedIndex < 0)
                return;

            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();

            int selected = lbClipboardMenu.SelectedIndex;
            var selData = list[selected];

            Settings.Default.settingClipboardList.RemoveAt(selected);

            Settings.Default.Save();

            var list2 = Settings.Default.settingClipboardList.Cast<string>().ToList();
            lbClipboardMenu.Items.Clear();
            listLoader(lbClipboardMenu, list2);
            GenerateActionMenu();
            GenerateClipboardMenu();
        }

        private void btnClipboardSave_Click(object sender, EventArgs e)
        {
            Settings.Default.settingClipboardList.Add(txtRTFLabel.Text + "ç" + txtRTFPath.Text + "ç" + txtRTFKey.Text + "ç" + "RTF");
            Settings.Default.Save();
            lbClipboardMenu.Items.Clear();
            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();
            listLoader(lbClipboardMenu, list);
            GenerateActionMenu();
            GenerateClipboardMenu();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Settings.Default.settingClipboardList.Add(txtStringLabel.Text + "ç" + txtStringString.Text + "ç" + txtStringKey.Text + "ç" + "String");
            Settings.Default.Save();
            lbClipboardMenu.Items.Clear();
            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();
            listLoader(lbClipboardMenu, list);
            GenerateActionMenu();
            GenerateClipboardMenu();
        }

        private void btnClipboardSeperator_Click(object sender, EventArgs e)
        {
            Settings.Default.settingClipboardList.Add("---------------ç%SEPARATOR%ç-ç-ç-");
            Settings.Default.Save();
            lbClipboardMenu.Items.Clear();
            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();
            listLoader(lbClipboardMenu, list);
            //---------------|%SEPARATOR%|-
        }

        private void lbClipboardMenu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbClipboardMenu.SelectedItem == null || lbClipboardMenu.SelectedIndex < 0)
                return;

            var list = Settings.Default.settingClipboardList.Cast<string>().ToList();

            int selected = lbClipboardMenu.SelectedIndex;


            try
            {
                var selData = list[selected];
                string[] Splitter = selData.Split(char.Parse("ç"));
                string aLabel = Splitter[0];
                string aPathorString = Splitter[1];
                string aKey = Splitter[2];
                string aType = Splitter[3];
                if (aType == "RTF")
                {
                    tabControl3.SelectedTab = tabPage3;
                    txtRTFPath.Text = aPathorString;
                    txtRTFKey.Text = aKey;
                    txtRTFLabel.Text = aLabel;
                    txtRTFPath.SelectionStart = txtRTFPath.Text.Length;
                    txtRTFPath.SelectionLength = 0;
                }
                if (aType == "String")
                {
                    tabControl3.SelectedTab = tabPage4;

                    txtStringString.Text = aPathorString;
                    txtStringKey.Text = aKey;
                    txtStringLabel.Text = aLabel;
                    txtStringString.SelectionStart = txtStringString.Text.Length;
                    txtStringString.SelectionLength = 0;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnShowClipboardMenu_Click(object sender, EventArgs e)
        {
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - cMenuStrip.Height;
            cMenuStrip.Show(0, y);
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }

        private void btnSetClipboardHotkey_Click(object sender, EventArgs e)
        {
            var hotkey = new SetHotKey();
            hotkey.ShowDialog();
            hook2.Dispose();
            if (hotkey.hkModkeys == string.Empty)
            {

            }
            else
            {
                KeyboardHook reDefineHook2 = new KeyboardHook();
                reDefineHook2.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(hook2_KeyPressed);

                string[] Splitter = hotkey.hkModkeys.Split(char.Parse(","));
                string Modkey1 = Splitter[0];
                string Modkey2 = Splitter[1];

                ModKeys modKeys1 = ModKeys.Shift;
                ModKeys modKeys2 = ModKeys.Control;
                if (Modkey1.Contains("Shift")) modKeys1 = ModKeys.Shift;
                if (Modkey1.Contains("Alt")) modKeys1 = ModKeys.Alt;
                if (Modkey1.Contains("Control")) modKeys1 = ModKeys.Control;
                if (Modkey2.Contains("Shift")) modKeys2 = ModKeys.Shift;
                if (Modkey2.Contains("Alt")) modKeys2 = ModKeys.Alt;
                if (Modkey2.Contains("Control")) modKeys2 = ModKeys.Control;

                reDefineHook2.RegisterHotKey(modKeys1 | modKeys2, hotkey.hkKeypress);

                Settings.Default.settingsClipboardHotkey = hotkey.hkModkeys + ", " + hotkey.hkKeypress;
                textBox1.Text = Settings.Default.settingsClipboardHotkey;
                textBox1.Text = textBox1.Text.Replace(",", "+");
                textBox1.Text = textBox1.Text.ToUpper();
                Settings.Default.Save();
                Process.Start(Application.ExecutablePath);
                this.Close();
            }

        }

        private void tabClipboard_Click(object sender, EventArgs e)
        {

        }
    }
}
