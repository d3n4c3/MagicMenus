using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MagicMenus
{
    /// <summary>
    /// "About Magic Menus" dialog shown from the tray icon's right-click
    /// menu. Displays the assembly version, author, and a clickable link
    /// to the author's site. Styled to match the rest of the application
    /// (custom dark/light title bar, accent palette, modal).
    /// </summary>
    public partial class AboutForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        private const string AuthorName = "Dennis Mayer";
        private const string AuthorUrl = "https://d3n4c3.com";
        private const string AuthorUrlDisplay = "d3n4c3.com";
        private const string GitHubUrl = "https://github.com/d3n4c3/MagicMenus";
        private const string GitHubUrlDisplay = "github.com/d3n4c3/MagicMenus";

        private Panel _titleBar;
        private Label _titleLabel;
        private Label _closeLabel;
        // Tracked so ApplyTheme() can re-apply accent / subtle colors that
        // ThemeManager.Apply would otherwise overwrite with p.Text.
        private PictureBox _brand;
        private Label _version;
        private Label _copyright;
        private LinkLabel _siteLink;
        private LinkLabel _githubLink;

        public AboutForm()
        {
            BuildLayout();
            ApplyTheme();
        }

        private void BuildLayout()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            Size = new Size(400, 360);
            Text = "About Magic Menus";
            KeyPreview = true;
            KeyDown += delegate (object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter)
                {
                    Close();
                    e.SuppressKeyPress = true;
                }
            };

            // ----- Title bar (matches TutorialForm / main settings dialog).
            _titleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 32,
                Tag = ThemeManager.TagTitleBar
            };
            _titleBar.MouseDown += TitleBar_MouseDown;

            _titleLabel = new Label
            {
                Text = "About Magic Menus",
                AutoSize = true,
                Location = new Point(14, 7),
                Tag = ThemeManager.TagTitleBar
            };
            _titleLabel.MouseDown += TitleBar_MouseDown;

            _closeLabel = new Label
            {
                Text = "X",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(36, 32),
                Dock = DockStyle.Right,
                Cursor = Cursors.Hand,
                Tag = ThemeManager.TagTitleBar
            };
            _closeLabel.Click += delegate { Close(); };
            _closeLabel.MouseEnter += delegate
            {
                _closeLabel.BackColor = ThemeManager.CloseHoverColor;
                _closeLabel.ForeColor = Color.White;
            };
            _closeLabel.MouseLeave += delegate
            {
                ThemeManager.Palette pal = ThemeManager.Current;
                _closeLabel.BackColor = pal.TitleBar;
                _closeLabel.ForeColor = pal.TitleText;
            };

            // Add close first so DockStyle.Right wins the rightmost slot.
            _titleBar.Controls.Add(_closeLabel);
            _titleBar.Controls.Add(_titleLabel);

            // ----- Content panel with the version / author / link block.
            ThemeManager.Palette p = ThemeManager.Current;
            int innerWidth = ClientSize.Width;

            // Project icon shown at its native 32x32 resolution so it
            // stays pixel-perfect instead of getting bicubic-blurred when
            // upscaled. ExtractAssociatedIcon returns a 32x32 icon on
            // Windows.
            _brand = new PictureBox
            {
                Size = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.Normal,
                Image = LoadBrandImage()
            };
            _brand.Location = new Point(0, 56);
            Controls.Add(_brand);

            Label productName = new Label
            {
                Text = "Magic Menus",
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 14F, FontStyle.Regular)
            };
            productName.Location = new Point(0, 100);
            Controls.Add(productName);

            _version = new Label
            {
                Text = "Version " + GetVersionDisplay(),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _version.Location = new Point(0, 128);
            Controls.Add(_version);

            Label authorLabel = new Label
            {
                Text = "Created by " + AuthorName,
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular)
            };
            authorLabel.Location = new Point(0, 168);
            Controls.Add(authorLabel);

            _siteLink = new LinkLabel
            {
                Text = AuthorUrlDisplay,
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                LinkBehavior = LinkBehavior.HoverUnderline,
                Cursor = Cursors.Hand
            };
            _siteLink.Location = new Point(0, 192);
            _siteLink.LinkClicked += delegate { OpenUrl(AuthorUrl); };
            Controls.Add(_siteLink);

            _githubLink = new LinkLabel
            {
                Text = GitHubUrlDisplay,
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                LinkBehavior = LinkBehavior.HoverUnderline,
                Cursor = Cursors.Hand
            };
            _githubLink.Location = new Point(0, 216);
            _githubLink.LinkClicked += delegate { OpenUrl(GitHubUrl); };
            Controls.Add(_githubLink);

            _copyright = new Label
            {
                Text = GetCopyrightDisplay(),
                AutoSize = true,
                Font = new Font("Segoe UI", 8F, FontStyle.Regular)
            };
            _copyright.Location = new Point(0, 252);
            Controls.Add(_copyright);

            // ----- Close button at the bottom right.
            Button closeBtn = new Button
            {
                Text = "Close",
                Size = new Size(90, 30),
                Cursor = Cursors.Hand,
                Tag = ThemeManager.TagPrimary
            };
            closeBtn.Click += delegate { Close(); };
            // Pinned to bottom-right; positioned in Layout() so it
            // accounts for the actual ClientSize.
            closeBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            closeBtn.Location = new Point(ClientSize.Width - closeBtn.Width - 18,
                                          ClientSize.Height - closeBtn.Height - 16);
            Controls.Add(closeBtn);

            // Center the labels horizontally once the form is shown.
            Shown += delegate
            {
                CenterHorizontally(_brand);
                CenterHorizontally(productName);
                CenterHorizontally(_version);
                CenterHorizontally(authorLabel);
                CenterHorizontally(_siteLink);
                CenterHorizontally(_githubLink);
                CenterHorizontally(_copyright);
            };
            // Make sure the title bar sits in front of any centered child.
            _titleBar.BringToFront();
            Controls.Add(_titleBar);
        }

        private void CenterHorizontally(Control c)
        {
            c.Left = (ClientSize.Width - c.Width) / 2;
        }

        private void ApplyTheme()
        {
            ThemeManager.Apply(this);
            ThemeManager.ApplyTitleBar(_titleBar, _titleLabel, _closeLabel);
            BackColor = ThemeManager.Current.FormBack;
            ApplyAccentColors();
        }

        /// <summary>
        /// ThemeManager.Apply paints every Label with p.Text and every
        /// LinkLabel with the accent. Re-apply the few overrides we want
        /// (accent "MM" mark, subtle version / copyright) here so the
        /// dialog reads like a marketing splash rather than a wall of
        /// uniform body copy.
        /// </summary>
        private void ApplyAccentColors()
        {
            ThemeManager.Palette p = ThemeManager.Current;
            if (_brand != null) _brand.BackColor = p.FormBack;
            if (_version != null) _version.ForeColor = p.SubtleText;
            if (_copyright != null) _copyright.ForeColor = p.SubtleText;
            StyleLink(_siteLink, p);
            StyleLink(_githubLink, p);
        }

        private static void StyleLink(LinkLabel link, ThemeManager.Palette p)
        {
            if (link == null) return;
            link.LinkColor = p.Accent;
            link.ActiveLinkColor = p.AccentHover;
            link.VisitedLinkColor = p.Accent;
        }

        private void OpenUrl(string url)
        {
            try { Process.Start(url); }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Couldn't open the link:\r\n" + ex.Message,
                    "Magic Menus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Extract the running executable's main icon at its native 32x32
        /// size. No scaling - the PictureBox displays the bitmap pixel-
        /// for-pixel so the artwork stays crisp. Returns null if
        /// extraction fails (in which case the PictureBox just stays
        /// empty).
        /// </summary>
        private static Image LoadBrandImage()
        {
            try
            {
                using (Icon raw = Icon.ExtractAssociatedIcon(Application.ExecutablePath))
                {
                    if (raw == null) return null;
                    return raw.ToBitmap();
                }
            }
            catch { return null; }
        }

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private static string GetVersionDisplay()
        {
            try
            {
                Version v = Assembly.GetExecutingAssembly().GetName().Version;
                if (v == null) return "1.0";
                // Trim trailing .0 components for a cleaner display
                // (1.0.0.0 -> 1.0).
                if (v.Revision == 0 && v.Build == 0) return v.Major + "." + v.Minor;
                if (v.Revision == 0) return v.Major + "." + v.Minor + "." + v.Build;
                return v.ToString();
            }
            catch { return "1.0"; }
        }

        private static string GetCopyrightDisplay()
        {
            try
            {
                object[] attrs = Assembly.GetExecutingAssembly()
                    .GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attrs != null && attrs.Length > 0)
                {
                    AssemblyCopyrightAttribute a = (AssemblyCopyrightAttribute)attrs[0];
                    if (!string.IsNullOrEmpty(a.Copyright)) return a.Copyright;
                }
            }
            catch { }
            return "Copyright " + AuthorName;
        }
    }
}
