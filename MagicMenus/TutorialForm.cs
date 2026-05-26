using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// The main form class is QuickMenu.MagicMenus, the surrounding namespace of
// this file is "MagicMenus" (matching the project name), so we alias the
// owner type once to keep the rest of the file readable.
using OwnerForm = QuickMenu.MagicMenus;

namespace MagicMenus
{
    /// <summary>
    /// First-run onboarding tutorial. Walks the user through Magic Menus'
    /// core concepts (menus, wildcards) and has them create their first two
    /// menu items hands-on: a Settings shortcut and a Google Search that
    /// uses %INPUT% to pass a query into a URL.
    ///
    /// The tutorial auto-opens on first launch (no menu items yet) and can
    /// be re-run anytime from the "Show tutorial" button on the General tab.
    /// </summary>
    public partial class TutorialForm : Form
    {
        // ----- Win32 plumbing for the custom borderless title bar drag. -----
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        // ----- Step definitions. Each step is described by a Tutorial step
        //       object that the renderer turns into a panel of controls. -----
        private enum StepKind
        {
            Info,
            ThemeChoice,
            ActionItemForm,
            ClipboardItemForm,
            HotkeyAndFinish
        }

        private class Step
        {
            public StepKind Kind;
            public string Title;
            public string Body;
            // For *ItemForm steps. ItemUrl doubles as the clipboard
            // content string for the clipboard form.
            public string ItemLabel;
            public string ItemUrl;
            public string ItemKey;
            public string ItemExplanation;
            // For HotkeyAndFinish steps:
            public string FinishHint;
        }

        private readonly OwnerForm _ownerForm;
        private readonly List<Step> _steps;
        private int _index;

        // Header / chrome controls (built once).
        private Panel _titleBar;
        private Label _titleLabel;
        private Label _closeLabel;
        private Panel _content;
        private Panel _footer;
        private Label _progressLabel;
        private Button _backButton;
        private Button _skipButton;
        private Button _nextButton;
        // A single-line status banner near the bottom of the content area
        // that confirms "Added Settings to your menu", etc.
        private Label _statusLabel;

        // Per-step interactive controls.
        // The "MenuItemForm" step displays pre-filled (read-only) fields so
        // the user sees exactly what's being created and clicks Add. The
        // "HotkeyAndFinish" step shows the configured action hotkey (or a
        // button to set one if there isn't one yet).

        public TutorialForm(OwnerForm owner)
        {
            _ownerForm = owner;
            _steps = BuildSteps();
            _index = 0;

            BuildLayout();
            ApplyInitialTheme();
            Render();
        }

        private static List<Step> BuildSteps()
        {
            List<Step> steps = new List<Step>();
            steps.Add(new Step
            {
                Kind = StepKind.Info,
                Title = "Welcome to Magic Menus",
                Body =
                    "Magic Menus lets you trigger custom commands from anywhere on Windows " +
                    "with a single hotkey. Open URLs, run scripts, paste boilerplate text, or " +
                    "even ask for input on the fly.\r\n\r\n" +
                    "Magic Menus lives in your system tray - right-click the icon at any time " +
                    "to open Settings, view About info, or quit the app.\r\n\r\n" +
                    "Let's set up your first menu in a few quick steps."
            });
            steps.Add(new Step
            {
                Kind = StepKind.ThemeChoice,
                Title = "Pick your look",
                Body =
                    "Magic Menus comes in two flavors. Click one to try it - you can switch " +
                    "anytime from Settings > General."
            });
            steps.Add(new Step
            {
                Kind = StepKind.Info,
                Title = "Two ideas to know",
                Body =
                    "MENU ITEM - a single command on your menu (a URL to open, a program " +
                    "to run, text to paste).\r\n\r\n" +
                    "WILDCARD - a placeholder like %INPUT% or %SETTINGS% that Magic Menus " +
                    "expands when you invoke the item. %INPUT% prompts you for text at " +
                    "runtime; %SETTINGS% opens the Settings dialog. You can define your own " +
                    "custom wildcards on the General tab."
            });
            steps.Add(new Step
            {
                Kind = StepKind.ActionItemForm,
                Title = "Action menu: Settings shortcut",
                Body =
                    "Add an item that opens the Settings dialog from your menu. The fields " +
                    "are pre-filled - just click Add to Menu.",
                ItemLabel = "Settings",
                ItemUrl = "%SETTINGS%",
                ItemKey = "S",
                ItemExplanation =
                    "%SETTINGS% is a built-in wildcard. When you select this menu item, " +
                    "Magic Menus opens the Settings dialog instead of trying to launch a URL " +
                    "- handy when you need to edit menu items from anywhere."
            });
            steps.Add(new Step
            {
                Kind = StepKind.ActionItemForm,
                Title = "Action menu: Google Search",
                Body =
                    "Now build something powerful: a Google Search that prompts you for a " +
                    "query and then opens the results in your browser.",
                ItemLabel = "Google Search",
                ItemUrl = "https://www.google.com/search?q=%INPUT%",
                ItemKey = "G",
                ItemExplanation =
                    "Notice %INPUT% inside the URL. When you select this menu item, Magic " +
                    "Menus pops up a small text box, takes whatever you type, and substitutes " +
                    "it into the URL before opening it. The same trick works for any program " +
                    "that takes its query on the command line."
            });
            steps.Add(new Step
            {
                Kind = StepKind.ClipboardItemForm,
                Title = "Clipboard menu: Shrug",
                Body =
                    "The clipboard menu pastes prebuilt snippets. Add a shrug emoji - perfect " +
                    "for chat when words fail you.",
                ItemLabel = "Shrug",
                // Pure-Unicode shrug so the file stays ASCII-safe.
                ItemUrl = "\u00AF\\_(\u30C4)_/\u00AF",
                ItemKey = "R",
                ItemExplanation =
                    "When you pick this from the clipboard menu, Magic Menus copies the text " +
                    "to your clipboard so you can paste it anywhere with Ctrl+V. " +
                    "Bonus: try (\u256F\u00B0\u25A1\u00B0)\u256F\uFE35 \u253B\u2501\u253B " +
                    "for the full table-flip experience."
            });
            steps.Add(new Step
            {
                Kind = StepKind.HotkeyAndFinish,
                Title = "Pick your hotkeys, then try it",
                Body =
                    "Each menu opens whenever you press its hotkey - anywhere on Windows, " +
                    "even when Magic Menus isn't focused.",
                FinishHint =
                    "Once your hotkeys are set:\r\n" +
                    "  - Press the action hotkey, then G, to search Google.\r\n" +
                    "  - Press the clipboard hotkey, then R, to copy the shrug onto your " +
                    "clipboard - paste it anywhere with Ctrl+V.\r\n\r\n" +
                    "Tip: right-click the tray icon to open Settings, see About info, or " +
                    "quit. You can re-run this tutorial anytime from Settings > General."
            });
            return steps;
        }

        // -----------------------------------------------------------------
        // Layout (built once on construction).
        // -----------------------------------------------------------------

        private void BuildLayout()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            Size = new Size(540, 480);
            Text = "Magic Menus Tutorial";
            KeyPreview = true;
            KeyDown += new KeyEventHandler(TutorialForm_KeyDown);
            // Sit above other application windows so the user doesn't lose
            // the tutorial when they Alt-Tab or click off it accidentally.
            TopMost = true;

            // Title bar.
            _titleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 32,
                Tag = ThemeManager.TagTitleBar
            };
            _titleBar.MouseDown += new MouseEventHandler(TitleBar_MouseDown);

            _titleLabel = new Label
            {
                Text = "Magic Menus Tutorial",
                AutoSize = true,
                Location = new Point(14, 7),
                Tag = ThemeManager.TagTitleBar
            };
            _titleLabel.MouseDown += new MouseEventHandler(TitleBar_MouseDown);

            _closeLabel = new Label
            {
                Text = "X",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(36, 32),
                // Pin to the right side of the title bar so it lines up
                // with the X on the main MagicMenus settings window.
                Dock = DockStyle.Right,
                Cursor = Cursors.Hand,
                Tag = ThemeManager.TagTitleBar
            };
            _closeLabel.Click += new EventHandler(Close_Click);
            _closeLabel.MouseEnter += new EventHandler(CloseLabel_MouseEnter);
            _closeLabel.MouseLeave += new EventHandler(CloseLabel_MouseLeave);

            // Add the close label first so DockStyle.Right wins the rightmost
            // slot before the title label is laid out at its fixed location.
            _titleBar.Controls.Add(_closeLabel);
            _titleBar.Controls.Add(_titleLabel);

            // Footer with navigation.
            _footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 54
            };

            _progressLabel = new Label
            {
                AutoSize = true,
                Location = new Point(18, 19),
                Tag = ThemeManager.TagSubtle
            };

            _backButton = new Button
            {
                Text = "Back",
                Size = new Size(80, 30),
                Cursor = Cursors.Hand,
                Tag = "secondary"
            };
            _backButton.Click += new EventHandler(Back_Click);

            _skipButton = new Button
            {
                Text = "Skip tutorial",
                Size = new Size(100, 30),
                Cursor = Cursors.Hand,
                Tag = "secondary"
            };
            _skipButton.Click += new EventHandler(Skip_Click);

            _nextButton = new Button
            {
                Text = "Get Started",
                Size = new Size(120, 30),
                Cursor = Cursors.Hand,
                Tag = ThemeManager.TagPrimary
            };
            _nextButton.Click += new EventHandler(Next_Click);

            _footer.Controls.Add(_progressLabel);
            _footer.Controls.Add(_backButton);
            _footer.Controls.Add(_skipButton);
            _footer.Controls.Add(_nextButton);
            _footer.Resize += new EventHandler(Footer_Resize);

            // Content area (filled per step).
            _content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(28, 22, 28, 12),
                AutoScroll = false
            };

            Controls.Add(_content);
            Controls.Add(_footer);
            Controls.Add(_titleBar);

            // Status label (created once, repositioned per step).
            _statusLabel = new Label
            {
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = false,
                Tag = ThemeManager.TagSubtle
            };
            _content.Controls.Add(_statusLabel);
        }

        private void Footer_Resize(object sender, EventArgs e)
        {
            // Right-align the primary action, then place Back to its left,
            // with Skip pinned to the bottom-left.
            int right = _footer.ClientSize.Width - 18;
            _nextButton.Location = new Point(right - _nextButton.Width, 12);
            _backButton.Location = new Point(_nextButton.Left - _backButton.Width - 8, 12);
            _skipButton.Location = new Point(18, 12);
            _progressLabel.Location = new Point(_skipButton.Right + 16, 19);
        }

        private void ApplyInitialTheme()
        {
            // Use the same dark theme as the main form. ThemeManager paints
            // GroupBox/TabControl chrome itself, but for our simple Panels/
            // Labels we set colors explicitly.
            ThemeManager.Apply(this);
            ApplyThemeChromeColors();
        }

        private void ApplyThemeChromeColors()
        {
            ThemeManager.ApplyTitleBar(_titleBar, _titleLabel, _closeLabel);
            BackColor = ThemeManager.Current.FormBack;
            _footer.BackColor = ThemeManager.Current.FormBack;
            _content.BackColor = ThemeManager.Current.FormBack;
        }

        // -----------------------------------------------------------------
        // Per-step rendering.
        // -----------------------------------------------------------------

        private void Render()
        {
            ClearContent();
            _statusLabel.Visible = false;

            Step step = _steps[_index];
            ThemeManager.Palette p = ThemeManager.Current;

            Label headerNumber = new Label
            {
                Text = "STEP " + (_index + 1) + " OF " + _steps.Count,
                AutoSize = true,
                Location = new Point(0, 0),
                ForeColor = p.Accent,
                Font = new Font("Segoe UI Semibold", 7.5F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            _content.Controls.Add(headerNumber);

            Label title = new Label
            {
                Text = step.Title,
                AutoSize = true,
                Location = new Point(0, 18),
                ForeColor = p.Text,
                Font = new Font("Segoe UI Semibold", 14F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            _content.Controls.Add(title);

            int bodyTop = title.Bottom + 8;
            int contentRight = _content.ClientSize.Width - _content.Padding.Horizontal;
            Label body = new Label
            {
                Text = step.Body,
                AutoSize = false,
                Location = new Point(0, bodyTop),
                Size = new Size(contentRight, 0),
                ForeColor = p.SubtleText,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            Size bodySize = TextRenderer.MeasureText(step.Body, body.Font,
                new Size(contentRight, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            body.Size = new Size(contentRight, bodySize.Height + 8);
            _content.Controls.Add(body);

            int interactiveTop = body.Bottom + 14;
            if (step.Kind == StepKind.ThemeChoice)
            {
                interactiveTop = RenderThemeChoice(step, interactiveTop, contentRight);
            }
            else if (step.Kind == StepKind.ActionItemForm ||
                     step.Kind == StepKind.ClipboardItemForm)
            {
                interactiveTop = RenderItemForm(step, interactiveTop, contentRight);
            }
            else if (step.Kind == StepKind.HotkeyAndFinish)
            {
                interactiveTop = RenderHotkeyAndFinish(step, interactiveTop, contentRight);
            }

            // Status label - reserve space at the bottom of the content
            // area but only show it once we have something to say.
            _statusLabel.Bounds = new Rectangle(0, interactiveTop + 8, contentRight, 20);

            // Footer state.
            _progressLabel.Text = "Step " + (_index + 1) + " of " + _steps.Count;
            _backButton.Visible = _index > 0;
            _skipButton.Visible = _index < _steps.Count - 1;
            _nextButton.Text = NextButtonCaption();
            // Some step captions ("Add to clipboard menu") are longer than
            // the default 120px next-button slot. Stretch the button to fit
            // its current caption so nothing ever gets clipped.
            Size captionSize = TextRenderer.MeasureText(_nextButton.Text, _nextButton.Font);
            _nextButton.Width = Math.Max(120, captionSize.Width + 32);
            Footer_Resize(this, EventArgs.Empty);
        }

        private string NextButtonCaption()
        {
            Step step = _steps[_index];
            if (_index == 0) return "Get started";
            if (step.Kind == StepKind.ActionItemForm) return "Add to menu";
            if (step.Kind == StepKind.ClipboardItemForm) return "Add to clipboard menu";
            if (_index == _steps.Count - 1) return "Finish";
            return "Next";
        }

        private void ClearContent()
        {
            // Detach status label first so we can re-add it without disposal.
            if (_statusLabel != null && _content.Controls.Contains(_statusLabel))
                _content.Controls.Remove(_statusLabel);

            List<Control> toDispose = new List<Control>();
            foreach (Control c in _content.Controls)
                toDispose.Add(c);
            foreach (Control c in toDispose)
            {
                _content.Controls.Remove(c);
                c.Dispose();
            }

            // Re-add status label so it sticks around between renders.
            _content.Controls.Add(_statusLabel);
        }

        /// <summary>
        /// Render two side-by-side "swatch" buttons that switch the active
        /// palette live. Clicking one updates both this tutorial and the
        /// settings dialog behind it, then re-renders so the just-clicked
        /// option shows its selected ring.
        /// </summary>
        private int RenderThemeChoice(Step step, int top, int contentWidth)
        {
            ThemeManager.Palette p = ThemeManager.Current;

            int gap = 14;
            int cardWidth = (contentWidth - gap) / 2;
            int cardHeight = 96;

            Button lightCard = BuildThemeCard(false, 0, top, cardWidth, cardHeight);
            Button darkCard = BuildThemeCard(true, cardWidth + gap, top, cardWidth, cardHeight);

            // Intentionally NOT calling ThemeManager.RestyleControl on the
            // cards: the default Button styling would overwrite the swatch
            // background/foreground we set in BuildThemeCard with the
            // active palette's ButtonBack/Text. The cards already have
            // FlatStyle.Flat plus their own colors set in BuildThemeCard.
            _content.Controls.Add(lightCard);
            _content.Controls.Add(darkCard);

            return top + cardHeight + 8;
        }

        private Button BuildThemeCard(bool dark, int left, int top, int width, int height)
        {
            ThemeManager.Palette p = ThemeManager.Current;
            bool selected = (dark == ThemeManager.IsDark);
            // Tint the card differently per theme so users can see at a
            // glance which is which - the active palette also gets a
            // ring (accent border) so the selection is unambiguous.
            Color cardBack = dark
                ? Color.FromArgb(30, 30, 30)
                : Color.FromArgb(245, 245, 245);
            Color cardFore = dark ? Color.White : Color.FromArgb(20, 20, 20);
            Button card = new Button
            {
                Text = (dark ? "Dark" : "Light") + (selected ? "  (current)" : ""),
                Location = new Point(left, top),
                Size = new Size(width, height),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                BackColor = cardBack,
                ForeColor = cardFore,
                Font = new Font("Segoe UI Semibold", 12F, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                UseVisualStyleBackColor = false,
                // Mark with our own tag so the theme manager doesn't paint
                // over our colors with the default button styling - we
                // want these to look like swatches, not buttons.
                Tag = "theme-card"
            };
            card.FlatAppearance.BorderSize = selected ? 2 : 1;
            card.FlatAppearance.BorderColor = selected ? p.Accent : p.ButtonBorder;
            card.FlatAppearance.MouseOverBackColor = dark
                ? Color.FromArgb(45, 45, 45)
                : Color.FromArgb(230, 230, 230);
            card.Click += delegate { ApplyTheme(dark); };
            return card;
        }

        private void ApplyTheme(bool dark)
        {
            if (ThemeManager.IsDark == dark) return;
            ThemeManager.IsDark = dark;
            // Persist immediately and refresh the settings dialog behind us
            // so the change is visible on both windows at once.
            try
            {
                Properties.Settings.Default.darkMode = dark;
                Properties.Settings.Default.Save();
            }
            catch { }
            if (_ownerForm != null) _ownerForm.RefreshTheme();
            // Re-apply to this form and re-render the step so the swatch
            // ring updates to point at the newly-selected card.
            ThemeManager.Apply(this);
            ApplyThemeChromeColors();
            Render();
        }

        private int RenderItemForm(Step step, int top, int contentWidth)
        {
            ThemeManager.Palette p = ThemeManager.Current;

            int labelColWidth = 78;
            // Clipboard items show "Content"; action items show "URL".
            string valueRowLabel = step.Kind == StepKind.ClipboardItemForm
                ? "Content" : "URL";

            top = AddFormRow(top, labelColWidth, contentWidth, "Label", step.ItemLabel, false);
            top = AddFormRow(top, labelColWidth, contentWidth, valueRowLabel, step.ItemUrl, true);
            top = AddFormRow(top, labelColWidth, contentWidth, "Shortcut", step.ItemKey, false);

            top += 6;
            Label explain = new Label
            {
                Text = step.ItemExplanation,
                AutoSize = false,
                Location = new Point(0, top),
                Size = new Size(contentWidth, 0),
                ForeColor = p.SubtleText,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                BackColor = p.FormBack
            };
            Size es = TextRenderer.MeasureText(step.ItemExplanation, explain.Font,
                new Size(contentWidth, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            explain.Size = new Size(contentWidth, es.Height + 4);
            _content.Controls.Add(explain);

            return explain.Bottom;
        }

        private int AddFormRow(int top, int labelColWidth, int contentWidth,
            string label, string value, bool monospace)
        {
            ThemeManager.Palette p = ThemeManager.Current;
            Label lbl = new Label
            {
                Text = label,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleRight,
                Location = new Point(0, top),
                Size = new Size(labelColWidth, 22),
                ForeColor = p.SubtleText,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            _content.Controls.Add(lbl);

            // Use a Panel + Label for the read-only value to avoid the
            // text box caret. The accent border tells the eye "this is a
            // configured field" without making it look editable.
            Panel valuePanel = new Panel
            {
                Location = new Point(labelColWidth + 12, top),
                Size = new Size(contentWidth - labelColWidth - 12, 24),
                BackColor = p.InputBack,
                BorderStyle = BorderStyle.FixedSingle
            };
            Label valueLabel = new Label
            {
                Text = value,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 8, 0),
                ForeColor = monospace ? p.Accent : p.Text,
                BackColor = p.InputBack,
                Font = monospace
                    ? new Font("Consolas", 9F, FontStyle.Regular)
                    : new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            valuePanel.Controls.Add(valueLabel);
            _content.Controls.Add(valuePanel);

            return top + 32;
        }

        private int RenderHotkeyAndFinish(Step step, int top, int contentWidth)
        {
            ThemeManager.Palette p = ThemeManager.Current;

            top = AddHotkeyRow(top, contentWidth, "Action menu hotkey",
                Properties.Settings.Default.settingsActionHotkey, true);
            top = AddHotkeyRow(top + 12, contentWidth, "Clipboard menu hotkey",
                Properties.Settings.Default.settingsClipboardHotkey, false);

            int finishTop = top + 18;
            Label hint = new Label
            {
                Text = step.FinishHint,
                AutoSize = false,
                Location = new Point(0, finishTop),
                Size = new Size(contentWidth, 0),
                ForeColor = p.SubtleText,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            Size hs = TextRenderer.MeasureText(step.FinishHint, hint.Font,
                new Size(contentWidth, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            hint.Size = new Size(contentWidth, hs.Height + 4);
            _content.Controls.Add(hint);

            return hint.Bottom;
        }

        /// <summary>
        /// Render a "Name: [VALUE]  [Change...]" row used twice in the final
        /// step: once for the action menu hotkey, once for the clipboard
        /// menu hotkey. Returns the y coordinate just below the row.
        /// </summary>
        private int AddHotkeyRow(int top, int contentWidth, string rowName,
            string rawValue, bool isActionMenu)
        {
            ThemeManager.Palette p = ThemeManager.Current;
            string hk = (rawValue ?? string.Empty).Trim();

            Label rowLabel = new Label
            {
                Text = rowName,
                AutoSize = true,
                Location = new Point(0, top),
                ForeColor = p.SubtleText,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                BackColor = p.FormBack
            };
            _content.Controls.Add(rowLabel);

            int rowTop = top + 20;
            Panel hkBox = new Panel
            {
                Location = new Point(0, rowTop),
                Size = new Size(260, 30),
                BackColor = p.InputBack,
                BorderStyle = BorderStyle.FixedSingle
            };
            Label hkValue = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Regular),
                BackColor = p.InputBack
            };
            if (string.IsNullOrEmpty(hk))
            {
                hkValue.Text = "(not set)";
                hkValue.ForeColor = p.SubtleText;
            }
            else
            {
                hkValue.Text = hk.Replace(",", " + ").ToUpperInvariant();
                hkValue.ForeColor = p.Accent;
            }
            hkBox.Controls.Add(hkValue);
            _content.Controls.Add(hkBox);

            Button setBtn = new Button
            {
                Text = string.IsNullOrEmpty(hk) ? "Set hotkey..." : "Change hotkey...",
                Location = new Point(hkBox.Right + 10, rowTop),
                Size = new Size(130, 30),
                Cursor = Cursors.Hand,
                Tag = ThemeManager.TagPrimary
            };
            setBtn.Click += delegate
            {
                bool changed = isActionMenu
                    ? (_ownerForm != null && _ownerForm.SetActionHotkeyInteractively())
                    : (_ownerForm != null && _ownerForm.SetClipboardHotkeyInteractively());
                if (changed) Render();
            };
            _content.Controls.Add(setBtn);
            ThemeManager.RestyleControl(setBtn);

            return hkBox.Bottom;
        }

        private void ShowStatus(string text, bool good)
        {
            ThemeManager.Palette p = ThemeManager.Current;
            _statusLabel.Text = (good ? "OK  " : "")  + text;
            _statusLabel.ForeColor = good ? p.Accent : p.SubtleText;
            _statusLabel.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Regular);
            _statusLabel.Visible = true;
            _statusLabel.BringToFront();
        }

        // -----------------------------------------------------------------
        // Navigation.
        // -----------------------------------------------------------------

        private void Next_Click(object sender, EventArgs e)
        {
            Step step = _steps[_index];

            if (step.Kind == StepKind.ActionItemForm)
            {
                bool added = TryAddActionItem(step.ItemLabel, step.ItemUrl, step.ItemKey);
                if (!added) return; // ShowStatus already explained why
            }
            else if (step.Kind == StepKind.ClipboardItemForm)
            {
                bool added = TryAddClipboardItem(step.ItemLabel, step.ItemUrl, step.ItemKey, "String");
                if (!added) return;
            }

            if (_index >= _steps.Count - 1)
            {
                MarkComplete();
                Close();
                return;
            }

            _index++;
            Render();
        }

        private void Back_Click(object sender, EventArgs e)
        {
            if (_index <= 0) return;
            _index--;
            Render();
        }

        private void Skip_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this,
                "Skip the rest of the tutorial?\r\n\r\n" +
                "You can re-run it anytime from Settings > General.",
                "Skip tutorial",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                MarkComplete();
                Close();
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            // Closing via the X is treated the same as Skip - mark complete
            // so we don't pester them on next launch.
            MarkComplete();
            Close();
        }

        private void MarkComplete()
        {
            try
            {
                Properties.Settings.Default.tutorialCompleted = true;
                Properties.Settings.Default.Save();
            }
            catch { }
        }

        /// <summary>
        /// Programmatically advance to the next step without firing any
        /// side effects (no menu items added, no settings saved). Used by
        /// the screenshot harness so each tutorial page can be captured.
        /// </summary>
        public void AdvanceForScreenshot()
        {
            if (_index >= _steps.Count - 1) return;
            _index++;
            Render();
        }

        private bool TryAddActionItem(string label, string url, string key)
        {
            string type = "Link";
            string arg = string.Empty;
            string entry = label + "ç" + url + "ç" + key + "ç" + arg + "ç" + type;

            // Skip if an item with the same Label or the same URL already
            // exists - re-running the tutorial shouldn't duplicate rows.
            System.Collections.Specialized.StringCollection list =
                Properties.Settings.Default.settingsActionList
                ?? new System.Collections.Specialized.StringCollection();

            foreach (string existing in list)
            {
                if (string.IsNullOrEmpty(existing)) continue;
                string[] parts = existing.Split('ç');
                if (parts.Length >= 2 &&
                    (string.Equals(parts[0], label, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(parts[1], url, StringComparison.OrdinalIgnoreCase)))
                {
                    ShowStatus("Already in your menu - moving on.", true);
                    if (_ownerForm != null) _ownerForm.RefreshActionMenuUI();
                    return true;
                }
            }

            list.Add(entry);
            Properties.Settings.Default.settingsActionList = list;
            Properties.Settings.Default.Save();
            if (_ownerForm != null) _ownerForm.RefreshActionMenuUI();
            ShowStatus("Added \"" + label + "\" to your menu.", true);
            return true;
        }

        private bool TryAddClipboardItem(string label, string content, string key, string type)
        {
            // Clipboard menu entries use the 4-field schema:
            //   Label ç Content ç Keypress ç Type   (no Arguments slot).
            string entry = label + "ç" + content + "ç" + key + "ç" + type;

            System.Collections.Specialized.StringCollection list =
                Properties.Settings.Default.settingClipboardList
                ?? new System.Collections.Specialized.StringCollection();

            foreach (string existing in list)
            {
                if (string.IsNullOrEmpty(existing)) continue;
                string[] parts = existing.Split('ç');
                if (parts.Length >= 2 &&
                    (string.Equals(parts[0], label, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(parts[1], content, StringComparison.Ordinal)))
                {
                    ShowStatus("Already in your clipboard menu - moving on.", true);
                    if (_ownerForm != null) _ownerForm.RefreshClipboardMenuUI();
                    return true;
                }
            }

            list.Add(entry);
            Properties.Settings.Default.settingClipboardList = list;
            Properties.Settings.Default.Save();
            if (_ownerForm != null) _ownerForm.RefreshClipboardMenuUI();
            ShowStatus("Added \"" + label + "\" to your clipboard menu.", true);
            return true;
        }

        // -----------------------------------------------------------------
        // Title bar behavior.
        // -----------------------------------------------------------------

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void CloseLabel_MouseEnter(object sender, EventArgs e)
        {
            _closeLabel.BackColor = ThemeManager.CloseHoverColor;
            _closeLabel.ForeColor = Color.White;
        }

        private void CloseLabel_MouseLeave(object sender, EventArgs e)
        {
            ThemeManager.Palette p = ThemeManager.Current;
            _closeLabel.BackColor = p.TitleBar;
            _closeLabel.ForeColor = p.TitleText;
        }

        private void TutorialForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Skip_Click(this, EventArgs.Empty);
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                Next_Click(this, EventArgs.Empty);
                e.SuppressKeyPress = true;
            }
        }
    }
}
