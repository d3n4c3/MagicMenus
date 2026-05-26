using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MagicMenus
{
    /// <summary>
    /// Centralized theming for the application. Reads the persisted dark-mode
    /// flag from <see cref="MagicMenus.Properties.Settings"/> and applies a
    /// modern flat palette to any form/control tree it is given.
    /// </summary>
    public static class ThemeManager
    {
        public class Palette
        {
            public Color FormBack;
            public Color Surface;
            public Color TitleBar;
            public Color TitleText;
            public Color Text;
            public Color SubtleText;
            public Color Accent;
            public Color AccentHover;
            public Color AccentPressed;
            public Color ButtonBack;
            public Color ButtonHover;
            public Color ButtonBorder;
            public Color InputBack;
            public Color InputBorder;
            public Color ListBack;
            public Color ListSelection;
            public Color GroupBoxFore;
            public Color GroupBoxBorder;
            public Color TabBack;
            public Color TabActive;
            public Color TabInactive;
            public Color TabHover;
            public Color TabBorder;
            public Color CloseHover;
            public Color MenuBack;
            public Color MenuText;
            public Color MenuHover;
            public Color MenuSeparator;
        }

        // ------------------------------------------------------------
        // Light palette: cool neutral whites with a Fluent / Win11 vibe.
        // ------------------------------------------------------------
        public static readonly Palette Light = new Palette
        {
            FormBack = Color.FromArgb(248, 249, 251),
            Surface = Color.White,
            TitleBar = Color.FromArgb(248, 249, 251),
            TitleText = Color.FromArgb(24, 26, 30),
            Text = Color.FromArgb(24, 26, 30),
            SubtleText = Color.FromArgb(108, 114, 124),
            Accent = Color.FromArgb(0, 122, 204),
            AccentHover = Color.FromArgb(20, 142, 224),
            AccentPressed = Color.FromArgb(0, 96, 170),
            ButtonBack = Color.White,
            ButtonHover = Color.FromArgb(240, 244, 250),
            ButtonBorder = Color.FromArgb(214, 218, 224),
            InputBack = Color.White,
            InputBorder = Color.FromArgb(214, 218, 224),
            ListBack = Color.White,
            ListSelection = Color.FromArgb(220, 235, 250),
            GroupBoxFore = Color.FromArgb(72, 78, 88),
            GroupBoxBorder = Color.FromArgb(228, 231, 236),
            TabBack = Color.FromArgb(248, 249, 251),
            TabActive = Color.White,
            TabInactive = Color.FromArgb(238, 240, 244),
            TabHover = Color.FromArgb(230, 234, 240),
            TabBorder = Color.FromArgb(228, 231, 236),
            CloseHover = Color.FromArgb(232, 17, 35),
            MenuBack = Color.White,
            MenuText = Color.FromArgb(24, 26, 30),
            MenuHover = Color.FromArgb(220, 235, 250),
            MenuSeparator = Color.FromArgb(228, 231, 236)
        };

        // ------------------------------------------------------------
        // Dark palette: cool, slightly-blue neutrals inspired by VS Code
        // Dark+. Single base tone for the chrome (form, title bar, tab
        // strip) so the window reads as one surface; controls sit on a
        // slightly lighter "surface" tone, active accents are vivid blue.
        // ------------------------------------------------------------
        public static readonly Palette Dark = new Palette
        {
            FormBack = Color.FromArgb(24, 26, 31),
            Surface = Color.FromArgb(34, 37, 43),
            TitleBar = Color.FromArgb(24, 26, 31),
            TitleText = Color.FromArgb(240, 242, 247),
            Text = Color.FromArgb(228, 231, 238),
            SubtleText = Color.FromArgb(160, 166, 178),
            Accent = Color.FromArgb(94, 158, 255),
            AccentHover = Color.FromArgb(130, 180, 255),
            AccentPressed = Color.FromArgb(72, 138, 232),
            ButtonBack = Color.FromArgb(40, 43, 50),
            ButtonHover = Color.FromArgb(54, 58, 68),
            ButtonBorder = Color.FromArgb(60, 65, 76),
            InputBack = Color.FromArgb(32, 35, 41),
            InputBorder = Color.FromArgb(58, 63, 73),
            ListBack = Color.FromArgb(32, 35, 41),
            ListSelection = Color.FromArgb(45, 102, 170),
            GroupBoxFore = Color.FromArgb(214, 218, 226),
            GroupBoxBorder = Color.FromArgb(46, 50, 58),
            TabBack = Color.FromArgb(24, 26, 31),
            TabActive = Color.FromArgb(34, 37, 43),
            TabInactive = Color.FromArgb(24, 26, 31),
            TabHover = Color.FromArgb(40, 43, 50),
            TabBorder = Color.FromArgb(46, 50, 58),
            CloseHover = Color.FromArgb(232, 17, 35),
            MenuBack = Color.FromArgb(34, 37, 43),
            MenuText = Color.FromArgb(228, 231, 238),
            MenuHover = Color.FromArgb(45, 102, 170),
            MenuSeparator = Color.FromArgb(60, 65, 76)
        };

        // Typography hierarchy.
        //   UiFont       — default body text (8pt, sized to match the original
        //                  Calibri 8.25pt layout so we don't clip captions).
        //   UiFontBold   — Semibold 8pt for emphasis (button captions on the
        //                  ↑ / ↓ / X buttons, GroupBox titles, etc.).
        //   HeaderFont   — Semibold 9pt for section headers that we render
        //                  ourselves above grouped content.
        //   CaptionFont  — 7.75pt for subtle helper/caption text.
        //   TitleFont    — 11pt Semibold for the window title bar.
        //   CloseFont    — 10pt for the title-bar close glyph.
        public static readonly Font UiFont = new Font("Segoe UI", 8F, FontStyle.Regular);
        public static readonly Font UiFontBold = new Font("Segoe UI Semibold", 8F, FontStyle.Regular);
        public static readonly Font HeaderFont = new Font("Segoe UI Semibold", 9F, FontStyle.Regular);
        public static readonly Font CaptionFont = new Font("Segoe UI", 7.75F, FontStyle.Regular);
        public static readonly Font TitleFont = new Font("Segoe UI Semibold", 11F, FontStyle.Regular);
        public static readonly Font CloseFont = new Font("Segoe UI", 10F, FontStyle.Regular);

        // Tag string conventions on Controls so the theme manager can give
        // them a distinct role:
        //   "primary"    — Button rendered with accent fill
        //   "header"     — Label rendered with HeaderFont
        //   "subtle"     — Label rendered with SubtleText color (and CaptionFont)
        //   "code"       — Label rendered with monospace text (for wildcards)
        //   "icon"       — Small square Button used as an icon affordance
        //   "titlebar"   — Container drawn as the title bar
        public const string TagPrimary = "primary";
        public const string TagHeader = "header";
        public const string TagSubtle = "subtle";
        public const string TagCode = "code";
        public const string TagIcon = "icon";
        public const string TagTitleBar = "titlebar";

        // Buttons known to be primary actions (filled accent). Matched by
        // Name so we can keep the call site declarative. Names below come
        // straight from MagicMenus.Designer.cs.
        private static readonly HashSet<string> _primaryButtonNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "btnWCString",          // Add - string wildcard
            "btnSetActionHotkey",   // SET - action hotkey
            "btnSetClipboardHotkey",// SET - clipboard hotkey
            "btnActionShell",       // Save to Menu - Action / Shell tab
            "btnActionURL",         // Save to Menu - Action / Link tab
            "btnClipboardSave",     // Save to Menu - Clipboard / Rich Text File tab
            "button3",              // Save to Menu - Clipboard / String tab (default designer name)
            "btnShowActionMenu",    // Show Menu - Action
            "btnShowClipboardMenu"  // Show Menu - Clipboard
        };

        public static bool IsDark
        {
            get
            {
                try { return MagicMenus.Properties.Settings.Default.darkMode; }
                catch { return false; }
            }
            set
            {
                MagicMenus.Properties.Settings.Default.darkMode = value;
                MagicMenus.Properties.Settings.Default.Save();
            }
        }

        public static Palette Current
        {
            get { return IsDark ? Dark : Light; }
        }

        public static Color TitleBarBackground { get { return Current.TitleBar; } }
        public static Color TitleBarForeground { get { return Current.TitleText; } }
        public static Color CloseHoverColor { get { return Current.CloseHover; } }

        /// <summary>Apply the active theme to a form and all of its controls.</summary>
        /// <summary>
        /// Apply the current palette to a single control that was added to
        /// a form after Apply(Form) ran. Used by the tutorial wizard, which
        /// builds controls per-step and needs them to pick up the dark/light
        /// theme without re-walking the entire form.
        /// </summary>
        public static void RestyleControl(Control c)
        {
            if (c == null) return;
            ApplyToControl(c, Current);
        }

        public static void Apply(Form form)
        {
            if (form == null) return;
            Palette p = Current;

            form.BackColor = p.FormBack;
            form.ForeColor = p.Text;
            // NOTE: we deliberately do not touch form.Font. Doing so triggers
            // WinForms' Font-based AutoScale, which quietly resizes every
            // child control and can clip button captions like "Save to Menu".
            // Each control type below sets its own font explicitly instead.

            ApplyToControls(form.Controls, p);
            PolishLayout(form, p);
            form.Invalidate(true);
        }

        /// <summary>
        /// Apply per-form layout fixes that the designer can't express well:
        /// re-flow the User Input / String / Settings wildcard description
        /// labels so they don't overlap with Segoe UI metrics, restyle the
        /// description text as subtle muted body copy, and embed the
        /// wildcard literal (%INPUT% / %SETTINGS% / etc.) as code-styled
        /// inline text rather than wrapping awkwardly.
        /// </summary>
        private static void PolishLayout(Form form, Palette p)
        {
            // Hide the outer container GroupBox titles - they duplicate the
            // tab labels ("General" / "Action Menu" / "Clipboard Menu"). The
            // GroupBox still acts as a structural container but reads as a
            // clean dark panel instead of a labelled frame.
            HideGroupBoxChrome(form, "gbGeneral");
            HideGroupBoxChrome(form, "gbActionMenu");
            HideGroupBoxChrome(form, "gbCopyMenu");

            // The three description boxes share the same shape:
            //   - intro label
            //   - numbered point 1
            //   - numbered point 2
            // Replace them in-place with cleaner, shorter copy and re-flow
            // them so they actually fit at Segoe UI 8pt.
            PolishWildcardDescription(form, "gbInput",
                "Prompt the user for input when a menu item is invoked. The value typed in replaces the placeholder in the command.",
                "Example: passing the prompt's result into a search engine URL, a script argument, or a shell command.",
                "%INPUT%");

            PolishWildcardDescription(form, "gbString",
                "Define a reusable string here and reference it from menus to keep values like an API endpoint or username in one place.",
                "Editing the value updates every menu item that references the wildcard - no more hunting through commands.",
                null);

            PolishWildcardDescription(form, "gbSettings",
                "Bind this Settings dialog to a menu item so you can re-open it from anywhere in the app.",
                "Useful when the menu hotkey is your primary launcher and you'd rather not minimize to the tray icon.",
                "%SETTINGS%");
        }

        /// <summary>
        /// Suppress the etched-frame look of a GroupBox: clear its caption
        /// (so we don't paint a header) and tag it so the overpaint hook
        /// skips drawing the inset border too. The container still hosts its
        /// children at the same positions, it just visually dissolves into
        /// the form.
        /// </summary>
        private static void HideGroupBoxChrome(Form form, string groupBoxName)
        {
            GroupBox gb = FindControl(form, groupBoxName) as GroupBox;
            if (gb == null) return;
            gb.Text = string.Empty;
            gb.Tag = "noborder";
            gb.Invalidate();
        }

        /// <summary>
        /// Re-populate the String Wildcard description GroupBox to show a
        /// specific user-defined wildcard's token and current value, with a
        /// helpful hint reminding the reader they can press Delete in the
        /// list to remove it.
        /// </summary>
        public static void ShowCustomStringWildcard(GroupBox gb, string token, string value)
        {
            if (gb == null) return;
            string displayValue = string.IsNullOrEmpty(value) ? "(empty)" : value;
            // Truncate very long values so the description doesn't run off
            // the bottom of the box.
            const int maxValueChars = 140;
            if (displayValue.Length > maxValueChars)
                displayValue = displayValue.Substring(0, maxValueChars) + "...";

            PolishDescriptionGroup(gb,
                "Reference " + token + " in any menu item to have it expanded to the value below at runtime.",
                "Current value: " + displayValue,
                token,
                "Right-click in the list (or press Delete) to remove this wildcard.");
        }

        private static void PolishWildcardDescription(Form form, string groupBoxName,
            string intro, string detail, string codeBadge)
        {
            GroupBox gb = FindControl(form, groupBoxName) as GroupBox;
            if (gb == null) return;
            PolishDescriptionGroup(gb, intro, detail, codeBadge, null);
        }

        /// <summary>
        /// Re-layout a GroupBox with up to four pieces of content stacked
        /// vertically: an intro body label, a detail body label, a code-styled
        /// wildcard badge, and an optional subtle hint footer. Any "extra"
        /// labels beyond the four we use are hidden.
        /// </summary>
        private static void PolishDescriptionGroup(GroupBox gb,
            string intro, string detail, string codeBadge, string footerHint)
        {
            if (gb == null) return;

            // Gather child Labels in vertical order.
            List<Label> labels = new List<Label>();
            foreach (Control child in gb.Controls)
            {
                Label l = child as Label;
                if (l != null) labels.Add(l);
            }
            labels.Sort(delegate(Label a, Label b) { return a.Top.CompareTo(b.Top); });
            if (labels.Count < 2) return;

            Label introLabel = labels[0];
            Label detailLabel = labels[1];
            Label codeLabel = labels.Count >= 3 ? labels[2] : null;

            const int leftPad = 10;
            const int rightPad = 10;
            int innerWidth = gb.ClientSize.Width - leftPad - rightPad;
            int top = 18;

            ConfigureBodyLabel(introLabel, intro, leftPad, top, innerWidth);
            top = introLabel.Bottom + 4;

            ConfigureBodyLabel(detailLabel, detail, leftPad, top, innerWidth);
            top = detailLabel.Bottom + 8;

            if (codeLabel != null)
            {
                if (!string.IsNullOrEmpty(codeBadge))
                {
                    codeLabel.Visible = true;
                    codeLabel.Tag = TagCode;
                    codeLabel.Text = codeBadge;
                    codeLabel.AutoSize = false;
                    codeLabel.MaximumSize = Size.Empty;
                    codeLabel.Font = new Font("Consolas", 9F, FontStyle.Regular);
                    codeLabel.ForeColor = Current.Accent;
                    Size codeSize = TextRenderer.MeasureText(codeBadge, codeLabel.Font,
                        new Size(innerWidth, int.MaxValue), TextFormatFlags.NoPadding);
                    codeLabel.SetBounds(leftPad, top, codeSize.Width + 4, codeSize.Height + 2);
                    codeLabel.TextAlign = ContentAlignment.MiddleLeft;
                    top = codeLabel.Bottom + 6;
                }
                else
                {
                    codeLabel.Text = string.Empty;
                    codeLabel.Visible = false;
                }
            }

            // Footer hint: small subtle helper text appended below everything
            // (used to remind the user how to delete a custom wildcard).
            // We re-use the first "extra" hidden label if one exists, or
            // synthesize one and parent it under the GroupBox.
            const string footerName = "_polishFooter";
            Label footerLabel = FindNamedChild(gb, footerName);
            if (!string.IsNullOrEmpty(footerHint))
            {
                if (footerLabel == null)
                {
                    footerLabel = new Label();
                    footerLabel.Name = footerName;
                    gb.Controls.Add(footerLabel);
                }
                footerLabel.Visible = true;
                ConfigureBodyLabel(footerLabel, footerHint, leftPad, top, innerWidth);
                top = footerLabel.Bottom + 4;
            }
            else if (footerLabel != null)
            {
                footerLabel.Visible = false;
            }

            // Make sure the GroupBox is tall enough to contain all of it.
            int extraSpace = 8;
            int neededHeight = top + 6 + extraSpace;
            if (gb.Height < neededHeight)
                gb.Height = neededHeight;
        }

        private static void ConfigureBodyLabel(Label lbl, string text, int left, int top, int maxWidth)
        {
            lbl.Tag = TagSubtle;
            lbl.AutoSize = false;
            lbl.AutoEllipsis = false;
            lbl.Font = CaptionFont;
            lbl.ForeColor = Current.SubtleText;
            lbl.BackColor = Current.FormBack;
            lbl.Text = text;
            lbl.MaximumSize = new Size(maxWidth, 0);
            // Measure with word wrap to find the natural height.
            Size measured = TextRenderer.MeasureText(text, lbl.Font,
                new Size(maxWidth, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            lbl.SetBounds(left, top, maxWidth, measured.Height + 2);
            lbl.TextAlign = ContentAlignment.TopLeft;
        }

        private static Label FindNamedChild(Control parent, string name)
        {
            if (parent == null) return null;
            foreach (Control child in parent.Controls)
            {
                if (child is Label && string.Equals(child.Name, name, StringComparison.Ordinal))
                    return (Label)child;
            }
            return null;
        }

        private static Control FindControl(Control root, string name)
        {
            if (root == null) return null;
            if (string.Equals(root.Name, name, StringComparison.Ordinal))
                return root;
            foreach (Control child in root.Controls)
            {
                Control found = FindControl(child, name);
                if (found != null) return found;
            }
            return null;
        }

        /// <summary>Apply title-bar styling for the custom borderless windows.</summary>
        public static void ApplyTitleBar(Control titleBar, Label titleLabel, Label closeLabel)
        {
            Palette p = Current;
            if (titleBar != null)
            {
                titleBar.BackColor = p.TitleBar;
            }
            if (titleLabel != null)
            {
                titleLabel.BackColor = p.TitleBar;
                titleLabel.ForeColor = p.TitleText;
                titleLabel.Font = TitleFont;
            }
            if (closeLabel != null)
            {
                closeLabel.BackColor = p.TitleBar;
                closeLabel.ForeColor = p.TitleText;
                closeLabel.Font = CloseFont;
            }
        }

        /// <summary>Style a ContextMenuStrip (used for popup action/clipboard menus).</summary>
        public static void ApplyMenu(ToolStrip strip)
        {
            if (strip == null) return;
            Palette p = Current;
            strip.Renderer = new FlatMenuRenderer(p);
            strip.BackColor = p.MenuBack;
            strip.ForeColor = p.MenuText;
            strip.Font = UiFont;
            foreach (ToolStripItem item in strip.Items)
                StyleToolStripItem(item, p);
        }

        private static void StyleToolStripItem(ToolStripItem item, Palette p)
        {
            item.BackColor = p.MenuBack;
            item.ForeColor = p.MenuText;
            item.Font = UiFont;
            ToolStripDropDownItem dropDown = item as ToolStripDropDownItem;
            if (dropDown != null && dropDown.HasDropDownItems)
            {
                foreach (ToolStripItem child in dropDown.DropDownItems)
                    StyleToolStripItem(child, p);
            }
        }

        private static void ApplyToControls(Control.ControlCollection controls, Palette p)
        {
            foreach (Control c in controls)
                ApplyToControl(c, p);
        }

        private static void ApplyToControl(Control c, Palette p)
        {
            if (c == null) return;

            if (c.Tag != null && string.Equals(c.Tag.ToString(), "titlebar", StringComparison.OrdinalIgnoreCase))
            {
                c.BackColor = p.TitleBar;
                c.ForeColor = p.TitleText;
                if (c.HasChildren) ApplyToControls(c.Controls, p);
                return;
            }

            if (c is TabControl)
            {
                TabControl tc = (TabControl)c;
                tc.Font = UiFont;
                tc.SizeMode = TabSizeMode.Normal;
                tc.Appearance = TabAppearance.Normal;
                // We paint the whole control ourselves below; using OwnerDrawFixed
                // would only owner-draw the tab buttons and would still let the
                // system paint a light "frame" around the active TabPage.
                tc.DrawMode = TabDrawMode.Normal;
                tc.BackColor = p.FormBack;
                // Force OptimizedDoubleBuffer on the TabControl. The system
                // ignores the public Control.DoubleBuffered property here
                // because the tab strip is drawn by comctl32, but flipping
                // the style bits directly works and is what eliminates the
                // hover flicker on the tab buttons.
                EnableDoubleBuffer(tc);
                // Same problem as GroupBox: Paint event isn't reliably the last
                // word under visual styles. Subclass WM_PAINT so we always
                // overpaint after the system has finished drawing.
                AttachTabControlOverpaintHook(tc);
            }
            else if (c is TabPage)
            {
                TabPage tp = (TabPage)c;
                tp.UseVisualStyleBackColor = false;
                tp.BackColor = p.FormBack;
                tp.ForeColor = p.Text;
                tp.Font = UiFont;
            }
            else if (c is GroupBox)
            {
                GroupBox gb = (GroupBox)c;
                gb.BackColor = p.FormBack;
                gb.ForeColor = p.GroupBoxFore;
                gb.Font = UiFontBold;
                // Subscribing to .Paint isn't enough: GroupBox draws its
                // etched border via OnPaint *and* lets Windows redraw the
                // border again as part of WM_PAINT post-processing under
                // visual styles, so any fill we do in the Paint event gets
                // covered up. Hook WM_PAINT directly and overpaint after.
                AttachGroupBoxOverpaintHook(gb);
            }
            else if (c is Button)
            {
                Button btn = (Button)c;
                bool isPrimary = IsPrimaryButton(btn);

                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = isPrimary ? 0 : 1;
                btn.FlatAppearance.BorderColor = isPrimary ? p.Accent : p.ButtonBorder;
                btn.FlatAppearance.MouseOverBackColor = isPrimary ? p.AccentHover : p.ButtonHover;
                btn.FlatAppearance.MouseDownBackColor = p.AccentPressed;
                btn.BackColor = isPrimary ? p.Accent : p.ButtonBack;
                btn.ForeColor = isPrimary ? Color.White : p.Text;
                btn.UseVisualStyleBackColor = false;
                // Preserve the original font weight - the tiny arrow / "X"
                // buttons were designed as Bold so the single-glyph caption
                // stays legible.
                bool wasBold = btn.Font != null && btn.Font.Bold;
                btn.Font = wasBold ? UiFontBold : UiFont;
                btn.AutoSize = false;
                btn.AutoEllipsis = true;
                btn.Cursor = Cursors.Hand;
                btn.TextAlign = ContentAlignment.MiddleCenter;
                btn.Padding = new Padding(0);
                // The system's button text rendering uses TextFormatFlags
                // with WordBreak, which was wrapping "Save to Menu" to two
                // lines that the 21-px button height then clipped. We
                // over-paint the caption with SingleLine | EndEllipsis so
                // it's guaranteed to stay on one line no matter what.
                btn.Paint -= Button_Paint;
                btn.Paint += Button_Paint;
            }
            else if (c is TextBox)
            {
                TextBox tb = (TextBox)c;
                tb.BorderStyle = BorderStyle.FixedSingle;
                tb.BackColor = tb.Enabled ? p.InputBack : p.FormBack;
                tb.ForeColor = p.Text;
                tb.Font = UiFont;
                // Subtle focus accent: tint the BackColor a touch on focus
                // so users can see which input is active without a heavy
                // outline. Detach old handlers first to be idempotent
                // across re-applies (toggling light/dark).
                tb.GotFocus -= TextBox_GotFocus;
                tb.LostFocus -= TextBox_LostFocus;
                tb.GotFocus += TextBox_GotFocus;
                tb.LostFocus += TextBox_LostFocus;
            }
            else if (c is ListBox)
            {
                ListBox lb = (ListBox)c;
                lb.BorderStyle = BorderStyle.FixedSingle;
                lb.BackColor = p.ListBack;
                lb.ForeColor = p.Text;
                lb.Font = UiFont;
                lb.IntegralHeight = false;
            }
            else if (c is CheckBox)
            {
                CheckBox cb = (CheckBox)c;
                cb.BackColor = p.FormBack;
                cb.ForeColor = p.Text;
                cb.Font = UiFont;
                cb.FlatStyle = FlatStyle.Flat;
                cb.FlatAppearance.BorderColor = p.ButtonBorder;
                cb.FlatAppearance.CheckedBackColor = p.Accent;
            }
            else if (c is LinkLabel)
            {
                LinkLabel ll = (LinkLabel)c;
                ll.BackColor = p.FormBack;
                ll.LinkColor = p.Accent;
                ll.ActiveLinkColor = p.AccentHover;
                ll.VisitedLinkColor = p.Accent;
                ll.LinkBehavior = LinkBehavior.HoverUnderline;
                ll.Font = UiFont;
                // The "Set wildcard variables..." label was bleeding past
                // its column. Pinning MaximumSize caused wrapping. Instead:
                // measure the text and shift the label LEFT just enough
                // that it fits as a single line inside the parent container.
                ll.MaximumSize = Size.Empty;
                if (ll.AutoSize && ll.Parent != null && !string.IsNullOrEmpty(ll.Text))
                {
                    Size textSize = TextRenderer.MeasureText(ll.Text, ll.Font);
                    int parentWidth = ll.Parent.ClientSize.Width;
                    const int rightPadding = 8;
                    int desiredLeft = parentWidth - textSize.Width - rightPadding;
                    if (desiredLeft < ll.Left)
                    {
                        if (desiredLeft < 6) desiredLeft = 6;
                        ll.Left = desiredLeft;
                    }
                }
            }
            else if (c is Label)
            {
                Label lbl = (Label)c;
                lbl.BackColor = p.FormBack;
                string role = lbl.Tag as string;
                if (string.Equals(role, TagHeader, StringComparison.OrdinalIgnoreCase))
                {
                    lbl.ForeColor = p.Text;
                    lbl.Font = HeaderFont;
                }
                else if (string.Equals(role, TagSubtle, StringComparison.OrdinalIgnoreCase))
                {
                    lbl.ForeColor = p.SubtleText;
                    lbl.Font = CaptionFont;
                }
                else if (string.Equals(role, TagCode, StringComparison.OrdinalIgnoreCase))
                {
                    lbl.ForeColor = p.Accent;
                    lbl.Font = new Font("Consolas", 8.5F, FontStyle.Regular);
                }
                else
                {
                    lbl.ForeColor = p.Text;
                    if (lbl.Font == null || lbl.Font.Name == "Calibri")
                        lbl.Font = UiFont;
                }
            }
            else if (c is PictureBox)
            {
                // leave picture boxes alone unless explicitly tagged.
            }
            else
            {
                c.BackColor = p.FormBack;
                c.ForeColor = p.Text;
                if (c.Font == null || c.Font.Name == "Calibri")
                    c.Font = UiFont;
            }

            if (c.HasChildren)
                ApplyToControls(c.Controls, p);
        }

        // -----------------------------------------------------------------
        // Owner-draw helpers so the theme reads consistently in dark mode.
        // -----------------------------------------------------------------

        private static bool IsPrimaryButton(Button btn)
        {
            if (btn == null) return false;
            if (btn.Tag is string)
            {
                string tag = (string)btn.Tag;
                if (string.Equals(tag, TagPrimary, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return !string.IsNullOrEmpty(btn.Name) && _primaryButtonNames.Contains(btn.Name);
        }

        private static void TextBox_GotFocus(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb == null || !tb.Enabled) return;
            // Lift the input one tone on focus - a clean, unobtrusive cue.
            Palette p = Current;
            tb.BackColor = p.Surface;
        }

        private static void TextBox_LostFocus(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb == null) return;
            Palette p = Current;
            tb.BackColor = tb.Enabled ? p.InputBack : p.FormBack;
        }

        /// <summary>
        /// Re-render the button caption on a single line. The system has
        /// already painted the background, border, and (possibly wrapped)
        /// text by the time this fires; we figure out the current hover /
        /// pressed background, fill the content rect, and draw the caption
        /// ourselves with SingleLine | EndEllipsis.
        /// </summary>
        private static void Button_Paint(object sender, PaintEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;
            if (string.IsNullOrEmpty(btn.Text)) return;

            Palette p = Current;
            Graphics g = e.Graphics;
            bool isPrimary = IsPrimaryButton(btn);

            Point cursorPos = btn.PointToClient(Cursor.Position);
            bool isHover = btn.Enabled && btn.ClientRectangle.Contains(cursorPos);
            bool isPressed = isHover &&
                (Control.MouseButtons & MouseButtons.Left) == MouseButtons.Left;

            Color bgColor;
            if (!btn.Enabled)
            {
                // Disabled primary buttons fall back to the muted ButtonBack
                // - a filled accent button that doesn't react to clicks is
                // confusing UX.
                bgColor = p.ButtonBack;
            }
            else if (isPrimary)
            {
                if (isPressed) bgColor = p.AccentPressed;
                else if (isHover) bgColor = p.AccentHover;
                else bgColor = p.Accent;
            }
            else
            {
                if (isPressed) bgColor = p.AccentPressed;
                else if (isHover) bgColor = p.ButtonHover;
                else bgColor = btn.BackColor;
            }

            // For primary buttons we own the whole client area (no border);
            // for secondary, stay one pixel inside the FlatStyle border.
            Rectangle contentRect = btn.ClientRectangle;
            if (!isPrimary) contentRect.Inflate(-1, -1);

            using (SolidBrush bg = new SolidBrush(bgColor))
                g.FillRectangle(bg, contentRect);

            Color foreColor;
            if (!btn.Enabled) foreColor = p.SubtleText;
            else if (isPrimary) foreColor = Color.White;
            else if (isPressed) foreColor = Color.White;
            else foreColor = btn.ForeColor;

            TextFormatFlags flags = TextFormatFlags.SingleLine |
                                    TextFormatFlags.HorizontalCenter |
                                    TextFormatFlags.VerticalCenter |
                                    TextFormatFlags.EndEllipsis |
                                    TextFormatFlags.NoPadding;

            TextRenderer.DrawText(g, btn.Text, btn.Font, contentRect, foreColor, flags);
        }

        // -----------------------------------------------------------------
        // WndProc subclass overpaint hooks.
        // Subscribing to .Paint isn't enough under visual styles: the system
        // redraws the GroupBox border (and the frame around the active
        // TabPage) AFTER the Paint event fires, so any fill we do there is
        // covered up again. The fix is to subclass WM_PAINT: let the system
        // do its painting, then immediately overpaint our themed version.
        // -----------------------------------------------------------------

        private const int WM_PAINT = 0x000F;
        private const int WM_ERASEBKGND = 0x0014;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct PAINTSTRUCT
        {
            public IntPtr hdc;
            public bool fErase;
            public RECT rcPaint;
            public bool fRestore;
            public bool fIncUpdate;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
            public byte[] rgbReserved;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr BeginPaint(IntPtr hWnd, out PAINTSTRUCT lpPaint);

        [DllImport("user32.dll")]
        private static extern bool EndPaint(IntPtr hWnd, [In] ref PAINTSTRUCT lpPaint);

        private static readonly Dictionary<IntPtr, OverpaintHook> _hooks =
            new Dictionary<IntPtr, OverpaintHook>();

        private static void AttachGroupBoxOverpaintHook(GroupBox gb)
        {
            if (!gb.IsHandleCreated)
            {
                gb.HandleCreated += delegate { AttachGroupBoxOverpaintHook(gb); };
                return;
            }
            AttachOverpaintHook(gb, PaintGroupBox);
        }

        private static void AttachTabControlOverpaintHook(TabControl tc)
        {
            if (!tc.IsHandleCreated)
            {
                tc.HandleCreated += delegate { AttachTabControlOverpaintHook(tc); };
                return;
            }

            // Disable Windows visual styles for this TabControl. With visual
            // styles ON, comctl32 paints an etched "frame" around the active
            // tab body and a subtle 3D rim around the active tab button that
            // bleeds through any overpaint we apply. Setting the theme to
            // "" / "" forces the classic renderer, which is a completely
            // blank surface our overpaint then completely owns.
            try { SetWindowTheme(tc.Handle, " ", " "); } catch { }

            if (AttachOverpaintHook(tc, PaintTabControl))
            {
                // First-time attach: also redraw on selection change so the
                // active-tab indicator and inactive tab colors update.
                tc.SelectedIndexChanged += delegate { tc.Invalidate(); };
                // NOTE: no MouseMove invalidate. With WM_ERASEBKGND
                // suppressed in OverpaintHook and double-buffering enabled,
                // the only paints that happen are the ones initiated by
                // comctl32 itself - and our overpaint runs on the same DC
                // before the screen sees the intermediate state, so there
                // is no white flash on hover.
            }
        }

        /// <summary>
        /// Turn on OptimizedDoubleBuffer + AllPaintingInWmPaint via
        /// reflection. WinForms exposes a public DoubleBuffered property,
        /// but for controls that wrap a native comctl32 window (TabControl,
        /// ListBox, ...) it's effectively a no-op; we have to call the
        /// protected SetStyle directly. We deliberately do NOT set UserPaint
        /// because the TabControl relies on comctl32 to draw the tab strip;
        /// our OverpaintHook then runs on top of that strip on the buffered
        /// DC, so the user only sees the final composited frame.
        /// </summary>
        private static void EnableDoubleBuffer(Control c)
        {
            try
            {
                System.Reflection.MethodInfo setStyle = typeof(Control).GetMethod(
                    "SetStyle",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                System.Reflection.MethodInfo updateStyles = typeof(Control).GetMethod(
                    "UpdateStyles",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (setStyle != null)
                {
                    ControlStyles flags = ControlStyles.OptimizedDoubleBuffer
                                        | ControlStyles.AllPaintingInWmPaint;
                    setStyle.Invoke(c, new object[] { flags, true });
                    if (updateStyles != null) updateStyles.Invoke(c, null);
                }
            }
            catch
            {
                // Style flags are an optimization, not a correctness issue;
                // ignore if reflection is blocked.
            }
        }

        // P/Invoke for SetWindowTheme - used to strip the visual-styles
        // appearance off the TabControl so we get a clean canvas to paint
        // on, without any system-drawn etched frame.
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        /// <summary>
        /// Returns true if a hook was newly attached (false if one was
        /// already present and we just invalidated). Either way the control
        /// is invalidated so the new theme repaints.
        /// </summary>
        private static bool AttachOverpaintHook(Control c, Action<Control, Graphics> paint)
        {
            IntPtr h = c.Handle;
            if (_hooks.ContainsKey(h))
            {
                c.Invalidate();
                return false;
            }
            OverpaintHook hook = new OverpaintHook(c, paint);
            _hooks[h] = hook;
            c.HandleDestroyed += delegate
            {
                _hooks.Remove(h);
                hook.Detach();
            };
            c.Invalidate();
            return true;
        }

        private static void PaintGroupBox(Control control, Graphics g)
        {
            GroupBox gb = control as GroupBox;
            if (gb == null) return;
            Palette p = Current;

            g.SmoothingMode = SmoothingMode.None;
            g.TextRenderingHint = TextRenderingHint.AntiAlias;

            // Wipe everything the system painted, including the etched border.
            using (SolidBrush fill = new SolidBrush(p.FormBack))
                g.FillRectangle(fill, gb.ClientRectangle);

            bool noBorder = string.Equals(gb.Tag as string, "noborder", StringComparison.OrdinalIgnoreCase);
            if (noBorder)
            {
                // Container-only mode: no caption, no border. The control
                // dissolves into the form so we get a single clean surface
                // rather than nested boxes-in-boxes.
                return;
            }

            SizeF textSize = string.IsNullOrEmpty(gb.Text)
                ? SizeF.Empty
                : g.MeasureString(gb.Text, gb.Font);

            int borderTop = (int)Math.Round(textSize.Height / 2f);
            Rectangle borderRect = new Rectangle(
                0,
                borderTop,
                gb.Width - 1,
                gb.Height - borderTop - 1);

            using (Pen pen = new Pen(p.GroupBoxBorder, 1f))
                g.DrawRectangle(pen, borderRect);

            if (!string.IsNullOrEmpty(gb.Text))
            {
                const int textPadding = 8;
                RectangleF textBg = new RectangleF(
                    textPadding - 3,
                    0,
                    textSize.Width + 6,
                    textSize.Height);
                using (SolidBrush bgBrush = new SolidBrush(p.FormBack))
                    g.FillRectangle(bgBrush, textBg);
                using (SolidBrush textBrush = new SolidBrush(p.GroupBoxFore))
                    g.DrawString(gb.Text, gb.Font, textBrush, textPadding, 0);
            }
        }

        private static void PaintTabControl(Control control, Graphics g)
        {
            TabControl tc = control as TabControl;
            if (tc == null) return;
            Palette p = Current;

            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = TextRenderingHint.AntiAlias;

            Rectangle clientRect = tc.ClientRectangle;

            // Wipe everything the system painted - both the tab strip area
            // AND the frame it draws around the active TabPage body. The
            // TabPage child will then paint its own background on top, so
            // we never see any system chrome.
            using (SolidBrush bg = new SolidBrush(p.FormBack))
                g.FillRectangle(bg, clientRect);

            for (int i = 0; i < tc.TabCount; i++)
            {
                Rectangle bounds = tc.GetTabRect(i);

                bool isSelected = (tc.SelectedIndex == i);
                Color back = isSelected ? p.TabActive : p.TabInactive;
                Color fore = isSelected ? p.Text : p.SubtleText;

                using (SolidBrush bgBrush = new SolidBrush(back))
                    g.FillRectangle(bgBrush, bounds);

                if (isSelected)
                {
                    using (SolidBrush accentBrush = new SolidBrush(p.Accent))
                    {
                        // 3-px-tall accent stripe slightly inset from the
                        // tab edges so it reads as a deliberate underline
                        // marker rather than a full-width line.
                        Rectangle indicator = new Rectangle(
                            bounds.X + 4,
                            bounds.Bottom - 3,
                            bounds.Width - 8,
                            3);
                        g.FillRectangle(accentBrush, indicator);
                    }
                }

                string text = tc.TabPages[i].Text ?? string.Empty;
                using (StringFormat sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap
                })
                using (SolidBrush textBrush = new SolidBrush(fore))
                {
                    Font textFont = isSelected ? UiFontBold : tc.Font;
                    g.DrawString(text, textFont, textBrush, bounds, sf);
                }
            }

            // NOTE: no bottom separator. The active-tab accent stripe alone
            // demarcates the strip; a horizontal line below it would
            // re-introduce the "framed" look the user wants gone.
        }

        /// <summary>
        /// NativeWindow subclass that, after WM_PAINT has completed, lets
        /// our PaintAction draw the themed appearance directly on the
        /// control's HDC. This guarantees we are the last paint, no matter
        /// what the system or visual styles draw.
        /// </summary>
        private sealed class OverpaintHook : NativeWindow
        {
            private readonly Control _control;
            private readonly Action<Control, Graphics> _paintAction;
            private bool _detached;

            public OverpaintHook(Control control, Action<Control, Graphics> paintAction)
            {
                _control = control;
                _paintAction = paintAction;
                AssignHandle(control.Handle);
            }

            public void Detach()
            {
                if (_detached) return;
                _detached = true;
                try { ReleaseHandle(); } catch { }
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_ERASEBKGND)
                {
                    // Tell the system we erased without actually erasing.
                    // The default behavior fills the invalid rect with the
                    // class brush (typically COLOR_WINDOW / white), which
                    // produces the bright flash users see on hover and on
                    // tab switch.
                    m.Result = (IntPtr)1;
                    return;
                }

                if (m.Msg == WM_PAINT)
                {
                    // Take full ownership of the paint cycle: skip
                    // base.WndProc so the system never paints its own
                    // (visual-styled or classic) appearance, then draw
                    // our themed look into the buffered DC via
                    // BeginPaint/EndPaint. PaintGroupBox / PaintTabControl
                    // both draw the entire client area from scratch, so
                    // nothing the system would have drawn is needed.
                    if (!_detached && _control != null && _control.IsHandleCreated && _control.Visible)
                    {
                        PAINTSTRUCT ps;
                        IntPtr hdc = BeginPaint(_control.Handle, out ps);
                        try
                        {
                            using (Graphics g = Graphics.FromHdc(hdc))
                            {
                                _paintAction(_control, g);
                            }
                        }
                        catch
                        {
                            // Swallow paint exceptions; missing a frame
                            // is preferable to bringing down the app.
                        }
                        finally
                        {
                            EndPaint(_control.Handle, ref ps);
                        }
                        m.Result = IntPtr.Zero;
                        return;
                    }

                    // Control not in a paintable state - let base handle it
                    // (very rare; mostly during teardown).
                    base.WndProc(ref m);
                    return;
                }

                base.WndProc(ref m);
            }
        }

        /// <summary>Flat renderer used for the popup menus.</summary>
        private class FlatMenuRenderer : ToolStripProfessionalRenderer
        {
            private readonly Palette _p;

            public FlatMenuRenderer(Palette p) : base(new FlatMenuColors(p))
            {
                _p = p;
                RoundedEdges = false;
            }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                e.TextColor = e.Item.Selected ? _p.TitleText : _p.MenuText;
                e.TextFormat |= TextFormatFlags.NoPadding;
                base.OnRenderItemText(e);
            }

            protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
            {
                using (Pen pen = new Pen(_p.MenuSeparator))
                {
                    int y = e.Item.Height / 2;
                    e.Graphics.DrawLine(pen, 4, y, e.Item.Width - 4, y);
                }
            }
        }

        private class FlatMenuColors : ProfessionalColorTable
        {
            private readonly Palette _p;
            public FlatMenuColors(Palette p) { _p = p; UseSystemColors = false; }

            public override Color ToolStripDropDownBackground { get { return _p.MenuBack; } }
            public override Color MenuBorder { get { return _p.ButtonBorder; } }
            public override Color MenuItemBorder { get { return _p.MenuHover; } }
            public override Color MenuItemSelected { get { return _p.MenuHover; } }
            public override Color MenuItemSelectedGradientBegin { get { return _p.MenuHover; } }
            public override Color MenuItemSelectedGradientEnd { get { return _p.MenuHover; } }
            public override Color MenuItemPressedGradientBegin { get { return _p.MenuHover; } }
            public override Color MenuItemPressedGradientEnd { get { return _p.MenuHover; } }
            public override Color ImageMarginGradientBegin { get { return _p.MenuBack; } }
            public override Color ImageMarginGradientMiddle { get { return _p.MenuBack; } }
            public override Color ImageMarginGradientEnd { get { return _p.MenuBack; } }
            public override Color SeparatorDark { get { return _p.MenuSeparator; } }
            public override Color SeparatorLight { get { return _p.MenuSeparator; } }
        }
    }
}
