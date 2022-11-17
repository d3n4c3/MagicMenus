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


            //Action Hook
            hook.KeyPressed +=
                new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);

            if (Settings.Default.settingsActionHotkey == "")
            {
                
                actionSettings = true;

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
                clipboardSettings = true;

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
            label1.BackColor = Color.Red;
            Point p = new Point();
            this.Height = 0;
            this.Width = 0;
            p.X = Screen.PrimaryScreen.WorkingArea.Left;
            p.Y = Screen.PrimaryScreen.WorkingArea.Top;
            this.Location = p;
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
            label1.BackColor = Color.FromArgb(64, 64, 64);
        }

        private void label1_MouseEnter(object sender, EventArgs e)
        {
            label1.BackColor = Color.Red;
        }
        private void label1_MouseLeave(object sender, EventArgs e)
        {
            label1.BackColor = Color.FromArgb(64, 64, 64);
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
            this.Height = 348;
            this.Width = 583;
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
            if (lbWildcards.SelectedIndex == 0)
            {
                gbInput.BringToFront();
            }
            if (lbWildcards.SelectedIndex == 1)
            {
                gbSettings.BringToFront();
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
