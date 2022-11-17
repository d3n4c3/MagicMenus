using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MagicMenus
{
    public partial class SetHotKey : Form
    {
        public string hkModkeys { get; set; }
        public Keys hkKeypress { get; set; }

        public bool settingHotkey = false;
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

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

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        public SetHotKey()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 5, 5));
        }

        private void SetHotKey_Load(object sender, EventArgs e)
        {
            settingHotkey = true;

        }

        private void SetHotKey_KeyDown(object sender, KeyEventArgs e)
        {
            if (settingHotkey)
            {

                    if (e.Modifiers.ToString().Contains(","))
                    {
                    string[] countingArray = e.Modifiers.ToString().Split(',');
                    if (countingArray.Length == 2)
                    {
                        txtModifer.Text = e.Modifiers.ToString();
                        if (e.KeyCode.ToString().Length < 3)
                        {
                            txtKey.Text = e.KeyCode.ToString();
                            settingHotkey = false;
                            var response = MessageBox.Show("Would you like to set the hotkey for this menu to: " + txtModifer.Text + " + " + txtKey.Text + "?" + "\n\nThe application will restart.", "Set Hotkey?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (response.ToString() == "Yes")
                            {
                                hkModkeys = e.Modifiers.ToString();
                                hkKeypress = e.KeyCode;
                                settingHotkey = false;
                                this.Close();
                            }
                            else
                            {
                                txtModifer.Text = string.Empty;
                                txtKey.Text = string.Empty;
                                settingHotkey = false;
                            }
                        }
                    }
                }
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnModifier_Click(object sender, EventArgs e)
        {
            txtModifer.Text = string.Empty;
            txtKey.Text = string.Empty;
            settingHotkey = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            label1.BackColor = Color.FromArgb(64, 64, 64);
            hkModkeys = "";
            hkKeypress = Keys.None;
            this.Close();

        }
        private void label1_MouseEnter(object sender, EventArgs e)
        {
            label1.BackColor = Color.Red;
        }
        private void label1_MouseLeave(object sender, EventArgs e)
        {
            label1.BackColor = Color.FromArgb(64, 64, 64);
        }
        private void DragAndDrop(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        private void label3_MouseDown(object sender, MouseEventArgs e)
        {
           DragAndDrop(e);
        }
    }
}
