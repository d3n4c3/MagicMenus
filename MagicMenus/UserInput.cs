using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MagicMenus.Settings;

namespace MagicMenus
{

    public partial class UserInput : Form
    {
        public string inputReturn { get { return txtInput.Text; } }

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

        public UserInput()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 5, 5));
        }

        private void UserInput_Load(object sender, EventArgs e)
        {
            this.Width = 383;
            this.Height = 33;
            Point p = new Point();
            p.X = 0;
            p.Y = Screen.PrimaryScreen.WorkingArea.Bottom - this.Height; 
            this.Location = p;
            this.TopMost = true;
            SetForegroundWindow(this.Handle);
        }

        private void btnReturn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DragAndDrop(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        private void label3_MouseDown(object sender, MouseEventArgs e)
        {
            DragAndDrop(e);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            label1.BackColor = Color.FromArgb(64, 64, 64);
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
