using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Siemens.Opc.DaClient.Controller;

namespace Dashboard
{
    public partial class ProgCOEX : Form
    {
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

        

        public ProgCOEX()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);
            nomPage.Text = ("Dashboard");
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);

            DashControlleur.Hide();
            RecetteControlleur.Hide();
            CorrectionControlleur.Hide();

        }




        private void label1_Click(object sender, EventArgs e)
        {

        }



        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        public void test()
        {
            nomPage.Text = ("DASHBOARD");
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            nomPage.Text = ("DASHBOARD");
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);

            RecetteControlleur.Hide();
            CorrectionControlleur.Hide();

            DashControlleur.Show();
            DashControlleur.BringToFront();

        }

        private void btnCorrection_Click(object sender, EventArgs e)
        {
            nomPage.Text = ("CORRECTION");
            pnlNav.Height = btnCorrection.Height;
            pnlNav.Top = btnCorrection.Top;
            btnCorrection.BackColor = Color.FromArgb(46, 51, 73);
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);


            CorrectionControlleur.Show();
            CorrectionControlleur.BringToFront();

            RecetteControlleur.Hide();

            DashControlleur.Hide();

        }

        private void btnRecettes_Click(object sender, EventArgs e)
        {
            nomPage.Text = ("RECETTES");
            pnlNav.Height = btnRecettes.Height;
            pnlNav.Top = btnRecettes.Top;
            btnRecettes.BackColor = Color.FromArgb(46, 51, 73);
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);


            RecetteControlleur.Show();
            RecetteControlleur.BringToFront();

            DashControlleur.Hide();

            CorrectionControlleur.Hide();

        }

        private void btnParametres_Click(object sender, EventArgs e)
        {
            nomPage.Text = ("PARAMETRES");
            pnlNav.Height = btnParametres.Height;
            pnlNav.Top = btnParametres.Top;
            btnParametres.BackColor = Color.FromArgb(46, 51, 73);
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);

            RecetteControlleur.Hide();

            DashControlleur.Hide();

            CorrectionControlleur.Hide();

        }
        private void btnDashboard_Leave(object sender, EventArgs e)
        {
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);
        }
        private void btnCorrection_Leave(object sender, EventArgs e)
        {
            btnCorrection.BackColor = Color.FromArgb(24, 30, 54);
        }
        private void btnRecettes_Leave(object sender, EventArgs e)
        {
            btnRecettes.BackColor = Color.FromArgb(24, 30, 54);
        }
        private void btnParametres_Leave(object sender, EventArgs e)
        {
            btnParametres.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnDashboard_Click(sender, e);
        }



        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dash1_Load(object sender, EventArgs e)
        {

        }

        private void ProgCOEX_MouseDown(object sender, MouseEventArgs e)
        {
            Show();
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Aide Aide = new Aide();
            Aide.Show();
        }

        private void correction1_Load(object sender, EventArgs e)
        {

        }
    }
}
