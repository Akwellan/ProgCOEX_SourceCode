using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Siemens.Opc.DaClient.Controller
{
    public partial class Aide : Form
    {
        int page = 4;
        int numPage = 1;
        public Aide()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnPage1_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage1.Height;
            pnlBtn.Top = btnPage1.Top;
            pnlBtn.Left = btnPage1.Left;
            btnPage1.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            pictureBox3.Show();
            numPage = 1;
            Page();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void btnPage2_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage2.Height;
            pnlBtn.Top = btnPage2.Top;
            pnlBtn.Left = btnPage2.Left;
            btnPage2.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            pictureBox4.Show();
            numPage = 2;
            Page();
        }

        private void btnPage3_Click(object sender, EventArgs e)
        {

            pnlBtn.Height = btnPage3.Height;
            pnlBtn.Top = btnPage3.Top;
            pnlBtn.Left = btnPage3.Left;
            btnPage3.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            pictureBox5.Show();
            numPage = 3;
            Page();

        }

        private void HidePicture()
        {
            pictureBox3.Hide();
            pictureBox4.Hide();
            pictureBox5.Hide();
            pictureBox6.Hide();
        }

        private void Aide_Load(object sender, EventArgs e)
        {
            HidePicture();
            btnPage1_Click(sender, e);
        }

        private void btnPage4_Click(object sender, EventArgs e)
        {

            pnlBtn.Height = btnPage4.Height;
            pnlBtn.Top = btnPage4.Top;
            pnlBtn.Left = btnPage4.Left;
            btnPage4.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            pictureBox6.Show();
            numPage = 4;
            Page();
        }

        private void btnPage5_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage5.Height;
            pnlBtn.Top = btnPage5.Top;
            pnlBtn.Left = btnPage5.Left;
            btnPage5.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 5;
            Page();

        }

        private void btnPage6_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage6.Height;
            pnlBtn.Top = btnPage6.Top;
            pnlBtn.Left = btnPage6.Left;
            btnPage6.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 6;
            Page();
        }

        private void btnPage7_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage7.Height;
            pnlBtn.Top = btnPage7.Top;
            pnlBtn.Left = btnPage7.Left;
            btnPage7.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 7;
            Page();
        }

        private void btnPage8_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage8.Height;
            pnlBtn.Top = btnPage8.Top;
            pnlBtn.Left = btnPage8.Left;
            btnPage8.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 8;
            Page();
        }

        private void btnPage9_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage9.Height;
            pnlBtn.Top = btnPage9.Top;
            pnlBtn.Left = btnPage9.Left;
            btnPage9.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage10.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 9;
            Page();
        }

        private void btnPage10_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = btnPage10.Height;
            pnlBtn.Top = btnPage10.Top;
            pnlBtn.Left = btnPage10.Left;
            btnPage10.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btnPage1.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage2.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage3.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage4.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage5.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage6.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage7.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage8.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btnPage9.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            HidePicture();
            //pictureBox3.Show();
            numPage = 10;
            Page();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void btnSuivant_Click(object sender, EventArgs e)
        {
            switch (numPage)
            {
                case 1:
                    btnPage2_Click(sender, e);
                    Page();
                    break;
                case 2:
                    btnPage3_Click(sender, e);
                    Page();
                    break;
                case 3:
                    btnPage4_Click(sender, e);
                    Page();
                    break;
                case 4:
                    Page();
                    break;
                default:
                    // code block
                    break;
            }
        }

        private void btnPrecedent_Click(object sender, EventArgs e)
        {
            switch (numPage)
            {
                case 1:
                    Page();
                    break;
                case 2:
                    btnPage1_Click(sender, e);
                    Page();
                    break;
                case 3:
                    btnPage2_Click(sender, e);
                    Page();
                    break;
                case 4:
                    btnPage3_Click(sender, e);
                    Page();
                    break;
                default:
                    // code block
                    break;
            }
        }

        private void Page()
        {
            if(numPage == 1)
            {
                btnPrecedent.Enabled = false;
            } else
            {
                btnPrecedent.Enabled = true;
            }
            if (numPage == page)
            {
                btnSuivant.Enabled = false;
            } else
            {
                btnSuivant.Enabled = true;
            }
        }
    }
}
