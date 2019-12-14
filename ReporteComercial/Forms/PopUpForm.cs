using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FYRASA.Forms
{
    public partial class PopUpForm : Form
    {
        public string Answer;
        public PopUpForm()
        {
            InitializeComponent();
        }

        public PopUpForm(string title, string question, int option)
        {
            InitializeComponent();

            if(option == 0)
            {
                this.lblTitle.Text = title;
                this.lblQuestion.Text = question;
            }
            else if(option == 1)
            {
                this.lblTitle.Text = title;
                this.lblQuestion.Text = question;
                this.txtAnswer.Visible = false;
            }

        }

        private void PopUpForm_Load(object sender, EventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            this.Answer = txtAnswer.Text;
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Answer = "";
            this.Close();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
