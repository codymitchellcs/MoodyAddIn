using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MoodysAddIn; 

namespace MoodysAddIn.UserControls
{
    public partial class EmailLoginForm : Form
    {
        public EmailLoginForm()
        {
            InitializeComponent();
        }

        private void EmailLoginForm_Load(object sender, EventArgs e)
        {
            
        }

        private void btnEmailLogin_Click(object sender, EventArgs e)
        {
            if (!isEmailValid(emailTextBox.Text))
            {
                MessageBox.Show("Invalid Email!", "Login Error");
                emailTextBox.Text = string.Empty;
            }
            else if (emailTextBox.Text == "user1@email.com")
            {
                ThisAddIn.CurrentUser = "User1";
                Globals.ThisAddIn.protectDocument();
                Globals.Ribbons.Ribbon1.btnLoad.Enabled = true;
                Globals.Ribbons.Ribbon1.btnSave.Enabled = true;

                this.Close();
            }
            else if (emailTextBox.Text == "user2@email.com")
            {
                ThisAddIn.CurrentUser = "User2";
                Globals.ThisAddIn.protectDocument();
                Globals.Ribbons.Ribbon1.btnLoad.Enabled = false;
                Globals.Ribbons.Ribbon1.btnSave.Enabled = false;

                this.Close(); 
            }
            else
            {
                ThisAddIn.CurrentUser = "OtherUser";
                Globals.ThisAddIn.protectDocument();
                Globals.Ribbons.Ribbon1.btnLoad.Enabled = false;
                Globals.Ribbons.Ribbon1.btnSave.Enabled = false;
                
                this.Close(); 
            }
        }

        private bool isEmailValid(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false;
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }
    }
}
