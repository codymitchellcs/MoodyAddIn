using Microsoft.Office.Tools.Ribbon;
using MoodysAddIn.UserControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MoodysAddIn
{
    public partial class CustomRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Document Saved Successfully!");
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Hello World!");
        }

        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            EmailLoginForm emailLoginForm = new EmailLoginForm();
            emailLoginForm.ShowDialog();
        }
    }
}
