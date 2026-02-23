using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MoodysAddIn
{
    public partial class WebPagePane : UserControl
    {
        public WebPagePane()
        {
            InitializeComponent();
            webBrowser2.Navigate("https://www.google.com");
        }

        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
        }
    }
}
