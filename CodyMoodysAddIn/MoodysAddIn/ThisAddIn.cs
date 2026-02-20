using System;
using System.IO; 
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using MoodysAddIn.UserControls;

namespace MoodysAddIn
{
    public class AddInUserSettings
    {
        public string AddInUserIdentity { get; set; }
    }
    public partial class ThisAddIn
    {
        public static string CurrentUser; 
        public CustomRibbon cRibbon => Globals.Ribbons.GetRibbon(typeof(CustomRibbon)) as CustomRibbon; 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
