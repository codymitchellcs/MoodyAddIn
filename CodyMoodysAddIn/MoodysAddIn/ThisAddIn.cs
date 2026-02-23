using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using MoodysAddIn.UserControls;
using System;
using System.Collections.Generic;
using System.IO; 
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

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
        private WebPagePane webPane;
        private Microsoft.Office.Tools.CustomTaskPane customTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            webPane = new WebPagePane();
            customTaskPane = this.CustomTaskPanes.Add(webPane, "Embedded Web Page");
            customTaskPane.Visible = true;
            MessageBox.Show(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            string filePath = @"C:/Users/terri/source/repos/CodyMoodysAddIn/MSWordCandidateEvaluation.docx";
            if (System.IO.File.Exists(filePath))
            {
                Word.Document document = this.Application.Documents.Open(filePath);
                document.Activate();
            }
            CurrentUser = "OtherUser";
            protectDocument();
           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void protectDocument()
        {  if (Globals.ThisAddIn.Application.ActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                this.Application.ActiveDocument.Unprotect();
                this.Application.ActiveDocument.DeleteAllEditableRanges(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone);
            }
                if (CurrentUser == "OtherUser")
            { 
                this.Application.ActiveDocument.Protect(Word.WdProtectionType.wdNoProtection, false, System.String.Empty, false, false);
            }
            else if (CurrentUser == "User1")
            {
                Word.Range userOneEditableRange = this.Application.ActiveDocument.Bookmarks["VersioningPractice"].Range;
                userOneEditableRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                this.Application.ActiveDocument.Protect(Word.WdProtectionType.wdAllowOnlyReading, false, System.String.Empty, false, false);
                        
            }
            else if (CurrentUser == "User2")
            {
                Word.Range userTwoEditableRange = this.Application.ActiveDocument.Bookmarks["LanguageSupport"].Range;
                userTwoEditableRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                this.Application.ActiveDocument.Protect(Word.WdProtectionType.wdAllowOnlyReading, false, System.String.Empty, false, false);
            }
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
