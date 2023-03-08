using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;

namespace WordDynamicControls
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void ToggleTextBoldOnText()
        {
            Word.Selection selection = this.Application.Selection;

            if (selection != null && selection.Range != null)
            {
                selection.Font.Bold = selection.Font.Bold == 0 ? 1:0;
            }
        }

        internal void CreateTable()
        {
            Word.Range tableLocation =
                this.Application.ActiveDocument.Range(0, 0);
            this.Application.ActiveDocument.Tables.Add(
                tableLocation, 3, 4);
            this.Application.ActiveDocument.Tables[1].Range.Font.Size = 8;
            this.Application.ActiveDocument.Tables[1].set_Style("Table Grid");
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
