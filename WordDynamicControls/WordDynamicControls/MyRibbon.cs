using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordDynamicControls
{
    public partial class MyRibbon
    {

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleTextBoldOnText();
        }
    }
}
