using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddInWithButtons
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleBold_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleBoldOnCells();
        }

        private void drawTable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CreateTable();
        }
    }
}
