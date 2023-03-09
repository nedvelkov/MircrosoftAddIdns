using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Data;

namespace ExcelAddInWithButtons
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void ToggleBoldOnCells()
        {
            var range = this.Application.ActiveWindow.RangeSelection;
            range.Font.Bold = !range.Font.Bold;
        }

        internal void CreateTable()
        {
            int[] Ages = { 32, 44, 28, 61 };
            string[] Names = { "Reggie", "Sally", "Henry", "Christine" };

            // Create a data table with two columns.
            DataSet ds = new DataSet();
            DataTable table = ds.Tables.Add("Customers");
            DataColumn column1 = new DataColumn("Names", typeof(string));
            DataColumn column2 = new DataColumn("Ages", typeof(int));
            table.Columns.Add(column1);
            table.Columns.Add(column2);

            // Add the four rows of data to the table.
            DataRow row;
            for (int i = 0; i < 4; i++)
            {
                row = table.NewRow();
                row["Names"] = Names[i];
                row["Ages"] = Ages[i];
                table.Rows.Add(row);
            }

            Worksheet worksheet = Globals.Factory.GetVstoObject(
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            Microsoft.Office.Tools.Excel.ListObject list1;
            Excel.Range cell = worksheet.Range["$A$1:$D$4"];
            list1 = worksheet.Controls.AddListObject(cell, "list1");


            // Bind the list object to the table.
            list1.SetDataBinding(ds, "Customers");
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
