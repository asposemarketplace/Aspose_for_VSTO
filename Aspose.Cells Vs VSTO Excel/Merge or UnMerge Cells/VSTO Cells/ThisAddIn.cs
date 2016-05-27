﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Reflection;

namespace VSTO_Cells
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Instantiate the Application object.
            Excel.Application excelApp = Application;

            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "MergeUnmergeCells.xlsx";

            //Open the excel file.
            excelApp.Workbooks.Open(fileName, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            //Get the range of cells i.e.., A1:C1.
            Excel.Range rng1 = excelApp.get_Range("A1", "C1");

            //Merge the cells.
            rng1.Merge(Missing.Value);

            rng1 = excelApp.get_Range("A1", Missing.Value);

            //UnMerge the cell.
            rng1.UnMerge();

            //Save the file.
            excelApp.ActiveWorkbook.Save();

            //Quit the Application.
            //excelApp.Quit();
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
