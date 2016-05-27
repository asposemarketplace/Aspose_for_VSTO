using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace VSTO_Cells
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "ProtectDocument.docx";
            //Instantiate the Application object.
            Excel.Application excelApp = Application;
            //Excel.Application excelApp = Application;

            //Open the excel file.

            excelApp.Workbooks.Open(fileName, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            //Protect the workbook specifying a password with Structure and Windows attributes.
            excelApp.ActiveWorkbook.Protect("007", true, true);

            //Save the file.
            excelApp.ActiveWorkbook.Save();

            //Quit the Application.
            excelApp.Quit();

            //Unprotect the workbook specifying its password.
            excelApp.ActiveWorkbook.Unprotect("007");
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
