﻿using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string mypath = "";
            Word.Application wordApp = Application;
            wordApp.Documents.Open(mypath + "Search and Replace.doc");
            object replaceAll = Word.WdReplace.wdReplaceAll;

            this.Application.Selection.Find.ClearFormatting();
            this.Application.Selection.Find.Text = "find me";

            this.Application.Selection.Find.Replacement.ClearFormatting();
            this.Application.Selection.Find.Replacement.Text = "Found";

            this.Application.Selection.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
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
