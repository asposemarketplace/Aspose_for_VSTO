using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "RemoveHeaderandFooter.docx";
            Word.Document doc = Application.Documents.Open(fileName);
           Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Last.Range.Delete();
           Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Last.Range.Delete();
           
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
