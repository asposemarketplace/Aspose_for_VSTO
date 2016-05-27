using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Inserting Table.docx";
            Word.Application wordApp = Application;
            wordApp.Documents.Open(fileName);
            CreateDocumentPropertyTable();
            this.Application.ActiveDocument.Save();
        }
        private void CreateDocumentPropertyTable()
        {
            
            object start = 0, end = 0;
            Word.Range rng = this.Application.ActiveDocument.Range(ref start, ref end);

            // Insert a title for the table and paragraph marks. 
            rng.InsertBefore("Document Statistics");
            rng.Font.Name = "Verdana";
            rng.Font.Size = 16;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);

            // Add the table.
            rng.Tables.Add(this.Application.ActiveDocument.Paragraphs[2].Range, 3, 2, ref missing, ref missing);

            // Format the table and apply a style. 
            Word.Table tbl = this.Application.ActiveDocument.Tables[1];
            tbl.Range.Font.Size = 12;
            tbl.Columns.DistributeWidth();

            object styleName = "Table Professional";
            tbl.set_Style(ref styleName);

            // Insert document properties into cells. 
            tbl.Cell(1, 1).Range.Text = "Document Property";
            tbl.Cell(1, 2).Range.Text = "Value";

            tbl.Cell(2, 1).Range.Text = "Subject";
            tbl.Cell(2, 2).Range.Text = ((Office.DocumentProperties)(this.Application.ActiveDocument.BuiltInDocumentProperties))
                [Word.WdBuiltInProperty.wdPropertySubject].Value.ToString();

            tbl.Cell(3, 1).Range.Text = "Author";
            tbl.Cell(3, 2).Range.Text = ((Office.DocumentProperties)(this.Application.ActiveDocument.BuiltInDocumentProperties))
                [Word.WdBuiltInProperty.wdPropertyAuthor].Value.ToString();
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
