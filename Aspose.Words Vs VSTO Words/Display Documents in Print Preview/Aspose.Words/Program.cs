using Aspose.Words;
using Aspose.Words.Rendering;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileName = "YourFileName.docx";
            Document doc = new Document(FileName);

            PrintPreviewDialog previewDlg = new PrintPreviewDialog();
            // Pass the Aspose.Words print document to the Print Preview dialog.
            previewDlg.Document = new AsposeWordsPrintDocument(doc);
            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "TestName.doc";
            previewDlg.WindowState = FormWindowState.Maximized;
            // Show the appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
        }
    }

}
