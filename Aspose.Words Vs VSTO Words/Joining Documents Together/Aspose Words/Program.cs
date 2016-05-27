using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // We should call this method to clear this document of any existing content.
            dstDoc.RemoveAllChildren();

            int recordCount = 1;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Document srcDoc = new Document(FilePath+"JoinningDocumenttogether(source).docx");

                // Append the source document at the end of the destination document.
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                Document doc2 = new Document(FilePath+"JoinningDocumenttogether(dest).docx");
                dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                // In automation you were required to insert a new section break at this point, however in Aspose.Words we
                // don't need to do anything here as the appended document is imported as separate sectons already.

                // If this is the second document or above being appended then unlink all headers footers in this section
                // from the headers and footers of the previous section.
                if (i > 1)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }
            dstDoc.Save(FilePath +"JoinningDocumenttogether.docx");
        }
    }
}
