﻿using Aspose.Words;
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
            Document doc = new Document();
            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all.
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.A4;
            doc.Save(FilePath + "ModifyPageSetupInAllSections.doc");
        }
    }
}
