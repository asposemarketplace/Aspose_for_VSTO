﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            //Specify the template excel file path.
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "InsertRemoveComments.xlsx";

            //Instantiate a new Workbook.
            //Open the excel file.
            Workbook workbook = new Workbook(fileName);

            //Add a Comment to A1 cell.
            int commentIndex = workbook.Worksheets[0].Comments.Add("A1");

            //Accessing the newly added comment
            Comment comment = workbook.Worksheets[0].Comments[commentIndex];

            //Setting the comment note
            comment.Note = "This is my comment";

            workbook.Worksheets[0].Comments.RemoveAt("A1");
            //Save As the excel file.
            workbook.Save(fileName);
        }
    }
}
