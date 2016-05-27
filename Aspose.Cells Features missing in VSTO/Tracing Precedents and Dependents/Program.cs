using Aspose.Cells;
using System;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Cells for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Cells for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample File.xlsx";

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(srcFileName);
            Aspose.Cells.Cells cells = workbook.Worksheets[0].Cells;
            Aspose.Cells.Cell cell = cells["B7"];

            //Tracing precedents of the cell B7.
            //The return array contains ranges and cells.
            ReferredAreaCollection ret = cell.GetPrecedents();

            //Printing all the precedent cells' name.
            if (ret != null)
            {
                for (int m = 0; m < ret.Count; m++)
                {
                    ReferredArea area = ret[m];
                    StringBuilder stringBuilder = new StringBuilder();
                    if (area.IsExternalLink)
                    {
                        stringBuilder.Append("[");
                        stringBuilder.Append(area.ExternalFileName);
                        stringBuilder.Append("]");
                    }
                    stringBuilder.Append(area.SheetName);
                    stringBuilder.Append("!");
                    stringBuilder.Append(CellsHelper.CellIndexToName(area.StartRow, area.StartColumn));
                    if (area.IsArea)
                    {
                        stringBuilder.Append(":");
                        stringBuilder.Append(CellsHelper.CellIndexToName(area.EndRow, area.EndColumn));
                    }


                    Console.WriteLine(stringBuilder.ToString());
                }
            }
        }
        static void Main2(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample File.xlsx";

            Workbook workbook = new Workbook(srcFileName);
            Worksheet worksheet = workbook.Worksheets[0];
            var c = worksheet.Cells["A1"];
            var dependents = c.GetDependents(true);
            foreach (var dependent in dependents)
            {
                Console.WriteLine(string.Format("{0} ---- {1} : {2}", dependent.Worksheet.Name, dependent.Name, dependent.Value));
            }
        }
    }
}
