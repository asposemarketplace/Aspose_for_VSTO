using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

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
            //Create Aspose.Cells Workbook
            Workbook workbook = new Workbook();

            //Access Aspose.Cells Worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add sample data for pie chart
            //Add headings in A1 and B1
            sheet.Cells["A1"].PutValue("Products");
            sheet.Cells["B1"].PutValue("Users");

            //Add data from A2 till B4
            sheet.Cells["A2"].PutValue("Aspose.Cells");
            sheet.Cells["B2"].PutValue(10000);
            sheet.Cells["A3"].PutValue("Aspose.Slides");
            sheet.Cells["B3"].PutValue(8000);
            sheet.Cells["A4"].PutValue("Aspose.Words");
            sheet.Cells["B4"].PutValue(12000);

            //Chart reference
            Chart productsChart;

            //Add a Pie Chart
            int chartIdx = sheet.Charts.Add(ChartType.Pie, 7, 0, 20, 6);
            productsChart = sheet.Charts[chartIdx];

            //Gets the cells that define the data to be charted
            int seriesIdx = productsChart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            Series nSeries = productsChart.NSeries[seriesIdx];
            nSeries.XValues = "=Sheet1!$A$2:$A$4";

            //Set chart title
            productsChart.Title.Text = "Users";

            //Autofit first column
            sheet.AutoFitColumn(0);

            //Save the copy of workbook as OutputAspose.xlsx
            workbook.Save(FilePath+"CreatePieChart.xlsx");
        }
    }
}
