using Aspose.Diagram;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Diagram for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Diagram for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace AsposeSourceCode.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {

            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Sample Diagram.vsdx";
            string destFileName = FilePath + "Sample Diagram Protected.vsdx";
            
            //Load diagram
            Diagram diagram = new Diagram(srcFileName);

            Page page0 = diagram.Pages[0];

            Shape shape = page0.Shapes[0];
            shape.Protection.LockAspect.Value = BOOL.True;
            shape.Protection.LockBegin.Value = BOOL.True;
            shape.Protection.LockCalcWH.Value = BOOL.True;
            shape.Protection.LockCrop.Value = BOOL.True;
            shape.Protection.LockCustProp.Value = BOOL.True;
            shape.Protection.LockDelete.Value = BOOL.True;
            shape.Protection.LockEnd.Value = BOOL.True;
            shape.Protection.LockFormat.Value = BOOL.True;
            shape.Protection.LockFromGroupFormat.Value = BOOL.True;
            shape.Protection.LockGroup.Value = BOOL.True;
            shape.Protection.LockHeight.Value = BOOL.True;
            shape.Protection.LockMoveX.Value = BOOL.True;
            shape.Protection.LockMoveY.Value = BOOL.True;
            shape.Protection.LockRotate.Value = BOOL.True;
            shape.Protection.LockSelect.Value = BOOL.True;
            shape.Protection.LockTextEdit.Value = BOOL.True;
            shape.Protection.LockThemeColors.Value = BOOL.True;
            shape.Protection.LockThemeEffects.Value = BOOL.True;
            shape.Protection.LockVtxEdit.Value = BOOL.True;
            shape.Protection.LockWidth.Value = BOOL.True;
            diagram.Save(destFileName, SaveFileFormat.VSDX);
        }
    }
}
