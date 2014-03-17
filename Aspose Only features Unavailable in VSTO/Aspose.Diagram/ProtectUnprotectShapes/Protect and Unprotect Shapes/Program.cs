using Aspose.Diagram;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Protect_and_Unprotect_Shapes
{
    class Program
    {
        static void Main(string[] args)
        {

            string MyDir = @"Files\";
            //Load diagram
            Diagram diagram = new Diagram(MyDir+"ProtectShape.vsd");

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
            diagram.Save(MyDir+"ProtectedShapesFile.vdx", SaveFileFormat.VDX);
        }
    }
}
