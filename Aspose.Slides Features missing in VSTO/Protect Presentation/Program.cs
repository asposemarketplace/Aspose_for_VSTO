using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
      private static  string FilePath = @"..\..\..\Sample Files\";
        static void Main(string[] args)
        {
            ApplyingProtection();
            RemovingProtection();
        }
        static void ApplyingProtection()
        {
            string srcFileName = FilePath + "Sample Presentation.pptx";
            string destFileName = FilePath + "ProtectedSample.pptx";
            
            //Instatiate Presentation class that represents a PPTX file
            Presentation pTemplate = new Presentation(srcFileName);//Instatiate Presentation class that represents a PPTX file
           

            //ISlide object for accessing the slides in the presentation
            ISlide slide = pTemplate.Slides[0];

            //IShape object for holding temporary shapes
            IShape shape;

            //Traversing through all the slides in the presentation
            for (int slideCount = 0; slideCount < pTemplate.Slides.Count; slideCount++)
            {
                slide = pTemplate.Slides[slideCount];

                //Travesing through all the shapes in the slides
                for (int count = 0; count < slide.Shapes.Count; count++)
                {
                    shape = slide.Shapes[count];

                    //if shape is autoshape
                    if (shape is AutoShape)
                    {
                        //Type casting to Auto shape and  getting auto shape lock
                        AutoShape Ashp = shape as AutoShape;
                        IAutoShapeLock AutoShapeLock = Ashp.ShapeLock;

                        //Applying shapes locks
                        AutoShapeLock.PositionLocked = true;
                        AutoShapeLock.SelectLocked = true;
                        AutoShapeLock.SizeLocked = true;
                    }

                    //if shape is group shape
                    else if (shape is GroupShape)
                    {
                        //Type casting to group shape and  getting group shape lock
                        GroupShape Group = shape as GroupShape;
                        IGroupShapeLock groupShapeLock = Group.ShapeLock;

                        //Applying shapes locks
                        groupShapeLock.GroupingLocked = true;
                        groupShapeLock.PositionLocked = true;
                        groupShapeLock.SelectLocked = true;
                        groupShapeLock.SizeLocked = true;
                    }

                    //if shape is a connector
                    else if (shape is Connector)
                    {
                        //Type casting to connector shape and  getting connector shape lock
                        Connector Conn = shape as Connector;
                        IConnectorLock ConnLock = Conn.ShapeLock;

                        //Applying shapes locks
                        ConnLock.PositionMove = true;
                        ConnLock.SelectLocked = true;
                        ConnLock.SizeLocked = true;
                    }

                    //if shape is picture frame
                    else if (shape is PictureFrame)
                    {
                        //Type casting to picture frame shape and  getting picture frame shape lock
                        PictureFrame Pic = shape as PictureFrame;
                        IPictureFrameLock PicLock = Pic.ShapeLock;

                        //Applying shapes locks
                        PicLock.PositionLocked = true;
                        PicLock.SelectLocked = true;
                        PicLock.SizeLocked = true;
                    }
                }
            }
            //Saving the presentation file
            pTemplate.Save(destFileName, Aspose.Slides.Export.SaveFormat.Pptx);
        }
        static void RemovingProtection()
        {
            string srcFileName = FilePath + "ProtectedSample.pptx";
            string destFileName = FilePath + "unProtectedSample.pptx";
            //Open the desired presentation
            Presentation pTemplate = new Presentation(srcFileName);

            //ISlide object for accessing the slides in the presentation
            ISlide slide = pTemplate.Slides[0];

            //IShape object for holding temporary shapes
            IShape shape;

            //Traversing through all the slides in presentation
            for (int slideCount = 0; slideCount < pTemplate.Slides.Count; slideCount++)
            {
                slide = pTemplate.Slides[slideCount];

                //Travesing through all the shapes in the slides
                for (int count = 0; count < slide.Shapes.Count; count++)
                {
                    shape = slide.Shapes[count];

                    //if shape is autoshape
                    if (shape is AutoShape)
                    {
                        //Type casting to Auto shape and  getting auto shape lock
                        AutoShape Ashp = shape as AutoShape;
                        IAutoShapeLock AutoShapeLock = Ashp.ShapeLock;

                        //Applying shapes locks
                        AutoShapeLock.PositionLocked = false;
                        AutoShapeLock.SelectLocked = false;
                        AutoShapeLock.SizeLocked = false;
                    }

                    //if shape is group shape
                    else if (shape is GroupShape)
                    {
                        //Type casting to group shape and  getting group shape lock
                        GroupShape Group = shape as GroupShape;
                        IGroupShapeLock groupShapeLock = Group.ShapeLock;

                        //Applying shapes locks
                        groupShapeLock.GroupingLocked = false;
                        groupShapeLock.PositionLocked = false;
                        groupShapeLock.SelectLocked = false;
                        groupShapeLock.SizeLocked = false;
                    }

                    //if shape is Connector shape
                    else if (shape is Connector)
                    {
                        //Type casting to connector shape and  getting connector shape lock
                        Connector Conn = shape as Connector;
                        IConnectorLock ConnLock = Conn.ShapeLock;

                        //Applying shapes locks
                        ConnLock.PositionMove = false;
                        ConnLock.SelectLocked = false;
                        ConnLock.SizeLocked = false;
                    }

                    //if shape is picture frame
                    else if (shape is PictureFrame)
                    {
                        //Type casting to pitcture frame shape and  getting picture frame shape lock
                        PictureFrame Pic = shape as PictureFrame;
                        IPictureFrameLock PicLock = Pic.ShapeLock;

                        //Applying shapes locks
                        PicLock.PositionLocked = false;
                        PicLock.SelectLocked = false;
                        PicLock.SizeLocked = false;
                    }
                }

            }
            //Saving the presentation file
            pTemplate.Save(destFileName, Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
