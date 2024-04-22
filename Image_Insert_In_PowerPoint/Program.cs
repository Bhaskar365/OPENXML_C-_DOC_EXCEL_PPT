// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;

Console.WriteLine("Hello, World!");

static void AddImage(string file, string image, int xPosition, int yPosition)
{
    using (var presentation = PresentationDocument.Open(file, true))
    {
        var slidePart = presentation.PresentationPart.SlideParts.First();

        var part = slidePart.AddImagePart(ImagePartType.Png);

        using (var stream = File.OpenRead(image))
        {
            part.FeedData(stream);
        }

        var tree = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>().First();

        var picture = new DocumentFormat.OpenXml.Presentation.Picture();

        picture.NonVisualPictureProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties();
        picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
        {
            Name = "My Shape",
            Id = (UInt32)tree.ChildElements.Count - 1
        });

        var nonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
        nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
        {
            NoChangeAspect = true
        });
        picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

        var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
        var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
        {
            Embed = slidePart.GetIdOfPart(part)
        };
        var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
        var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
        {
            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
        };
        var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
        {
            Val = false
        };
        useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
        blipExtension1.Append(useLocalDpi1);
        blipExtensionList1.Append(blipExtension1);
        blip1.Append(blipExtensionList1);
        var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
        stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
        blipFill.Append(blip1);
        blipFill.Append(stretch);
        picture.Append(blipFill);

        picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
        picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();

        // Set X-axis and Y-axis positions
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
        {
            X = xPosition, // X-axis position

            Y = yPosition, // Y-axis position
        });

        // Set the size and type of the image
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
        {
            Cx = 3000000, // Width of the image
            Cy = 3000000, // Height of the image
        });
        picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
        {
            Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle // Shape type
        });

        tree.Append(picture);
        Console.WriteLine("Image Added In PPT");
    }
}

AddImage("C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Image_Insert_In_PowerPoint\\PPT_Folder\\blankPPT.pptx",
         "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Image_Insert_In_PowerPoint\\PPT_Folder\\photo.jpg",
         100000, 100000);