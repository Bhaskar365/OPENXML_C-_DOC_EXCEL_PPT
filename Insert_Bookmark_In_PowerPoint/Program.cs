// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

Console.WriteLine("Hello, World!");

static void AddBookmark(string file, string bookmarkName, int xPosition, int yPosition)
{
    try
    {
        using (var presentation = PresentationDocument.Open(file, true))
        {
            var slidePart = presentation.PresentationPart.GetPartsOfType<SlidePart>().FirstOrDefault();
            if (slidePart != null)
            {
                var tree = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>().FirstOrDefault();
                if (tree != null)
                {
                    var bookmarkShape = new DocumentFormat.OpenXml.Drawing.Shape(
                        new DocumentFormat.OpenXml.Drawing.NonVisualShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties { Name = bookmarkName, Id = (UInt32)tree.ChildElements.Count },
                            new DocumentFormat.OpenXml.Drawing.NonVisualShapeDrawingProperties(),
                            new DocumentFormat.OpenXml.Office2010.Drawing.ChartDrawing.ApplicationNonVisualDrawingProperties()),
                        new DocumentFormat.OpenXml.Drawing.ShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.Transform2D(new DocumentFormat.OpenXml.Drawing.Offset { X = xPosition, Y = yPosition }),
                            new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle },
                            new DocumentFormat.OpenXml.Drawing.NoFill(),
                            new Inline()));

                    tree.AppendChild(bookmarkShape);
                    Console.WriteLine("Bookmark added successfully!");
                }
                else
                {
                    Console.WriteLine("Slide does not contain a ShapeTree.");
                }
            }
            else
            {
                Console.WriteLine("Presentation does not contain any slides.");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error adding bookmark: " + ex.Message);
    }
}
AddBookmark("C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Insert_Bookmark_In_PowerPoint\\PPTFolder\\blankPPTForBookmark.pptx", "MyBookmark", 200000, 200000);