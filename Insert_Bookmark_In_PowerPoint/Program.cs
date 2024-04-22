// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;

Console.WriteLine("Hello, World!");

static void AddBookmark(string file, string bookmarkName, int xPosition, int yPosition)
{
    using (var presentation = PresentationDocument.Open(file, true))
    {
        Console.WriteLine("Bookmark add started");
        // Get the first slide part
        var slidePart = presentation.PresentationPart.SlideParts.First();

        // Create a new shape tree
        var tree = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>().First();

        // Create a new bookmark shape
        var bookmarkShape = new DocumentFormat.OpenXml.Presentation.Shape();

        // Set non-visual properties of the bookmark
        bookmarkShape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties();
        bookmarkShape.NonVisualShapeProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
        {
            Name = bookmarkName,
            Id = (UInt32)tree.ChildElements.Count // Generate a unique Id for the bookmark
        });

        // Set visual properties of the bookmark
        bookmarkShape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
        bookmarkShape.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();

        // Set X-axis and Y-axis positions
        bookmarkShape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
        {
            X = xPosition, // X-axis position
            Y = yPosition, // Y-axis position
        });

        // Set the size and type of the bookmark
        bookmarkShape.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
        {
            Cx = 50, // Width of the bookmark (set to 0)
            Cy = 50, // Height of the bookmark (set to 0)
        });
        bookmarkShape.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
        {
            Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle // Shape type
        });

        // Append the bookmark to the shape tree
        tree.Append(bookmarkShape);
        Console.WriteLine("Bookmark added");
    }
}

AddBookmark("C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Insert_Bookmark_In_PowerPoint\\PPTFolder\\blankPPTForBookmark.pptx", "MyBookmark", 200000, 200000);