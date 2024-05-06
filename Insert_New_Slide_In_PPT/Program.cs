// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.Drawing;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

Console.WriteLine("Hello, World!");

InsertNewSlide();

// Insert a slide into the specified presentation.
static void InsertNewSlide()
{
    Console.WriteLine("Insertion code started");
    string presentationFile = "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Insert_New_Slide_In_PPT\\PPT_Folder\\NewBlankPpt.pptx";
    int position = 1;
    string slideTitle = "First Slide Insertion";
    // Open the source document as read/write. 
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Pass the source document and the position and title of the slide to be inserted to the next method.
        InsertNewSlideFromPresentation(presentationDocument, position, slideTitle);
    }
}

// Insert the specified slide into the presentation at the specified position.
static void InsertNewSlideFromPresentation(PresentationDocument presentationDocument, int position, string slideTitle)
{
    PresentationPart presentationPart = presentationDocument.PresentationPart;

    // Verify that the presentation is not empty.
    if (presentationPart == null)
    {
        throw new InvalidOperationException("The presentation document is empty.");
    }

    // Declare and instantiate a new slide.
    Slide slide = new Slide(new CommonSlideData(new DocumentFormat.OpenXml.Presentation.ShapeTree()));
    uint drawingObjectId = 1;

    // Construct the slide content.            
    // Specify the non-visual properties of the new slide.
    CommonSlideData commonSlideData = slide.CommonSlideData ?? slide.AppendChild(new CommonSlideData());
    DocumentFormat.OpenXml.Presentation.ShapeTree shapeTree = commonSlideData.ShapeTree ?? commonSlideData.AppendChild(new DocumentFormat.OpenXml.Presentation.ShapeTree());
    DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeProperties nonVisualProperties = shapeTree.AppendChild(new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeProperties());
    nonVisualProperties.NonVisualDrawingProperties = new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "" };
    nonVisualProperties.NonVisualGroupShapeDrawingProperties = new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeDrawingProperties();

    // Specify the group shape properties of the new slide.
    shapeTree.AppendChild(new DocumentFormat.OpenXml.Presentation.GroupShapeProperties());

    // Declare and instantiate the title shape of the new slide.
    DocumentFormat.OpenXml.Presentation.Shape titleShape = shapeTree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape());

    drawingObjectId++;

    // Specify the required shape properties for the title shape. 
    titleShape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties
        (new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
        new DocumentFormat.OpenXml.Drawing.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
    titleShape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();

    // Specify the text of the title shape.
    titleShape.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.Text() { Text = slideTitle })));

    // Declare and instantiate the body shape of the new slide.
    DocumentFormat.OpenXml.Presentation.Shape bodyShape = shapeTree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape());
    drawingObjectId++;

    // Specify the required shape properties for the body shape.
    bodyShape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
            new DocumentFormat.OpenXml.Drawing.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
    bodyShape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();

    // Specify the text of the body shape.
    bodyShape.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph());

    // Create the slide part for the new slide.
    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

    // Save the new slide part.
    slide.Save(slidePart);

    // Modify the slide ID list in the presentation part.
    // The slide ID list should not be null.
    SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

    // Find the highest slide ID in the current list.
    uint maxSlideId = 1;
    SlideId prevSlideId = null;

    foreach (SlideId slideId in slideIdList.ChildElements)
    {
        if (slideId.Id > maxSlideId)
        {
            maxSlideId = slideId.Id;
        }

        position--;
        if (position == 0)
        {
            prevSlideId = slideId;
        }

    }

    maxSlideId++;

    // Get the ID of the previous slide.
    SlidePart lastSlidePart;

    if (prevSlideId != null && prevSlideId.RelationshipId != null)
    {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
    }
    else
    {
        string firstRelId = ((SlideId)slideIdList.ChildElements[0]).RelationshipId;
        // If the first slide does not contain a relationship ID, throw an exception.
        if (firstRelId == null)
        {
            throw new ArgumentNullException(nameof(firstRelId));
        }

        lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
    }

    // Use the same slide layout as that of the previous slide.
    if (lastSlidePart.SlideLayoutPart != null)
    {
        slidePart.AddPart(lastSlidePart.SlideLayoutPart);
    }

    // Insert the new slide into the slide list after the previous slide.
    SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
    newSlideId.Id = maxSlideId;
    newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

    // Save the modified presentation.
    presentationPart.Presentation.Save();
    Console.WriteLine("Insertion code finished");
    Console.WriteLine("Slide insertion successful");
}
