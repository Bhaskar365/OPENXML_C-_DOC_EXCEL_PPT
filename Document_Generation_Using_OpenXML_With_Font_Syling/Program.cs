
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

Console.WriteLine("Hello, World!");

void DOC()
{
    string filePath = "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Document_Generation_Using_OpenXML_With_Font_Syling\\DocumentFolder\\generatedDoc.docx";

    // Check if the file exists
    if (!System.IO.File.Exists(filePath))
    {
        // Create a new Word document
        using (WordprocessingDocument newDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            // Add the main document part
            MainDocumentPart mainPart = newDoc.AddMainDocumentPart();

            // Create the document structure
            mainPart.Document = new Document();

            // Add some initial content
            mainPart.Document.Append(new Body(
                new Paragraph(
                    new Run(new Text("Initial content in the document."))
                )
            ));
        }
    }

    // Open the existing Word document
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
    {
        // Get the main document part
        MainDocumentPart mainPart = doc.MainDocumentPart;

        // Create a run with the text content
        Run run = new Run(new Text("Text with changed font and background color."));

        RunProperties runProperties = new RunProperties();

        // Create a font size element
        FontSize fontSize = new FontSize() { Val = "24" }; // Change the font size as needed

        // Create a color element for the font
        DocumentFormat.OpenXml.Wordprocessing.Color fontColor = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "FF0000" }; // Change the color as needed

        // Create a background color element using Shading
        Shading shading = new Shading() { Fill = "FFFF00" }; // Change the color as needed

        // Create a run fonts element to specify the font family
        RunFonts runFonts = new RunFonts() { Ascii = "Bauhaus 93" }; // Change "Arial" to the desired font family

        // Apply the font size, font color, background color, and font family to the run properties
        runProperties.Append(fontSize);
        runProperties.Append(fontColor);
        runProperties.Append(shading); // Apply shading to the run
        runProperties.Append(runFonts); // Apply font family

        // Apply the run properties to the run containing the text
        run.RunProperties = runProperties;

        // Create a paragraph and add the run to it
        Paragraph paragraph = new Paragraph(run);

        // Append the paragraph to the main body
        mainPart.Document.Body.AppendChild(paragraph);

        // Save the changes
        mainPart.Document.Save();
        Console.WriteLine("Generating Document");
        Console.WriteLine("Done");
    }
}

DOC();

