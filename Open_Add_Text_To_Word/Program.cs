// See https://aka.ms/new-console-template for more information
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Metadata;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

Console.WriteLine("Hello, World!");

OpenAndAddTextToWordDocument();

static void OpenAndAddTextToWordDocument()
{
    Console.WriteLine("inserting text");
    string filepath = @"C:\OPENXML_C#\Document_Generation_Using_OpenXML_With_Font_Syling\Open_Add_Text_To_Word\Document_Folder\blankDoc1.docx";
    string txt = "Append text in body - OpenAndAddTextToWordDocument";

    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new DocumentFormat.OpenXml.Wordprocessing.Document();
        mainDocumentPart.Document.Body ??= mainDocumentPart.Document.AppendChild(new Body());
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
        Console.WriteLine("insertion successful");
    }
}
