1.) Document file generation using OpenXML instruction

Clone the repo and then give the local path of your system for triggering document file in the folder I am using . If you do not want the folder , use the location of your project tree.

// Create a color element for the font
DocumentFormat.OpenXml.Wordprocessing.Color fontColor = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "FF0000" }; // Change the color as needed

// Create a background color element using Shading
Shading shading = new Shading() { Fill = "FFFF00" }; // Change the color as needed

// Create a run fonts element to specify the font family
RunFonts runFonts = new RunFonts() { Ascii = "Bauhaus 93" }; // Change "Bauhaus 93" to the desired font family

run.RunProperties = runProperties;

// Create a paragraph and add the run to it
Paragraph paragraph = new Paragraph(run);

// Append the paragraph to the main body
mainPart.Document.Body.AppendChild(paragraph);

// Insert an additional paragraph
Paragraph additionalParagraph = new Paragraph(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit .......... ")));
mainPart.Document.Body.AppendChild(additionalParagraph);

2.) Insert Image into document using OpenXML -

  For reference, go through this link - https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-picture-into-a-word-processing-document?tabs=cs-0%2Ccs-1%2Ccs-2%2Ccs-3%2Ccs

Use as required.
