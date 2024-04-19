1.) Document file download instruction

Clone the repo and then give the local path of your system for triggering document file in the folder I am using . If you do not want the folder , use the location of your project tree.

// Create a color element for the font
DocumentFormat.OpenXml.Wordprocessing.Color fontColor = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "FF0000" }; // Change the color as needed

// Create a background color element using Shading
Shading shading = new Shading() { Fill = "FFFF00" }; // Change the color as needed

// Create a run fonts element to specify the font family
RunFonts runFonts = new RunFonts() { Ascii = "Bauhaus 93" }; // Change "Bauhaus 93" to the desired font family

Use as required.
