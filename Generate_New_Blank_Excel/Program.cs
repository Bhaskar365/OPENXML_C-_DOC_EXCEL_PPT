using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");

CreateSpreadsheetWorkbook();

static void CreateSpreadsheetWorkbook()
{
    Console.WriteLine("Generating new Excel");
    string filepath = "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Generate_New_Blank_Excel\\Excel_Generation_Folder\\mySheet.xlsx";
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

    // Add a WorkbookPart to the document.
    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    // Add a WorksheetPart to the WorkbookPart.
    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData());

    // Add Sheets to the Workbook.
    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

    // Append a new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
    sheets.Append(sheet);

    workbookPart.Workbook.Save();

    // Dispose the document.
    spreadsheetDocument.Dispose();
    Console.WriteLine("Generated Excel");
}