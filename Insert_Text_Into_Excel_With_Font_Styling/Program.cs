
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

CreateSpreadsheetWorkbook();

static void CreateSpreadsheetWorkbook()
{
    Console.WriteLine("Excel writing started");
    string filepath = "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Insert_Text_Into_Excel_With_Font_Styling\\Excel_Destination\\sample.xlsx";

    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
    {
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        Cell cell = InsertTextInCell("Hello World!", worksheetPart, "A1");
        cell.StyleIndex = 1; // Apply the cell format with styling

        workbookPart.AddNewPart<WorkbookStylesPart>();
        var stylesheet = workbookPart.WorkbookStylesPart.Stylesheet = GenerateStylesheet();
        stylesheet.Save();

        workbookPart.Workbook.Save();
        Console.WriteLine("Excel writing successful");
    }
}

static Stylesheet GenerateStylesheet()
{
    Stylesheet stylesheet = new Stylesheet();

    Fonts fonts = new Fonts() { Count = 1U, KnownFonts = true };
    Font font = new Font();
    font.Append(new FontSize() { Val = 11D });
    font.Append(new Color() { Rgb = "DF0001" }); // Red color
    font.Append(new FontName() { Val = "Bradley Hand ITC" });
    font.Append(new Bold());
    fonts.Append(font);

    Fills fills = new Fills() { Count = 2U };
    Fill fill = new Fill();
    PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
    patternFill.Append(new ForegroundColor() { Rgb = "00FF00" }); // Green color
    fill.Append(patternFill);
    fills.Append(fill);

    CellFormats cellFormats = new CellFormats() { Count = 2U };
    cellFormats.AppendChild(new CellFormat() { FontId = 0U, FillId = 0U }); // Bold, Red font with Green fill
    cellFormats.AppendChild(new CellFormat() { FontId = 0U, FillId = 0U, ApplyFont = false }); // No styling (default)

    stylesheet.Append(fonts);
    stylesheet.Append(fills);
    stylesheet.Append(cellFormats);

    return stylesheet;
}

static Cell InsertTextInCell(string text, WorksheetPart worksheetPart, string cellReference)
{
    Cell cell = InsertCellInWorksheet(cellReference, worksheetPart);

    cell.DataType = new EnumValue<CellValues>(CellValues.String);
    cell.CellValue = new CellValue(text);

    return cell;
}

static Cell InsertCellInWorksheet(string cellReference, WorksheetPart worksheetPart)
{
    Worksheet worksheet = worksheetPart.Worksheet;
    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
    Row row;
    Cell cell;

    string[] cellRefParts = SplitCellReference(cellReference);

    uint rowIndex = uint.Parse(cellRefParts[1]);
    row = GetRow(sheetData, rowIndex);

    cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
    if (cell == null)
    {
        cell = new Cell() { CellReference = cellReference };
        row.Append(cell);
    }

    return cell;
}

static Row GetRow(SheetData sheetData, uint rowIndex)
{
    Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
    if (row == null)
    {
        row = new Row() { RowIndex = rowIndex };
        sheetData.Append(row);
    }
    return row;
}

static string[] SplitCellReference(string cellReference)
{
    System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(cellReference, @"([A-Za-z]+)(\d+)");
    return new string[] { match.Groups[1].Value, match.Groups[2].Value };
}

//static void CreateSpreadsheetWorkbook()
//{
//    Console.WriteLine("Creating document");
//    string filepath = "C:\\OPENXML_C#\\Document_Generation_Using_OpenXML_With_Font_Syling\\Insert_Text_Into_Excel_With_Font_Styling\\Excel_Destination\\sample.xlsx";
//    using (var spreadsheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
//    {
//        Console.WriteLine("Creating workbook");
//        spreadsheet.AddWorkbookPart();
//        spreadsheet.WorkbookPart.Workbook = new Workbook();
//        Console.WriteLine("Creating worksheet");
//        var wsPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
//        wsPart.Worksheet = new Worksheet();

//        var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
//        stylesPart.Stylesheet = new Stylesheet();

//        Console.WriteLine("Creating styles");

//        // blank font list
//        stylesPart.Stylesheet.Fonts = new Fonts();
//        stylesPart.Stylesheet.Fonts.Count = 1;
//        stylesPart.Stylesheet.Fonts.AppendChild(new Font());

//        // create fills
//        stylesPart.Stylesheet.Fills = new Fills();

//        // create a solid red fill
//        var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
//        solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
//        solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

//        stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
//        stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
//        stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
//        stylesPart.Stylesheet.Fills.Count = 3;

//        // blank border list
//        stylesPart.Stylesheet.Borders = new Borders();
//        stylesPart.Stylesheet.Borders.Count = 1;
//        stylesPart.Stylesheet.Borders.AppendChild(new Border());

//        // blank cell format list
//        stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
//        stylesPart.Stylesheet.CellStyleFormats.Count = 1;
//        stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

//        // cell format list
//        stylesPart.Stylesheet.CellFormats = new CellFormats();
//        // empty one for index 0, seems to be required
//        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
//        // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
//        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
//        stylesPart.Stylesheet.CellFormats.Count = 2;

//        stylesPart.Stylesheet.Save();

//        Console.WriteLine("Creating sheet data");
//        var sheetData = wsPart.Worksheet.AppendChild(new SheetData());

//        Console.WriteLine("Adding rows / cells...");

//        var row = sheetData.AppendChild(new Row());
//        row.AppendChild(new Cell() { CellValue = new CellValue("This"), DataType = CellValues.String });
//        row.AppendChild(new Cell() { CellValue = new CellValue("is"), DataType = CellValues.String });
//        row.AppendChild(new Cell() { CellValue = new CellValue("a"), DataType = CellValues.String });
//        row.AppendChild(new Cell() { CellValue = new CellValue("test."), DataType = CellValues.String });

//        sheetData.AppendChild(new Row());

//        row = sheetData.AppendChild(new Row());
//        row.AppendChild(new Cell() { CellValue = new CellValue("Value:"), DataType = CellValues.String });
//        row.AppendChild(new Cell() { CellValue = new CellValue("123"), DataType = CellValues.Number });
//        row.AppendChild(new Cell() { CellValue = new CellValue("Formula:"), DataType = CellValues.String });
//        // style index = 1, i.e. point at our fill format
//        row.AppendChild(new Cell() { CellFormula = new CellFormula("B3"), DataType = CellValues.Number, StyleIndex = 1 });

//        Console.WriteLine("Saving worksheet");
//        wsPart.Worksheet.Save();

//        Console.WriteLine("Creating sheet list");
//        var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
//        sheets.AppendChild(new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Test" });

//        Console.WriteLine("Saving workbook");
//        spreadsheet.WorkbookPart.Workbook.Save();

//        Console.WriteLine("Done.");
//    }
//}