
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

        string content = "asdfasdf af asda fad f";          // specific content

        //Cell cell = InsertTextInCell("Hello World!", worksheetPart, "A1");
        Cell cell = InsertTextInCell(content, worksheetPart, "A1");

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
    //PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
    //patternFill.Append(new ForegroundColor() { Rgb = "00FF00" }); // Green color
    //fill.Append(patternFill);
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
