using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var bookPath = Path.Combine(baseDir, "Book1.xlsx");

        DuplicateExcelSheetByClosedXML(baseDir, bookPath);
        DuplicateExcelSheetByEPPlus(baseDir, bookPath);
    }

    private static void DuplicateExcelSheetByEPPlus(string baseDir, string bookPath)
    {
        var outPath = Path.Combine(baseDir, "Book1 - ClosedXML.xlsx");
        using (var book = new ClosedXML.Excel.XLWorkbook(bookPath))
        {
            book.Worksheet("Sheet1").CopyTo("Sheet2");
            book.SaveAs(outPath);
        }
    }

    private static void DuplicateExcelSheetByClosedXML(string baseDir, string bookPath)
    {
        var outPath = Path.Combine(baseDir, "Book1 - EPPlus.xlsx");
        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(bookPath)))
        using (var book = package.Workbook)
        {
            book.Worksheets.Add("Sheet2", book.Worksheets["Sheet1"]);
            package.SaveAs(new FileInfo(outPath));
        }
    }
}
