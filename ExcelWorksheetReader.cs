namespace Tabel;

using System.Runtime.InteropServices;

internal sealed record ExcelWorksheetData(object[,] Values, int Rows, int Columns);

internal static class ExcelWorksheetReader
{
    private const int XlFormulas = -4123;
    private const int XlPart = 2;
    private const int XlByRows = 1;
    private const int XlByColumns = 2;
    private const int XlPrevious = 2;

    public static ExcelWorksheetData Read(dynamic sheet, int? maxRows = null, int? maxColumns = null)
    {
        dynamic? cells = null;
        dynamic? lastRowCell = null;
        dynamic? lastColumnCell = null;
        dynamic? firstCell = null;
        dynamic? endCell = null;
        dynamic? range = null;

        try
        {
            cells = sheet.Cells;
            lastRowCell = cells.Find(
                "*", Type.Missing, XlFormulas, XlPart, XlByRows, XlPrevious,
                false, Type.Missing, false);
            lastColumnCell = cells.Find(
                "*", Type.Missing, XlFormulas, XlPart, XlByColumns, XlPrevious,
                false, Type.Missing, false);

            int rows = lastRowCell == null ? 1 : Convert.ToInt32(lastRowCell.Row);
            int columns = lastColumnCell == null ? 1 : Convert.ToInt32(lastColumnCell.Column);

            if (maxRows.HasValue)
                rows = Math.Min(rows, maxRows.Value);
            if (maxColumns.HasValue)
                columns = Math.Min(columns, maxColumns.Value);

            firstCell = cells[1, 1];
            endCell = cells[rows, columns];
            range = sheet.Range[firstCell, endCell];

            object value = range.Value2;
            return new ExcelWorksheetData(ToArray2D(value), rows, columns);
        }
        finally
        {
            ReleaseComObject(range);
            ReleaseComObject(endCell);
            ReleaseComObject(firstCell);
            ReleaseComObject(lastColumnCell);
            ReleaseComObject(lastRowCell);
            ReleaseComObject(cells);
        }
    }

    private static object[,] ToArray2D(object value)
    {
        if (value is object[,] array)
            return array;

        var result = new object[2, 2];
        result[1, 1] = value;
        return result;
    }

    private static void ReleaseComObject(object? value)
    {
        if (value == null)
            return;

        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}
