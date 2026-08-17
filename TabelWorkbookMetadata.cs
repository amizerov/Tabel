namespace Tabel;

using System.Runtime.InteropServices;

internal static class TabelWorkbookMetadata
{
    private const string RoleMarkerName = "_TabelFileRole";

    public const string FilledTimesheetRole = "FilledTimesheet";
    public const string ComparisonRole = "Comparison";

    public static void SetRole(dynamic workbook, string role)
    {
        dynamic? oldMarker = null;
        dynamic? newMarker = null;

        try
        {
            try
            {
                oldMarker = workbook.Names.Item(RoleMarkerName);
                oldMarker.Delete();
            }
            catch { }
            finally
            {
                ReleaseComObject(oldMarker);
            }

            newMarker = workbook.Names.Add(RoleMarkerName, $"=\"{role}\"", false);
        }
        finally
        {
            ReleaseComObject(newMarker);
        }
    }

    public static string GetExplicitRole(dynamic workbook)
    {
        dynamic? marker = null;

        try
        {
            marker = workbook.Names.Item(RoleMarkerName);
            string refersTo = Convert.ToString(marker.RefersTo)?.Trim() ?? string.Empty;

            if (refersTo.StartsWith("=\"", StringComparison.Ordinal)
                && refersTo.EndsWith("\"", StringComparison.Ordinal)
                && refersTo.Length >= 3)
                return refersTo[2..^1];

            return refersTo.TrimStart('=').Trim('"');
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            ReleaseComObject(marker);
        }
    }

    public static bool IsLegacyComparison(dynamic workbook)
    {
        return HasLegacyComparisonIndicators(workbook);
    }

    private static bool HasLegacyComparisonIndicators(dynamic workbook)
    {
        int sheetCount;
        try { sheetCount = workbook.Worksheets.Count; }
        catch { return false; }

        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            dynamic? sheet = null;
            dynamic? comments = null;
            dynamic? usedRange = null;

            try
            {
                sheet = workbook.Worksheets[sheetIndex];
                comments = sheet.Comments;
                int commentCount = comments.Count;

                for (int commentIndex = 1; commentIndex <= commentCount; commentIndex++)
                {
                    dynamic? comment = null;
                    try
                    {
                        comment = comments.Item(commentIndex);
                        string text = Convert.ToString(comment.Text()) ?? string.Empty;
                        if (text.Contains("Сравнение:", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    catch { }
                    finally
                    {
                        ReleaseComObject(comment);
                    }
                }

                usedRange = sheet.UsedRange;
                if (RangeContainsColor(workbook.Application, usedRange, RGB(255, 255, 0))
                    || RangeContainsColor(workbook.Application, usedRange, RGB(198, 239, 206)))
                    return true;
            }
            catch { }
            finally
            {
                ReleaseComObject(usedRange);
                ReleaseComObject(comments);
                ReleaseComObject(sheet);
            }
        }

        return false;
    }

    private static bool RangeContainsColor(dynamic application, dynamic range, int color)
    {
        dynamic? found = null;

        try
        {
            application.FindFormat.Clear();
            application.FindFormat.Interior.Color = color;
            found = range.Find(
                "*",
                Type.Missing,
                -4163, // xlValues
                2,     // xlPart
                1,     // xlByRows
                1,     // xlNext
                false,
                Type.Missing,
                true); // SearchFormat
            return found != null;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(found);
            try { application.FindFormat.Clear(); } catch { }
        }
    }

    private static int RGB(int red, int green, int blue)
    {
        return red + (green << 8) + (blue << 16);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value == null)
            return;

        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
