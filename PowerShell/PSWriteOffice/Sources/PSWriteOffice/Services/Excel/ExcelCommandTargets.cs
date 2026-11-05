using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
namespace PSWriteOffice.Services.Excel;
/// <summary>Admits Excel-only native workbook sources while retaining explicit save and ownership boundaries.</summary>
internal static class ExcelCommandTargets {
    public static ExcelDocument Document(ExcelDocument? document) => document ?? OfficeDocumentContext.Find<ExcelDocument>() ?? ExcelDslContext.Current?.Document ?? throw new PSArgumentException("Supply -Document or use an Excel content block.");
    public static IEnumerable<ExcelWorkbookCommandScope> Workbooks(PSCmdlet cmdlet, string set, string[] paths, string[] literalPaths, ExcelDocument? document, bool readOnly) {
        if (set == "Document" || set == "DocumentTable" || set == "Default") {
            yield return new ExcelWorkbookCommandScope(Document(document), ownsDocument: false);
            yield break;
        }
        bool literal = set.StartsWith("LiteralPath", StringComparison.Ordinal);
        foreach (string path in OfficePath.ResolveInputs(cmdlet, literal ? literalPaths : paths, literal)) {
            using var scope = new ExcelWorkbookCommandScope(ExcelDocumentService.LoadDocument(path, readOnly, autoSave: false), ownsDocument: true);
            yield return scope;
        }
    }
    public static int Column(object? value) {
        value = OfficeArgumentValue.Unwrap(value);
        if (value == null) throw new PSArgumentException("A column is required.");
        if (value is string text) {
            if (int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int numeric)) return ValidateColumn(numeric);
            int index = ExcelA1.ColumnLettersToIndex(text);
            return ValidateColumn(index);
        }
        if (value is not (byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal))
            throw new PSArgumentException("Column accepts an integer position or Excel column letters.");
        decimal number = Convert.ToDecimal(value, System.Globalization.CultureInfo.InvariantCulture);
        if (number != decimal.Truncate(number)) throw new PSArgumentException("Column positions must be whole numbers.");
        return ValidateColumn(Convert.ToInt32(number));
    }
    private static int ValidateColumn(int index) => index >= 1 && index <= 16384 ? index : throw new PSArgumentException("Column positions range from 1 to 16384.");
    public static (int First, int Last) Rows(string value) {
        string[] parts = value.Replace("$", string.Empty).Split(':');
        if (parts.Length > 2 || !int.TryParse(parts[0], out int first) || first < 1 || first > 1048576) throw new PSArgumentException("RepeatRows uses a row span such as 1:2.");
        int last = parts.Length == 1 ? first : int.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
        if (last < first || last > 1048576) throw new PSArgumentException("RepeatRows must be an ascending row span.");
        return (first, last);
    }
    public static (int First, int Last) Columns(string value) {
        string[] parts = value.Replace("$", string.Empty).Split(':');
        if (parts.Length > 2) throw new PSArgumentException("RepeatColumns uses a column span such as A:B.");
        int first = Column(parts[0]), last = parts.Length == 1 ? first : Column(parts[1]);
        if (last < first) throw new PSArgumentException("RepeatColumns must be an ascending column span.");
        return (first, last);
    }
}
