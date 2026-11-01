using PSWriteOffice.Models.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelPageBreakRecordService {
    internal static ExcelPageBreakInfo Create(string kind, int index, string sheetName, string? path) =>
        new(kind, index, sheetName, path);
}
