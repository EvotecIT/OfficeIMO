using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelDelimitedImportTests
{
    [Fact]
    public void ImportDelimitedText_DetectsDelimiterAfterLeadingBlankRecord()
    {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream, autoSave: false);

        var result = document.ImportDelimitedText("\nName;Amount\nAlpha;10\n");

        Assert.Equal(';', result.Delimiter);
        Assert.Equal(1, result.ImportResult.RowCount);
        Assert.Equal(2, result.ImportResult.ColumnCount);
    }
}
