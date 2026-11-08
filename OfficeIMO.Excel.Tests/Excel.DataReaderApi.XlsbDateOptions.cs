using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OpenReader_XlsbDateOptionChangesKeepCachedAndOrdinaryRowsConsistent(bool initiallyTreatDates) {
        DateTime firstDate = new(2024, 1, 1);
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xlsb, output, leaveOpen: true)) {
            writer.WriteTable("Data", Enumerable.Range(0, 6), new[] { "When", "Id" }, (row, index) => {
                row.Write(firstDate.AddDays(index));
                row.Write(index);
            });
        }

        var options = new ExcelReaderOptions { TreatDatesUsingNumberFormat = initiallyTreatDates };
        using var reader = ExcelDocument.OpenReader(new MemoryStream(output.ToArray()), options);
        for (int index = 0; index < 6; index++) {
            if (index == 1) options.TreatDatesUsingNumberFormat = !initiallyTreatDates;
            Assert.True(reader.Read());
            DateTime expected = firstDate.AddDays(index);
            if (options.TreatDatesUsingNumberFormat) Assert.Equal(expected, Assert.IsType<DateTime>(reader.GetValue(0)));
            else Assert.Equal(expected.ToOADate(), Assert.IsType<double>(reader.GetValue(0)));
            Assert.Equal(index, reader.GetInt32(1));
        }
        Assert.False(reader.Read());
    }
}
