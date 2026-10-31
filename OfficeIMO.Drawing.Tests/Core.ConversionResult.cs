using System.Threading.Tasks;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeConversionResultTests {
    [Theory]
    [InlineData(OfficeConversionLossKind.None)]
    [InlineData(OfficeConversionLossKind.Omission)]
    public async Task AsyncResultsCarryTheValueAndUnifiedLossReport(OfficeConversionLossKind kind) {
        var value = new byte[] { 1, 2, 3 };
        var diagnostic = new OfficeDiagnostic("feature-affected", OfficeDiagnosticSeverity.Information,
            "Content assessed.", lossKind: kind);
        var report = new OfficeConversionReport("Word", "Pdf", diagnostics: new[] { diagnostic });
        Task<OfficeConversionResult<byte[]>> operation = Task.FromResult(new OfficeConversionResult<byte[]>(value, report));

        OfficeConversionResult<byte[]> result = await operation;
        var (converted, conversionReport) = result;
        Assert.Same(value, result.Value);
        Assert.Same(value, converted);
        Assert.Same(report, result.Report);
        Assert.Same(report, conversionReport);
        Assert.Equal(kind != OfficeConversionLossKind.None, result.HasLoss);
        if (result.HasLoss)
            Assert.Same(report, Assert.Throws<OfficeConversionException>(conversionReport.ThrowIfLoss).Report);
        else
            conversionReport.ThrowIfLoss();
    }

    [Fact]
    public void ResultsSupportValueTypesAndRejectMissingValuesOrReports() {
        var report = new OfficeConversionReport("Word", "Text");
        var (value, conversionReport) = new OfficeConversionResult<int>(0, report);
        Assert.Equal(0, value);
        Assert.Same(report, conversionReport);
        Assert.Throws<ArgumentNullException>(() => new OfficeConversionResult<string>(null!, report));
        Assert.Throws<ArgumentNullException>(() => new OfficeConversionResult<string>("text", null!));
    }
}
