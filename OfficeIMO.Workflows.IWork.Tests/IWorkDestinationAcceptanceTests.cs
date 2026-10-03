using OfficeIMO.IWork;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Known_destination_omission_requires_partial_acceptance_before_publication(bool allowPartial) {
        using var files = new Files("numbers", "xlsx");
        byte[] cell = new byte[24]; cell[0] = 5; cell[1] = 2; cell[8] = 0x22; cell[20] = 1;
        BitConverter.GetBytes(42d).CopyTo(cell, 12);
        byte[] row = Join(V(1, 0), V(2, 1), B(6, cell), B(7, [0, 0]));
        byte[] store = Join(R(5, 13), B(3, B(1, Join(V(1, 0), R(2, 12)))));
        byte[] padding = Join(U((ulong)((1 << 3) | 5)), BitConverter.GetBytes(4f));
        byte[] records = Join(
            A(1, 1, R(1, 2)), A(2, 2, Join(B(1, System.Text.Encoding.UTF8.GetBytes("Sheet")), R(2, 10))),
            A(10, 6000, R(2, 11)), A(11, 6001, Join(V(6, 1), V(7, 1), B(4, store))),
            A(12, 6002, Join(V(1, 1), V(2, 0), V(3, 1), V(4, 1), B(5, row), V(6, 5), V(7, 1))),
            A(13, 6005, Join(V(1, 4), B(3, Join(V(1, 1), R(4, 30))))),
            A(30, 6004, B(11, B(9, padding))));
        WritePackage(files.Input, records);
        Assert.True(IWorkSourceDocument.Open(files.Input).ReadNumbers().HasEditableContent);
        var request = files.Request("numbers-xlsx");
        request.RegisteredConversionSettings = allowPartial ? new IWorkWorkflowSettings {
            ConversionOptions = new IWorkConversionOptions { AllowPartialEditableReconstruction = true }
        } : null;

        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.Equal(allowPartial, result.Succeeded);
        Assert.Equal(allowPartial, File.Exists(files.Output));
        if (allowPartial) {
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("True", evidence.Facts["partialEditableReconstruction"]);
            Assert.Contains(evidence.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(files.Output);
            Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        } else {
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        }
    }
}
