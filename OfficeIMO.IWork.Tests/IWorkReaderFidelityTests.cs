using System.Text.Json;
using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Reader_preserves_explicit_source_fidelity_categories_in_json_diagnostics() {
        var source = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.key"), IWorkDocumentKind.Keynote);
        var kinds = new[] { OfficeConversionLossKind.None, OfficeConversionLossKind.Approximation,
            OfficeConversionLossKind.Omission, OfficeConversionLossKind.Failure, OfficeConversionLossKind.Unassessed };
        var diagnostics = kinds.Select(kind => new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_FIDELITY_" + kind, "Source fidelity assessment.", lossKind: kind)).ToArray();
        var projection = new IWorkKeynoteProjection(source, Array.Empty<IWorkKeynoteSlide>(), null,
            diagnostics, supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var reader = new IWorkReadProjection(result, "source.key", new ReaderOptions(),
            new ReaderIWorkOptions(), CancellationToken.None);
        reader.AddKeynote(projection);
        reader.Complete(source);
        // Attributes are the existing Reader transport for source-specific metadata.
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(result.Diagnostics));
        foreach (var kind in kinds) {
            var diagnostic = json.RootElement.EnumerateArray().Single(item =>
                item.GetProperty("Code").GetString() == "IWORK_FIDELITY_" + kind);
            Assert.Equal(kind.ToString(), diagnostic.GetProperty("Attributes").GetProperty("lossKind").GetString());
        }
    }
}
