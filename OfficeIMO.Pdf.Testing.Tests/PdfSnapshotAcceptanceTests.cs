namespace OfficeIMO.Pdf.Testing.Tests;

[CollectionDefinition(nameof(AcceptEnvironmentCollection), DisableParallelization = true)]
public sealed class AcceptEnvironmentCollection;

[Collection(nameof(AcceptEnvironmentCollection))]
public sealed class PdfSnapshotAcceptanceTests : IDisposable {
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-snapshot-" + Guid.NewGuid().ToString("N"));

    public void Dispose() {
        Environment.SetEnvironmentVariable(PdfSnapshot.AcceptEnvironmentVariable, null);
        if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
    }

    [Fact]
    public void MissingBaselineFailsUnlessAccepting() {
        var options = new PdfSnapshotOptions { Directory = _directory };
        byte[] pdf = SampleReport.Create("10");

        PdfSnapshotMismatchException missing = Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(pdf, "first", options));
        Assert.Contains("has no approved baseline", missing.Message);
        Assert.Null(missing.ReviewPath);
        Assert.False(File.Exists(Path.Combine(_directory, "first.structure.json")));

        Environment.SetEnvironmentVariable(PdfSnapshot.AcceptEnvironmentVariable, "1");
        PdfSnapshot.Verify(pdf, "first", options);
        Environment.SetEnvironmentVariable(PdfSnapshot.AcceptEnvironmentVariable, null);

        Assert.True(File.Exists(Path.Combine(_directory, "first.structure.json")));
        Assert.True(File.Exists(Path.Combine(_directory, "first.page-001.png")));
        Assert.Empty(Directory.GetFiles(_directory, "*.received.*"));
        PdfSnapshot.Verify(pdf, "first", options);
    }
}
