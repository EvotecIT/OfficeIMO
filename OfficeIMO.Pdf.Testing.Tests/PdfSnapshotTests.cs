namespace OfficeIMO.Pdf.Testing.Tests;

public sealed class PdfSnapshotTests : IDisposable {
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-snapshot-" + Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
    }

    [Fact]
    public void OneGlyphChangeFailsTextAndPixelChecks() {
        Approve("report", SampleReport.Create("1,250.00"));

        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(
            () => PdfSnapshot.Verify(SampleReport.Create("1,256.00"), "report", Options()));

        AssertDiffLine(failure.Message, "- ", "| Total | 1,250.00 |");
        AssertDiffLine(failure.Message, "+ ", "| Total | 1,256.00 |");
        Assert.Contains("Page 1 pixels differ", failure.Message);
    }

    [Fact]
    public void IgnoredRegionAbsorbsPixelChangesOnlyInsideIt() {
        Approve("badge", SampleReport.Create("10", badge: OfficeColor.FromRgb(0, 128, 0)));
        byte[] changedBadge = SampleReport.Create("10", badge: OfficeColor.FromRgb(200, 0, 0));

        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(
            () => PdfSnapshot.Verify(changedBadge, "badge", Options()));
        Assert.Contains("Page 1 pixels differ", failure.Message);
        Assert.DoesNotContain("Structure differs", failure.Message);

        PdfSnapshotOptions ignoringBadge = Options();
        ignoringBadge.Visual.IgnoredRegions.Add(SampleReport.BadgeRegion);
        PdfSnapshot.Verify(changedBadge, "badge", ignoringBadge);

        PdfSnapshotOptions ignoringElsewhere = Options();
        ignoringElsewhere.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, SampleReport.BadgeRegion.Y + SampleReport.BadgeRegion.Height, 1836, 1200));
        Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(changedBadge, "badge", ignoringElsewhere));
    }

    [Fact]
    public void PageCountChangeIsAStructuralFailure() {
        Approve("pages", SampleReport.Create("10"));

        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(
            () => PdfSnapshot.Verify(SampleReport.Create("10", pages: 2), "pages", Options()));

        AssertDiffLine(failure.Message, "+ ", "\"pageCount\": 2");
        Assert.Contains("Pages: PageCount: expected 1, actual 2.", failure.Message);
    }

    [Fact]
    public void MissingApprovedPageImagesAreNamed() {
        byte[] pdf = SampleReport.Create("10");
        Approve("partial", pdf);
        File.Delete(Path.Combine(_directory, "partial.page-001.png"));

        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(
            () => PdfSnapshot.Verify(pdf, "partial", Options()));

        Assert.Contains("Approved page images are missing: partial.page-001.png.", failure.Message);
    }

    [Fact]
    public void ReceivedFilesUseStablePathsAndAreRemovedOnceTheSnapshotMatches() {
        Approve("stable", SampleReport.Create("10"));
        byte[] changed = SampleReport.Create("11");
        string[] expected = {
            "stable.received.structure.json",
            "stable.received.page-001.png",
            "stable.received.html"
        };

        PdfSnapshotMismatchException first = Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(changed, "stable", Options()));
        PdfSnapshotMismatchException second = Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(changed, "stable", Options()));

        Assert.Equal(expected.Select(name => Path.Combine(_directory, name)), first.ReceivedFiles);
        Assert.Equal(first.ReceivedFiles, second.ReceivedFiles);
        Assert.Equal(Path.Combine(_directory, "stable.received.html"), first.ReviewPath);
        Assert.All(first.ReceivedFiles, path => Assert.Contains(path, first.Message));

        PdfSnapshot.Verify(SampleReport.Create("10"), "stable", Options());
        Assert.Empty(Directory.GetFiles(_directory, "*.received.*"));
    }

    [Fact]
    public void ExactModeFailsOnByteChangesThatRenderTheSame() {
        PdfSnapshotOptions exact = Options();
        exact.Pdf = PdfSnapshotPdfMode.Exact;
        byte[] approved = SampleReport.Create("10");
        Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(approved, "exact", exact));
        PdfSnapshot.Accept(_directory);
        byte[] retitled = PdfDocument.Load(new MemoryStream(approved)).UpdateProperties(title: "Renamed").ToBytes();

        PdfSnapshot.Verify(approved, "exact", exact);
        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(retitled, "exact", exact));

        Assert.StartsWith("PDF bytes differ", failure.Message.Split('\n')[1], StringComparison.Ordinal);
        Assert.DoesNotContain("Structure differs", failure.Message);
    }

    [Fact]
    public void ADifferentEmbeddedFontIsReportedAsAFontChange() {
        Approve("font", SampleReport.Create("10"));

        PdfSnapshotMismatchException failure = Assert.Throws<PdfSnapshotMismatchException>(
            () => PdfSnapshot.Verify(SampleReport.Create("10", font: SampleReport.OpenSans), "font", Options()));

        AssertDiffLine(failure.Message, "- ", "OfficeIMOBaselineSans-Regular");
        AssertDiffLine(failure.Message, "+ ", "OpenSans");
        Assert.Contains("Embedded fonts differ", failure.Message);
    }

    [Fact]
    public void FontsFromTheHostMachineFailWithGuidance() {
        byte[] hostFonts = SampleReport.Create("10", embedFonts: false);

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() => PdfSnapshot.Verify(hostFonts, "fonts", Options()));

        Assert.Contains("not embedded", failure.Message);
        Assert.Contains(nameof(PdfSnapshotOptions.AllowHostFonts), failure.Message);
        Assert.False(Directory.Exists(_directory));
    }

    private PdfSnapshotOptions Options() => new() { Directory = _directory };

    private void Approve(string name, byte[] pdf) {
        Assert.Throws<PdfSnapshotMismatchException>(() => PdfSnapshot.Verify(pdf, name, Options()));
        Assert.Equal(1, PdfSnapshot.Accept(_directory));
        PdfSnapshot.Verify(pdf, name, Options());
    }

    private static void AssertDiffLine(string message, string prefix, string text) =>
        Assert.Contains(message.Split('\n'), line => line.StartsWith(prefix, StringComparison.Ordinal) && line.Contains(text));
}
