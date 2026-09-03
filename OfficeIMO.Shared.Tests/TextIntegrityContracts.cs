using OfficeIMO.Provenance;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class TextIntegrityContracts {
    [Fact]
    public void InspectionReportsExactUnicodeSignalsWithoutAuthorshipClaims() {
        string text = "\uFEFFstart\u200Bfa\u200Crsi\u200D \u202Eabc\u202C \U000E0061\u00A0tail";

        OfficeTextIntegrityReport report = OfficeTextIntegrityInspector.Inspect(text, location: "Document/Paragraph[1]");

        Assert.Collection(report.Findings,
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.ZeroWidthSpace, 6, 0x200B),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.ZeroWidthNonJoiner, 9, 0x200C),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.ZeroWidthJoiner, 13, 0x200D),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.BidirectionalControl, 15, 0x202E),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.BidirectionalControl, 19, 0x202C),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.UnicodeTag, 21, 0xE0061),
            item => AssertFinding(item, OfficeTextIntegrityFindingKind.TypographicSpace, 23, 0x00A0));
        Assert.True(report.HasPotentiallyDangerousFindings);
        Assert.Equal(2, report.Findings[5].TextLength);
        Assert.Equal("U+E0061", report.Findings[5].UnicodeNotation);
        Assert.All(report.Findings, item => Assert.Equal("Document/Paragraph[1]", item.Location));
    }

    [Fact]
    public void LeadingBomAndOptionalTypographyCanBeExcluded() {
        OfficeTextIntegrityReport report = OfficeTextIntegrityInspector.Inspect(
            "\uFEFFword\u00A0word\uFE0F",
            new OfficeTextIntegrityOptions {
                IgnoreLeadingByteOrderMark = true,
                IncludeTypographicSpaces = false,
                IncludeVariationSelectors = false
            });

        Assert.Empty(report.Findings);
    }

    [Fact]
    public void CleanerRemovesOnlyExplicitlySelectedFindings() {
        string text = "keep\u200B فارسی\u200C text\u00A0";
        OfficeTextIntegrityReport report = OfficeTextIntegrityInspector.Inspect(text);
        OfficeTextIntegrityFinding zeroWidthSpace = Assert.Single(
            report.Findings,
            item => item.Kind == OfficeTextIntegrityFindingKind.ZeroWidthSpace);

        string cleaned = OfficeTextIntegrityCleaner.RemoveSelected(text, new[] { zeroWidthSpace });

        Assert.Equal("keep فارسی\u200C text\u00A0", cleaned);
    }

    [Fact]
    public void CleanerRejectsStaleOrOverlappingSelections() {
        string text = "a\u200Bb";
        OfficeTextIntegrityFinding finding = Assert.Single(OfficeTextIntegrityInspector.Inspect(text).Findings);
        var overlapping = new OfficeTextIntegrityFinding(
            OfficeTextIntegrityFindingKind.ZeroWidthSpace,
            OfficeTextIntegrityRisk.ContextDependent,
            finding.TextOffset,
            finding.TextLength,
            finding.CodePoint);

        Assert.Throws<ArgumentException>(() => OfficeTextIntegrityCleaner.RemoveSelected("axb", new[] { finding }));
        Assert.Throws<ArgumentException>(() => OfficeTextIntegrityCleaner.RemoveSelected(text, new[] { finding, overlapping }));
    }

    [Fact]
    public void InspectionBoundsFindingsAndMalformedUtf16() {
        Assert.Throws<InvalidDataException>(() => OfficeTextIntegrityInspector.Inspect(
            "\u200B\u200B",
            new OfficeTextIntegrityOptions { MaxFindings = 1 }));

        OfficeTextIntegrityFinding finding = Assert.Single(
            OfficeTextIntegrityInspector.Inspect("x\uD800y").Findings);
        Assert.Equal(OfficeTextIntegrityFindingKind.UnpairedSurrogate, finding.Kind);
        Assert.Equal(OfficeTextIntegrityRisk.PotentiallyDangerous, finding.Risk);
    }

    [Fact]
    public void InspectionReportsOtherUnicodeFormatControls() {
        OfficeTextIntegrityFinding finding = Assert.Single(
            OfficeTextIntegrityInspector.Inspect("left\u206Aright").Findings);

        AssertFinding(finding, OfficeTextIntegrityFindingKind.InvisibleFormatCharacter, 4, 0x206A);
        Assert.Equal(OfficeTextIntegrityRisk.ContextDependent, finding.Risk);
    }

    [Fact]
    public void FileInspectionEnforcesEncodedByteLimitDuringTheRead() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllBytes(path, new byte[] { 0x61, 0x62 });
        try {
            Assert.Throws<InvalidDataException>(() => OfficeTextIntegrityInspector.InspectFile(
                path,
                new OfficeTextIntegrityOptions { MaxEncodedBytes = 1 }));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void StringInspectionObservesCancellationDuringTraversal() {
        string text = new string('a', 16 * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        Assert.Throws<OperationCanceledException>(() => OfficeTextIntegrityInspector.Inspect(
            text,
            new OfficeTextIntegrityOptions { MaxCharacters = text.Length },
            "LargeText",
            cancellation.Token));
    }

    [Fact]
    public void FileInspectionObservesCancellationBeforeDecodeAndScan() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text", new UTF8Encoding(false));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try {
            Assert.Throws<OperationCanceledException>(() => OfficeTextIntegrityInspector.InspectFile(
                path,
                new OfficeTextIntegrityOptions(),
                "Text",
                cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    private static void AssertFinding(
        OfficeTextIntegrityFinding finding,
        OfficeTextIntegrityFindingKind kind,
        int offset,
        int codePoint) {
        Assert.Equal(kind, finding.Kind);
        Assert.Equal(offset, finding.TextOffset);
        Assert.Equal(codePoint, finding.CodePoint);
    }
}
