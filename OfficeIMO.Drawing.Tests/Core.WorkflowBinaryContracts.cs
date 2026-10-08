using System.IO.Compression;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WorkflowCoreBinaryContractTests {
    [Fact]
    public void PublishedRegularFileSignatureRetainsPhysicalRootGuard() {
        Func<string, string, int, FileStream> open = OfficePathIdentity.OpenRegularFileForRead;
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-binary-path-" + Guid.NewGuid().ToString("N"));
        string root = Path.Combine(directory, "root");
        Directory.CreateDirectory(root);
        try {
            string inside = Path.Combine(root, "inside.txt");
            string outside = Path.Combine(directory, "outside.txt");
            File.WriteAllText(inside, "inside");
            File.WriteAllText(outside, "outside");
            string physicalRoot = OfficePathIdentity.ResolvePhysicalPath(root);

            using (FileStream stream = open(inside, physicalRoot, 4096))
            using (var reader = new StreamReader(stream)) {
                Assert.Equal("inside", reader.ReadToEnd());
                Assert.False(stream.CanWrite);
            }
            Assert.Throws<InvalidDataException>(() => open(outside, physicalRoot, 4096));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PublishedZipScanSignatureRetainsEntryBoundAndSourcePosition() {
        Func<Stream, long, int, OfficeArchiveSafety.ZipCentralDirectoryScanResult> scan =
            OfficeArchiveSafety.ScanZipCentralDirectory;
        using var archiveBytes = new MemoryStream();
        using (var archive = new ZipArchive(archiveBytes, ZipArchiveMode.Create, leaveOpen: true)) {
            archive.CreateEntry("first.txt");
            archive.CreateEntry("second.txt");
        }
        using var source = new MemoryStream();
        source.Write(new byte[] { 1, 2, 3 }, 0, 3);
        archiveBytes.Position = 0;
        archiveBytes.CopyTo(source);
        source.Position = 3;

        var complete = scan(source, archiveBytes.Length, 2);
        Assert.True(complete.IsValid);
        Assert.False(complete.LimitExceeded);
        Assert.Equal(2L, complete.EntryCount);
        Assert.Equal(3L, source.Position);

        var limited = scan(source, archiveBytes.Length, 1);
        Assert.True(limited.IsValid && limited.LimitExceeded);
        Assert.Equal(3L, source.Position);
        Assert.Throws<ArgumentOutOfRangeException>(() => scan(source, archiveBytes.Length, -1));
    }

    [Fact]
    public void PublishedSnapshotSignatureRetainsInputBoundCancellationAndDisposal() {
        Func<string, long, CancellationToken, OfficeProvenanceFileSnapshot> capture =
            OfficeProvenanceFileSnapshot.Capture;
        string path = Path.Combine(Path.GetTempPath(), "officeimo-binary-snapshot-" + Guid.NewGuid().ToString("N") + ".txt");
        byte[] payload = Encoding.UTF8.GetBytes("snapshot input");
        File.WriteAllBytes(path, payload);
        try {
            string snapshotPath;
            using (OfficeProvenanceFileSnapshot snapshot = capture(path, payload.Length, CancellationToken.None)) {
                snapshotPath = snapshot.FilePath;
                Assert.NotEqual(path, snapshotPath);
                Assert.Equal(payload.LongLength, snapshot.Length);
                Assert.Equal(payload, File.ReadAllBytes(snapshotPath));
            }
            Assert.False(File.Exists(snapshotPath));
            Assert.Equal(payload, File.ReadAllBytes(path));
            Assert.Throws<InvalidDataException>(() => capture(path, payload.Length - 1, CancellationToken.None));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => capture(path, payload.Length, cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void PublishedAssessmentSignatureRetainsTextEvidenceAndCancellation() {
        Func<string, string, OfficeProvenanceReport, OfficeProvenanceAssessmentOptions?,
            IOfficeProvenanceVerifier?, IEnumerable<IOfficeProvenanceSignalDetector>?, CancellationToken,
            Encoding?, OfficeProvenanceAssessmentReport> assess = OfficeProvenanceAssessment.AssessSnapshotFile;
        string path = Path.Combine(Path.GetTempPath(), "officeimo-binary-assessment-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text\u200B", new UTF8Encoding(false));
        try {
            OfficeProvenanceReport structural = OfficeProvenanceInspector.InspectFile(path);
            OfficeProvenanceAssessmentReport report = assess(path, path, structural, null, null, null,
                CancellationToken.None, Encoding.UTF8);
            Assert.Same(structural, report.Structural);
            Assert.Equal(OfficeTextIntegrityFindingKind.ZeroWidthSpace, Assert.Single(report.TextIntegrity!.Findings).Kind);
            Assert.Null(report.Verification);
            Assert.Empty(report.ProviderSignals);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => assess(path, path, structural, null, null, null,
                cancellation.Token, Encoding.UTF8));
        } finally {
            File.Delete(path);
        }
    }
}
