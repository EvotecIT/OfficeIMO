using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderProjectionProbeEligibilityTests {
    private const string Password = "controlled-probe-fixture";
    private static readonly Lazy<byte[]> EncryptedWord = new(() => {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Eligible encrypted source");
        using var output = new MemoryStream(); document.Save(output,new OfficeIMO.Word.WordSaveOptions{Password=Password});
        return output.ToArray();
    });

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void IneligibleExtensionValidatorCannotPreventEligibleContentRouting(bool subset, bool pathInput) {
        int probes = 0;
        ReaderHandlerRegistration extension = Handler("extension", ReaderInputKind.Text, ".txt", ".allowed");
        extension.SupportsTextAndMetadata = subset;
        extension.TextAndMetadataExtensions = subset ? new[] { ".allowed" } : Array.Empty<string>();
        extension.ExtensionValidationProbeStream = (_, _, _, _) => { probes++; throw new ReaderResourceLimitException("unqualified-native-probe", 1); };
        ReaderHandlerRegistration fallback = Handler("detected", ReaderInputKind.Markdown, ".md");
        fallback.SupportsTextAndMetadata = true;
        var reader = new OfficeDocumentReaderBuilder().AddHandler(extension).AddHandler(fallback).Build();
        OfficeDocumentReadResult result = Read(reader, Encoding.UTF8.GetBytes("# Detected Markdown\n\nBody"), "payload.txt",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, DetectionMode = ReaderDetectionMode.PreferContent }, pathInput);
        Assert.Equal(ReaderInputKind.Markdown, result.Kind);
        Assert.Equal(0, probes);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", result.CapabilitiesUsed);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void IneligibleEncryptedProbeCannotPreventAnEligibleProbe(bool subset, bool pathInput) {
        int rejectedProbes = 0, eligibleProbes = 0;
        ReaderHandlerRegistration rejected = Handler("a-ineligible", ReaderInputKind.Excel, ".bad");
        rejected.SupportsTextAndMetadata = subset;
        rejected.TextAndMetadataExtensions = subset ? new[] { ".bad" } : Array.Empty<string>();
        rejected.ProbeStream = (_, _, _, _) => { rejectedProbes++; throw new ReaderResourceLimitException("unqualified-native-probe", 1); };
        ReaderHandlerRegistration eligible = Handler("z-eligible", ReaderInputKind.Word, ".docx");
        eligible.SupportsTextAndMetadata = true;
        eligible.ProbeStream = (_, _, _, _) => { eligibleProbes++; return true; };
        var reader = new OfficeDocumentReaderBuilder().AddHandler(rejected).AddHandler(eligible).Build();
        OfficeDocumentReadResult result = Read(reader, EncryptedWord.Value, "payload.bin",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, OpenPassword = Password }, pathInput);
        Assert.Equal(ReaderInputKind.Word, result.Kind);
        Assert.Equal(0, rejectedProbes);
        Assert.Equal(1, eligibleProbes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void QualifiedExtensionValidationAndFullProjectionRetainTheirCallbacks(bool pathInput) {
        int probes = 0;
        ReaderHandlerRegistration extension = Handler("extension", ReaderInputKind.Text, ".txt");
        extension.SupportsTextAndMetadata = true;
        extension.ExtensionValidationProbeStream = (_, _, _, _) => { probes++; return true; };
        ReaderHandlerRegistration fallback = Handler("detected", ReaderInputKind.Markdown, ".md"); fallback.SupportsTextAndMetadata = true;
        var reader = new OfficeDocumentReaderBuilder().AddHandler(extension).AddHandler(fallback).Build();
        byte[] bytes = Encoding.UTF8.GetBytes("# Text accepted by its native owner");
        foreach (ReaderProjection projection in new[] { ReaderProjection.TextAndMetadata, ReaderProjection.Full }) {
            Assert.Equal(ReaderInputKind.Text, Read(reader, bytes, "payload.txt", new ReaderOptions {
                Projection = projection, DetectionMode = ReaderDetectionMode.PreferContent
            }, pathInput).Kind);
        }
        Assert.Equal(2, probes);
        // Full still invokes a probe whose handler does not advertise lean support.
        extension.SupportsTextAndMetadata = false;
        reader = new OfficeDocumentReaderBuilder().AddHandler(extension).AddHandler(fallback).Build();
        Assert.Equal(ReaderInputKind.Text, Read(reader, bytes, "payload.txt", new ReaderOptions {
            DetectionMode = ReaderDetectionMode.PreferContent
        }, pathInput).Kind);
        Assert.Equal(3, probes);
    }

    [Theory]
    [InlineData(ReaderDetectionMode.ContentWhenUnknown, false)]
    [InlineData(ReaderDetectionMode.ContentWhenUnknown, true)]
    [InlineData(ReaderDetectionMode.PreferContent, false)]
    [InlineData(ReaderDetectionMode.PreferContent, true)]
    public void UnqualifiedEncryptedSourceIsRefusedBeforeNativePartAdmission(ReaderDetectionMode mode, bool pathInput) {
        var reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, DetectionMode = mode,
            OpenPassword = Password, OpenXmlMaxPartUncompressedBytes = 1
        };
        Assert.Throws<NotSupportedException>(() => Read(reader, EncryptedWord.Value, "payload.bin", options, pathInput));
        options.Projection = ReaderProjection.Full; options.OpenXmlMaxPartUncompressedBytes = null;
        Assert.Contains(Read(reader, EncryptedWord.Value, "payload.bin", options, pathInput).Chunks,
            chunk => chunk.Text!.Contains("Eligible encrypted source", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EligibleEncryptedProbeStillEnforcesNativeAdmission(bool pathInput) {
        var reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, DetectionMode = ReaderDetectionMode.PreferContent,
            OpenPassword = Password, OpenXmlMaxPartUncompressedBytes = 1
        };
        ReaderResourceLimitException exception = Assert.Throws<ReaderResourceLimitException>(() => Read(reader, EncryptedWord.Value, "payload.docx", options, pathInput));
        Assert.Equal(nameof(ReaderOptions.OpenXmlMaxPartUncompressedBytes), exception.LimitName);
        Assert.Equal(1, exception.Maximum);
        options.OpenXmlMaxPartUncompressedBytes = null;
        Assert.Contains(Read(reader, EncryptedWord.Value, "payload.docx", options, pathInput).Chunks,
            chunk => chunk.Text!.Contains("Eligible encrypted source", StringComparison.Ordinal));
    }

    private static ReaderHandlerRegistration Handler(string id, ReaderInputKind kind, params string[] extensions) => new() {
        Id = id, Kind = kind, Extensions = extensions,
        ReadDocumentPath = (path, _, _) => new OfficeDocumentReadResult { Kind = kind, Source = new OfficeDocumentSource { Path = path } },
        ReadDocumentStream = (_, name, _, _) => new OfficeDocumentReadResult { Kind = kind, Source = new OfficeDocumentSource { Path = name } }
    };

    private static OfficeDocumentReadResult Read(OfficeDocumentReader reader, byte[] bytes, string name, ReaderOptions options, bool pathInput) {
        if (!pathInput) return reader.ReadDocument(bytes, name, options);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(name));
        try { File.WriteAllBytes(path, bytes); return reader.ReadDocument(path, options); }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
