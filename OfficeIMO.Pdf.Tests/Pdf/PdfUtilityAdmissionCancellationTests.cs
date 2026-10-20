using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfUtilityAdmissionCancellationTests {
    [Fact]
    public void RangeAndSplitRoutesHonorExactRemainingStreamLimitAndKeepCallerStreamsOpen() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 5];
        Buffer.BlockCopy(pdf, 0, prefixed, 5, pdf.Length);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
        Func<Stream, PdfLoadOptions, IReadOnlyList<byte[]>>[] routes = {
            (input, limits) => new[] { PdfPageExtractor.ExtractPages(input, new[] { 1 }, limits, CancellationToken.None) },
            (input, limits) => new[] { PdfPageExtractor.ExtractPageRange(input, 1, 1, limits) },
            (input, limits) => new[] { PdfPageExtractor.ExtractPageRanges(input, new[] { PdfPageRange.From(1, 1) }, limits) },
            (input, limits) => PdfPageExtractor.SplitPages(input, limits, CancellationToken.None),
            (input, limits) => PdfPageExtractor.SplitPageRanges(input, new[] { PdfPageRange.From(1, 1) }, limits)
        };

        foreach (bool seekable in new[] { true, false }) {
            foreach (var route in routes) {
                using var input = new ControlledInputStream(prefixed, 5, seekable);
                IReadOnlyList<byte[]> parts = route(input, options);
                Assert.NotEmpty(parts);
                Assert.All(parts, part => Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(part), StringComparison.Ordinal));
                Assert.True(input.CanRead);
                Assert.Equal(prefixed.Length, input.BytesConsumed);
            }
        }
    }

    [Fact]
    public void RangeAndSplitRoutesRejectCustomLimitsWithoutChangingOutputOrReadingOversizedSeekableInput() {
        byte[] pdf = CreatePdf();
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
        Action<Stream, Stream>[] routes = {
            (input, output) => PdfPageExtractor.ExtractPages(input, output, new[] { 1 }, options, CancellationToken.None),
            (input, output) => PdfPageExtractor.ExtractPageRange(input, output, 1, 1, options),
            (input, output) => PdfPageExtractor.ExtractPageRange(input, output, PdfPageRange.From(1, 1), options),
            (input, output) => PdfPageExtractor.ExtractPageRanges(input, output, new[] { PdfPageRange.From(1, 1) }, options),
            (input, _) => PdfPageExtractor.SplitPages(input, options, CancellationToken.None),
            (input, _) => PdfPageExtractor.SplitPageRanges(input, new[] { PdfPageRange.From(1, 1) }, options)
        };

        foreach (bool seekable in new[] { true, false }) {
            foreach (var route in routes) {
                using var input = new ControlledInputStream(pdf, 0, seekable);
                using var output = new MemoryStream();
                output.WriteByte(42);
                PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => route(input, output));
                Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
                Assert.Equal(pdf.Length - 1, error.Limit);
                Assert.Equal(new byte[] { 42 }, output.ToArray());
                Assert.True(input.CanRead);
                if (seekable) Assert.Equal(0, input.BytesConsumed);
            }
        }
    }

    [Fact]
    public void CancellationDuringRemainingInputLeavesEveryRangeDestinationUnchanged() {
        byte[] pdf = CreatePdf();
        string root = Path.Combine(Path.GetTempPath(), "officeimo-utility-cancellation-" + Guid.NewGuid().ToString("N"));
        string splitDirectory = Path.Combine(root, "parts");
        Action<Stream, Stream, CancellationToken>[] routes = {
            (input, output, token) => PdfPageExtractor.ExtractPageRanges(input, output, new[] { PdfPageRange.From(1, 1) }, null, token),
            (input, _, token) => PdfPageExtractor.SplitPages(input, splitDirectory, "part", null, token),
            (input, _, token) => PdfPageExtractor.SplitPageRanges(input, splitDirectory, "part", new[] { PdfPageRange.From(1, 1) }, null, token)
        };

        foreach (var route in routes) {
            using var cancellation = new CancellationTokenSource();
            using var input = new ControlledInputStream(pdf, 0, seekable: false, cancellation);
            using var output = new MemoryStream();
            output.WriteByte(42);
            Assert.Throws<OperationCanceledException>(() => route(input, output, cancellation.Token));
            Assert.InRange(input.BytesConsumed, 1, pdf.Length - 1);
            Assert.Equal(new byte[] { 42 }, output.ToArray());
            Assert.True(input.CanRead);
            Assert.False(Directory.Exists(root));
        }
    }

    [Fact]
    public void BoundedFileRoutesRejectBeforePublishingOrReplacingExistingOutput() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-utility-file-admission-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string outputPath = Path.Combine(root, "existing.pdf");
        string splitDirectory = Path.Combine(root, "parts");
        try {
            Directory.CreateDirectory(root);
            byte[] pdf = CreatePdf();
            File.WriteAllBytes(inputPath, pdf);
            File.WriteAllBytes(outputPath, new byte[] { 42 });
            var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
            var ranges = new[] { PdfPageRange.From(1, 1) };
            Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.ExtractPages(inputPath, outputPath, new[] { 1 }, options));
            Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.ExtractPageRange(inputPath, outputPath, 1, 1, options));
            Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.ExtractPageRanges(inputPath, outputPath, ranges, options));
            Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.SplitPages(inputPath, splitDirectory, options));
            Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.SplitPageRanges(inputPath, splitDirectory, ranges, options));
            Assert.Throws<PdfReadLimitException>(() => PdfMerger.MergeFiles(new[] { inputPath }, outputPath, options, CancellationToken.None));
            Assert.Throws<PdfReadLimitException>(() => PdfIncrementalUpdater.UpdateMetadata(inputPath, outputPath,
                title: "Rejected", readOptions: options, cancellationToken: CancellationToken.None));
            var importOptions = new PdfPageImportOptions { SourceReadOptions = options };
            Assert.Throws<PdfReadLimitException>(() => PdfPageImporter.AppendPages(importOptions, inputPath, inputPath, outputPath, PdfLoadOptions.Default, 1));
            Assert.Throws<PdfReadLimitException>(() => PdfPageImporter.PrependPages(importOptions, inputPath, inputPath, outputPath, PdfLoadOptions.Default, 1));
            Assert.Throws<PdfReadLimitException>(() => PdfPageImporter.InsertPages(importOptions, inputPath, inputPath, outputPath, 1, PdfLoadOptions.Default, 1));
            Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(outputPath));
            Assert.Equal(pdf, File.ReadAllBytes(inputPath));
            Assert.False(Directory.Exists(splitDirectory));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void DocumentRangeAndSplitCancellationPreservesSourceIncludingCancellationDuringRangeEnumeration() {
        byte[] pdf = CreatePdf();
        PdfDocument document = PdfDocument.Load(pdf);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Action[] routes = {
            () => document.Pages.Extract(PdfPageRange.From(1, 1), cancelled.Token),
            () => document.Pages.Extract(PdfPageSelection.From(1), cancelled.Token),
            () => document.Pages.Extract("1", cancelled.Token),
            () => document.Pages.Extract(PdfPageSelector.Parse("last"), cancelled.Token),
            () => document.Pages.Split(cancelled.Token),
            () => document.Pages.Split(1, cancelled.Token),
            () => document.Pages.Split(cancelled.Token, PdfPageSelection.From(1)),
            () => document.Pages.Split("1", cancelled.Token)
        };
        foreach (var route in routes) Assert.Throws<OperationCanceledException>(route);

        using var cancellation = new CancellationTokenSource();
        IEnumerable<PdfPageRange> Ranges() {
            yield return PdfPageRange.From(1, 1);
            cancellation.Cancel();
            yield return PdfPageRange.From(1, 1);
        }
        Assert.Throws<OperationCanceledException>(() => document.Pages.Split(Ranges(), cancellation.Token));
        Assert.Equal(pdf, document.ToBytes());
        Assert.Contains("Utility admission", document.Pages.Extract("1", CancellationToken.None).Read().Text, StringComparison.Ordinal);
        Assert.Single(document.Pages.Split(1, CancellationToken.None));
    }

    [Fact]
    public void ImportRejectsCallerSourceLimitsForBytesAndStreamsWithoutChangingTarget() {
        byte[] pdf = CreatePdf();
        var options = new PdfPageImportOptions {
            SourceReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }
        };
        Action<PdfDocument>[] byteRoutes = {
            target => target.Pages.Append(pdf, options),
            target => target.Pages.Prepend(pdf, options),
            target => target.Pages.Insert(1, pdf, options)
        };
        foreach (var route in byteRoutes) {
            PdfDocument target = PdfDocument.Load(pdf);
            Assert.Throws<PdfReadLimitException>(() => route(target));
            Assert.Equal(pdf, target.ToBytes());
        }
        foreach (bool seekable in new[] { true, false }) {
            PdfDocument target = PdfDocument.Load(pdf);
            using var input = new ControlledInputStream(pdf, 0, seekable);
            Assert.Throws<PdfReadLimitException>(() => target.Pages.Append(input, options));
            Assert.Equal(pdf, target.ToBytes());
            Assert.True(input.CanRead);
            if (seekable) Assert.Equal(0, input.BytesConsumed);
        }
    }

    [Fact]
    public void ImportConsumesRemainingNonseekableInputAndHonorsCancellationBeforeChangingTarget() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        Func<PdfDocument, Stream, PdfPageImportOptions, PdfDocument>[] routes = {
            (target, input, options) => target.Pages.Append(input, options),
            (target, input, options) => target.Pages.Prepend(input, options),
            (target, input, options) => target.Pages.Insert(1, input, options)
        };
        foreach (var route in routes) {
            PdfDocument target = PdfDocument.Load(pdf);
            var options = new PdfPageImportOptions {
                SourceReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } }
            };
            using (var input = new ControlledInputStream(prefixed, 3, seekable: false)) {
                PdfDocument result = route(target, input, options);
                Assert.Equal(2, result.Inspect().PageCount);
                Assert.True(input.CanRead);
                Assert.Equal(prefixed.Length, input.BytesConsumed);
            }
            Assert.Equal(pdf, target.ToBytes());

            using var cancellation = new CancellationTokenSource();
            using var cancelledInput = new ControlledInputStream(prefixed, 3, seekable: false, cancellation);
            options.CancellationToken = cancellation.Token;
            Assert.Throws<OperationCanceledException>(() => route(target, cancelledInput, options));
            Assert.Equal(pdf, target.ToBytes());
            Assert.True(cancelledInput.CanRead);
            Assert.InRange(cancelledInput.BytesConsumed, 4, prefixed.Length - 1);
        }
    }

    [Fact]
    public void OptimizationSanitizationAnnotationFlatteningAndPlanningShareRemainingInputAdmissionAndCancellation() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        Action<Stream, PdfLoadOptions, CancellationToken>[] routes = {
            (input, limits, token) => Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(
                PdfOptimizer.Optimize(input, new PdfOptimizationOptions { CancellationToken = token }, limits).Bytes), StringComparison.Ordinal),
            (input, limits, token) => Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(
                PdfSanitizer.Sanitize(input, new PdfSanitizationOptions { CancellationToken = token }, limits).ToBytes()), StringComparison.Ordinal),
            (input, limits, token) => Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(
                PdfAnnotationFlattener.FlattenVisualAnnotations(input, new PdfAnnotationFlattenOptions { CancellationToken = token }, limits)), StringComparison.Ordinal),
            (input, limits, token) => Assert.True(PdfMutationPlanner.Plan(input, PdfMutationOperation.ExtractPages,
                limits, cancellationToken: token).CanExecute)
        };
        foreach (var route in routes) {
            var exact = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
            using (var input = new ControlledInputStream(prefixed, 3, seekable: false)) {
                route(input, exact, CancellationToken.None);
                Assert.Equal(prefixed.Length, input.BytesConsumed);
                Assert.True(input.CanRead);
            }
            var tooSmall = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
            using (var input = new ControlledInputStream(prefixed, 3, seekable: true)) {
                Assert.Throws<PdfReadLimitException>(() => route(input, tooSmall, CancellationToken.None));
                Assert.Equal(3, input.BytesConsumed);
            }
            using var cancellation = new CancellationTokenSource();
            using var cancellingInput = new ControlledInputStream(prefixed, 3, seekable: false, cancellation);
            Assert.Throws<OperationCanceledException>(() => route(cancellingInput, exact, cancellation.Token));
            Assert.InRange(cancellingInput.BytesConsumed, 4, prefixed.Length - 1);
            Assert.True(cancellingInput.CanRead);
        }
    }

    [Fact]
    public void MergeAndIncrementalUtilitiesKeepRemainingPositionOwnershipAndOutputIntegrity() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        var exact = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
        using (var first = new ControlledInputStream(prefixed, 3, seekable: false))
        using (var second = new ControlledInputStream(prefixed, 3, seekable: false)) {
            byte[] merged = PdfMerger.Merge(new Stream[] { first, second }, exact, CancellationToken.None);
            Assert.Equal(2, PdfInspector.Inspect(merged).PageCount);
            Assert.Equal(prefixed.Length, first.BytesConsumed);
            Assert.Equal(prefixed.Length, second.BytesConsumed);
            Assert.True(first.CanRead);
            Assert.True(second.CanRead);
        }
        using (var input = new ControlledInputStream(prefixed, 3, seekable: false)) {
            byte[] revised = PdfIncrementalUpdater.UpdateMetadata(input, title: "Updated", readOptions: exact);
            Assert.Equal(pdf, revised.Take(pdf.Length).ToArray());
            Assert.Equal("Updated", PdfInspector.Inspect(revised).Metadata.Title);
            Assert.Equal(prefixed.Length, input.BytesConsumed);
            Assert.True(input.CanRead);
        }
        using var cancellation = new CancellationTokenSource();
        using var earlier = new ControlledInputStream(prefixed, 3, seekable: false);
        using var cancellingInput = new ControlledInputStream(prefixed, 3, seekable: false, cancellation);
        using var output = new MemoryStream();
        output.WriteByte(42);
        Assert.Throws<OperationCanceledException>(() => PdfMerger.Merge(new Stream[] { earlier, cancellingInput }, output, exact, cancellation.Token));
        Assert.Equal(new byte[] { 42 }, output.ToArray());
        Assert.Equal(prefixed.Length, earlier.BytesConsumed);
        Assert.InRange(cancellingInput.BytesConsumed, 4, prefixed.Length - 1);
        Assert.True(earlier.CanRead);
        Assert.True(cancellingInput.CanRead);
    }

    [Fact]
    public void MetadataStreamsHonorRemainingInputLimitsAndNeverPublishAfterReadCancellation() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        Action<Stream, Stream, PdfLoadOptions, CancellationToken>[] routes = {
            (input, output, limits, token) => PdfMetadataEditor.UpdateMetadata(input, output, limits, token, title: "Changed"),
            (input, output, limits, token) => PdfMetadataEditor.ReplaceMetadata(input, output, new PdfMetadata { Title = "Changed" }, limits, token),
            (input, output, limits, token) => PdfMetadataEditor.SynchronizeMetadata(input, output, limits, token, title: "Changed")
        };
        foreach (var route in routes) {
            var exact = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
            foreach (bool seekable in new[] { true, false }) {
                using var input = new ControlledInputStream(prefixed, 3, seekable);
                using var output = new MemoryStream();
                route(input, output, exact, CancellationToken.None);
                byte[] rewritten = output.ToArray();
                Assert.Equal("Changed", PdfInspector.Inspect(rewritten).Metadata.Title);
                Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(rewritten), StringComparison.Ordinal);
                Assert.Equal(prefixed.Length, input.BytesConsumed);
                Assert.True(input.CanRead);
            }
            using (var input = new ControlledInputStream(prefixed, 3, seekable: true))
            using (var output = new MemoryStream()) {
                output.WriteByte(42);
                Assert.Throws<PdfReadLimitException>(() => route(input, output,
                    new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }, CancellationToken.None));
                Assert.Equal(3, input.BytesConsumed);
                Assert.Equal(new byte[] { 42 }, output.ToArray());
            }
            using var cancellation = new CancellationTokenSource();
            using var cancellingInput = new ControlledInputStream(prefixed, 3, seekable: false, cancellation);
            using var untouchedOutput = new MemoryStream();
            untouchedOutput.WriteByte(42);
            Assert.Throws<OperationCanceledException>(() => route(cancellingInput, untouchedOutput, exact, cancellation.Token));
            Assert.Equal(new byte[] { 42 }, untouchedOutput.ToArray());
            Assert.True(cancellingInput.CanRead);
        }
    }

    [Fact]
    public void MetadataAndSecurityFileAdmissionLeavesExistingOutputAndSourceUntouched() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-metadata-security-admission-" + Guid.NewGuid().ToString("N"));
        string sourcePath = Path.Combine(root, "source.pdf");
        string outputPath = Path.Combine(root, "existing.pdf");
        byte[] pdf = CreatePdf();
        var encryption = new PdfStandardEncryptionOptions("reader") { OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128 };
        Action<PdfLoadOptions, CancellationToken>[] routes = {
            (limits, token) => PdfMetadataEditor.UpdateMetadata(sourcePath, outputPath, limits, token, title: "Changed"),
            (limits, token) => PdfMetadataEditor.ReplaceMetadata(sourcePath, outputPath, new PdfMetadata { Title = "Changed" }, limits, token),
            (limits, token) => PdfMetadataEditor.SynchronizeMetadata(sourcePath, outputPath, limits, token, title: "Changed"),
            (limits, token) => PdfSecurityEditor.Encrypt(sourcePath, outputPath, encryption, limits, token),
            (limits, token) => PdfSecurityEditor.Decrypt(sourcePath, outputPath, "owner", limits, token)
        };
        try {
            Directory.CreateDirectory(root);
            File.WriteAllBytes(sourcePath, pdf);
            File.WriteAllBytes(outputPath, new byte[] { 42 });
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            foreach (var route in routes) {
                Assert.Throws<PdfReadLimitException>(() => route(
                    new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }, CancellationToken.None));
                Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(outputPath));
                Assert.Throws<OperationCanceledException>(() => route(PdfLoadOptions.Default, cancellation.Token));
                Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(outputPath));
                Assert.Equal(pdf, File.ReadAllBytes(sourcePath));
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void DocumentMetadataAndSecurityFacadesPropagateCancellationAndPreserveTheirSource() {
        byte[] pdf = CreatePdf();
        PdfDocument document = PdfDocument.Load(pdf, new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } });
        var metadata = new PdfMetadata { Title = "Changed" };
        var encryption = new PdfStandardEncryptionOptions("reader") { OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128 };
        Func<CancellationToken, object>[] routes = {
            token => document.UpdateMetadata(token, title: "Changed"),
            token => document.ReplaceMetadata(metadata, token),
            token => document.SynchronizeMetadata(token, title: "Changed"),
            token => document.AppendMetadataRevision(token, title: "Changed"),
            token => document.UpdateMetadataResult(token, title: "Changed"),
            token => document.ReplaceMetadataResult(metadata, token),
            token => document.SynchronizeMetadataResult(token, title: "Changed"),
            token => document.AppendMetadataRevisionResult(token, title: "Changed"),
            token => document.Security.Encrypt(encryption, token),
            token => document.Security.EncryptResult(encryption, token),
            token => document.Security.Decrypt("owner", token),
            token => document.Security.DecryptResult("owner", token),
            token => document.Security.Reencrypt("owner", encryption, token),
            token => document.Security.ReencryptResult("owner", encryption, token)
        };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        foreach (var route in routes) {
            Assert.Throws<OperationCanceledException>(() => route(cancellation.Token));
            Assert.Equal(pdf, document.ToBytes());
        }
        Assert.Equal("Changed", document.UpdateMetadata(CancellationToken.None, title: "Changed").Inspect().Metadata.Title);
        Assert.Equal("Changed", document.SynchronizeMetadata(CancellationToken.None, title: "Changed").Inspect().Metadata.Title);
        byte[] revision = document.AppendMetadataRevision(CancellationToken.None, title: "Changed").ToBytes();
        Assert.Equal(pdf, revision.Take(pdf.Length).ToArray());
        PdfSecurityMutationResult protectedResult = document.Security.Encrypt(encryption, CancellationToken.None);
        PdfSecurityMutationResult unlocked = protectedResult.ToDocument().Security.Decrypt("owner", CancellationToken.None);
        Assert.Contains("Utility admission", PdfTextExtractor.ExtractAllText(unlocked.Pdf), StringComparison.Ordinal);
        Assert.Equal(pdf, document.ToBytes());
    }

    [Fact]
    public void ExplicitPolicyMergeCancelsDuringEnumerationAndValidatesEveryStreamBeforeReading() {
        byte[] pdf = CreatePdf();
        PdfDocument document = PdfDocument.Load(pdf);
        var options = new PdfMergeOptions();
        using var cancellation = new CancellationTokenSource();
        IEnumerable<byte[]> Inputs() { yield return pdf; cancellation.Cancel(); yield return pdf; }
        Assert.Throws<OperationCanceledException>(() => PdfDocument.MergeBytesResult(options, Inputs(), cancellation.Token));
        Func<object>[] cancelled = {
            () => PdfDocument.MergeBytes(options, new[] { pdf }, cancellation.Token),
            () => PdfDocument.MergeToBytes(options, new[] { pdf }, cancellation.Token),
            () => PdfDocument.Merge(options, new[] { document }, cancellation.Token),
            () => document.MergeWith(document, options, cancellation.Token),
            () => document.MergeWithResult(document, options, cancellation.Token)
        };
        foreach (var route in cancelled) Assert.Throws<OperationCanceledException>(() => route());
        Assert.Equal(pdf, document.ToBytes());
        Assert.Equal(2, PdfDocument.MergeBytesResult(options, new[] { pdf, pdf }, CancellationToken.None).RequireValue().Inspect().PageCount);

        using var input = new ControlledInputStream(pdf, 0, seekable: false);
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(new Stream[] { input, null! }));
        Assert.Equal(0, input.BytesConsumed);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void LegacyImportStreamsHonorSeparateSourceAndTargetAdmissionAndCancelBeforeOutput() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        var exact = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
        Action<PdfPageImportOptions, Stream, Stream, Stream, PdfLoadOptions>[] routes = {
            (options, target, source, output, limits) => PdfPageImporter.AppendPages(options, target, source, output, limits, 1),
            (options, target, source, output, limits) => PdfPageImporter.PrependPages(options, target, source, output, limits, 1),
            (options, target, source, output, limits) => PdfPageImporter.InsertPages(options, target, source, output, 1, limits, 1)
        };
        foreach (var route in routes) {
            foreach (bool seekable in new[] { true, false }) {
                using var target = new ControlledInputStream(prefixed, 3, seekable);
                using var source = new ControlledInputStream(prefixed, 3, seekable);
                using var output = new MemoryStream();
                route(new PdfPageImportOptions { SourceReadOptions = exact }, target, source, output, exact);
                Assert.Equal(2, PdfInspector.Inspect(output.ToArray()).PageCount);
                Assert.Equal(prefixed.Length, target.BytesConsumed);
                Assert.Equal(prefixed.Length, source.BytesConsumed);
                Assert.True(target.CanRead);
                Assert.True(source.CanRead);
            }
            using (var target = new ControlledInputStream(prefixed, 3, seekable: true))
            using (var source = new ControlledInputStream(prefixed, 3, seekable: true))
            using (var output = new MemoryStream()) {
                output.WriteByte(42);
                Assert.Throws<PdfReadLimitException>(() => route(new PdfPageImportOptions { SourceReadOptions = exact }, target, source, output,
                    new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }));
                Assert.Equal(3, target.BytesConsumed);
                Assert.Equal(3, source.BytesConsumed);
                Assert.Equal(new byte[] { 42 }, output.ToArray());
                Assert.Throws<PdfReadLimitException>(() => route(new PdfPageImportOptions {
                    SourceReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }
                }, target, source, output, exact));
                Assert.Equal(3, source.BytesConsumed);
                Assert.Equal(new byte[] { 42 }, output.ToArray());
            }
            using var cancellation = new CancellationTokenSource();
            using var uncancelledTarget = new ControlledInputStream(prefixed, 3, seekable: false);
            using var cancellingSource = new ControlledInputStream(prefixed, 3, seekable: false, cancellation);
            using var untouchedOutput = new MemoryStream();
            untouchedOutput.WriteByte(42);
            Assert.Throws<OperationCanceledException>(() => route(new PdfPageImportOptions {
                SourceReadOptions = exact, CancellationToken = cancellation.Token
            }, uncancelledTarget, cancellingSource, untouchedOutput, exact));
            Assert.Equal(new byte[] { 42 }, untouchedOutput.ToArray());
            Assert.True(uncancelledTarget.CanRead);
            Assert.True(cancellingSource.CanRead);
        }
    }

    [Fact]
    public void UtilityStreamWritersObserveCancellationRaisedByTheFinalWriteAndKeepCallerOwnership() {
        byte[] pdf = CreatePdf();
        byte[] original = (byte[])pdf.Clone();
        Action<Stream, CancellationToken>[] routes = {
            (output, token) => PdfPageExtractor.ExtractPages(pdf, output, new[] { 1 }, null, token),
            (output, token) => PdfPageExtractor.ExtractPageRange(pdf, output, PdfPageRange.From(1, 1), null, token),
            (output, token) => PdfMetadataEditor.UpdateMetadata(pdf, output, null, token, title: "Changed"),
            (output, token) => PdfMetadataEditor.ReplaceMetadata(pdf, output, new PdfMetadata { Title = "Changed" }, null, token),
            (output, token) => PdfMetadataEditor.SynchronizeMetadata(pdf, output, null, token, title: "Changed"),
            (output, token) => PdfPageImporter.AppendPages(new PdfPageImportOptions { CancellationToken = token }, pdf, pdf, output, null, 1),
            (output, token) => {
                using var first = new MemoryStream(pdf, writable: false);
                using var second = new MemoryStream(pdf, writable: false);
                PdfMerger.Merge(new Stream[] { first, second }, output, null, token);
            }
        };
        foreach (var route in routes) {
            using var cancellation = new CancellationTokenSource();
            using var output = new CancellingOutputStream(cancellation);
            output.WriteByte(42);
            Assert.Throws<OperationCanceledException>(() => route(output, cancellation.Token));
            Assert.True(output.CanWrite);
            Assert.True(output.Length > 1);
            Assert.Equal(42, output.ToArray()[0]);
            Assert.Equal(original, pdf);
        }
    }

    private static byte[] CreatePdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Utility admission"))
        .ToBytes();

    private sealed class CancellingOutputStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancellingOutputStream(CancellationTokenSource cancellation) => _cancellation = cancellation;
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }

    private sealed class ControlledInputStream : Stream {
        private readonly byte[] _bytes;
        private readonly bool _seekable;
        private readonly CancellationTokenSource? _cancelAfterRead;
        private int _position;
        private bool _disposed;
        internal ControlledInputStream(byte[] bytes, int position, bool seekable, CancellationTokenSource? cancelAfterRead = null) {
            _bytes = bytes;
            _position = position;
            _seekable = seekable;
            _cancelAfterRead = cancelAfterRead;
        }
        internal int BytesConsumed => _position;
        public override bool CanRead => !_disposed;
        public override bool CanSeek => _seekable;
        public override bool CanWrite => false;
        public override long Length => _seekable ? _bytes.Length : throw new NotSupportedException();
        public override long Position { get => _seekable ? _position : throw new NotSupportedException(); set => _position = checked((int)value); }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = Math.Min(Math.Min(count, 7), _bytes.Length - _position);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, read);
            _position += read;
            if (read > 0) _cancelAfterRead?.Cancel();
            return read;
        }
        protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
