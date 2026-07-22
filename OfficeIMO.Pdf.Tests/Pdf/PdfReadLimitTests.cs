using System.Diagnostics;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReadLimitTests {
    [Fact]
    public void DefaultReadLimitsAreIndependentAcrossOptionInstances() {
        var first = new PdfReadOptions();
        var second = new PdfReadOptions();
        var setter = typeof(PdfReadLimits).GetProperty(nameof(PdfReadLimits.MaxInputBytes))!;

        setter.SetValue(first.Limits, 1L);

        Assert.Equal(1L, first.Limits.MaxInputBytes);
        Assert.Equal(512L * 1024L * 1024L, second.Limits.MaxInputBytes);
        Assert.Equal(512L * 1024L * 1024L, PdfReadOptions.Default.Limits.MaxInputBytes);
        Assert.NotSame(PdfReadOptions.Default, PdfReadOptions.Default);
        Assert.NotSame(PdfReadLimits.Default, PdfReadLimits.Default);
    }

    [Fact]
    public void ExternalSignatureCompletionAllowsItsOwnPreparedRevisionBeyondTheSourceBudget() {
        byte[] pdf = BuildPdf();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length }
        };

        PdfDocument source = PdfDocument.Open(pdf, readOptions);
        PdfExternalSignaturePreparation preparation = source.PrepareExternalSignature(
            new PdfExternalSignatureOptions {
                FieldName = "BudgetedSignature",
                ReservedSignatureContentsBytes = 512
            });

        Assert.True(preparation.PreparedPdf.LongLength > readOptions.Limits.MaxInputBytes);
        PdfDocument completed = preparation.Complete(new byte[] { 0x30, 0x01, 0x00 });

        Assert.True(completed.ToBytes().LongLength > readOptions.Limits.MaxInputBytes);
        Assert.Single(completed.Inspect().FormFields);
    }

    [Fact]
    public void OneShotExternalSignatureEnforcesStoredPageLimitBeforeCallingSigner() {
        byte[] twoPages = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .ToBytes();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPages = 1 }
        };
        var signer = new RecordingExternalSigner();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => PdfDocument.Open(twoPages, readOptions).SignExternal(signer));

        Assert.Equal(PdfReadLimitKind.Pages, exception.Kind);
        Assert.False(signer.WasCalled);
    }

    [Fact]
    public void OneShotExternalSignatureCompletionRetainsPolicyForItsOwnedOutput() {
        byte[] pdf = BuildPdf();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length }
        };

        PdfExternalSignatureCompletion completion = PdfDocument.Open(pdf, readOptions)
            .SignExternal(new RecordingExternalSigner());

        Assert.True(completion.Pdf.LongLength > readOptions.Limits.MaxInputBytes);
        Assert.Single(completion.ToDocument().Inspect().FormFields);
    }

    [Fact]
    public void OneShotExternalSignatureBoundsAndCancelsStreamInputBeforeCallingSigner() {
        byte[] pdf = BuildPdf();
        byte[] prefixed = new byte[pdf.Length + 5];
        Buffer.BlockCopy(pdf, 0, prefixed, 5, pdf.Length);
        var signer = new RecordingExternalSigner();
        using var stream = new MemoryStream(prefixed);
        stream.Position = 5;

        PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() =>
            PdfIncrementalUpdater.SignExternal(
                stream,
                signer,
                new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length - 1L }));
        Assert.Equal(PdfReadLimitKind.InputBytes, limit.Kind);
        Assert.Equal(5, stream.Position);
        Assert.False(signer.WasCalled);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfIncrementalUpdater.SignExternal(
                stream,
                signer,
                new PdfExternalSignatureOptions { CancellationToken = cancellation.Token }));
        Assert.False(signer.WasCalled);

        using var currentPositionStream = new MemoryStream(prefixed);
        currentPositionStream.Position = 5;
        var currentPositionSigner = new RecordingExternalSigner();

        PdfExternalSignatureCompletion completion = PdfIncrementalUpdater.SignExternal(
            currentPositionStream,
            currentPositionSigner,
            new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length });

        Assert.True(currentPositionSigner.WasCalled);
        Assert.Equal(currentPositionStream.Length, currentPositionStream.Position);
        Assert.True(completion.Pdf.Length > pdf.Length);
    }

    [Fact]
    public void ExternalSignaturePreparationBoundsByteStreamAndPathInputs() {
        byte[] pdf = BuildPdf();

        PdfReadLimitException bytesLimit = Assert.Throws<PdfReadLimitException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(
                pdf,
                new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length - 1L }));
        Assert.Equal(PdfReadLimitKind.InputBytes, bytesLimit.Kind);

        byte[] prefixed = new byte[pdf.Length + 5];
        Buffer.BlockCopy(pdf, 0, prefixed, 5, pdf.Length);
        using var stream = new MemoryStream(prefixed);
        stream.Position = 5;
        PdfReadLimitException streamLimit = Assert.Throws<PdfReadLimitException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(
                stream,
                new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length - 1L }));
        Assert.Equal(PdfReadLimitKind.InputBytes, streamLimit.Kind);
        Assert.Equal(5, stream.Position);

        using var successfulStream = new MemoryStream(prefixed);
        successfulStream.Position = 5;
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            successfulStream,
            new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length });
        Assert.True(preparation.PreparedPdf.Length > pdf.Length);
        Assert.Equal(successfulStream.Length, successfulStream.Position);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(
                pdf,
                new PdfExternalSignatureOptions { CancellationToken = cancellation.Token }));

        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(inputPath, pdf);
            PdfReadLimitException pathLimit = Assert.Throws<PdfReadLimitException>(() =>
                PdfIncrementalUpdater.PrepareExternalSignature(
                    inputPath,
                    outputPath,
                    new PdfExternalSignatureOptions { MaxInputBytes = pdf.Length - 1L }));
            Assert.Equal(PdfReadLimitKind.InputBytes, pathLimit.Kind);
            Assert.False(File.Exists(outputPath));
        } finally {
            File.Delete(inputPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void PageCompositionRejectsOversizedRemainingStreamBeforeReading() {
        using var stream = new LengthOnlyReadStream(
            PdfReadOptions.Default.Limits.MaxInputBytes + 1L);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageEditor.DeletePages(stream, 1));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.False(stream.WasRead);
    }

    [Fact]
    public void AnnotationResultAllowsItsOwnRevisionBeyondTheSourceBudget() {
        byte[] pdf = BuildPdf();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length }
        };

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(
            pdf,
            new PdfAnnotationCreateOptions {
                Subtype = "Text",
                Contents = "Budgeted annotation"
            },
            readOptions);

        Assert.True(result.Bytes.LongLength > readOptions.Limits.MaxInputBytes);
        Assert.Single(result.ToDocument().Inspect().GetAnnotationsBySubtype("Text"));
    }

    [Fact]
    public void MergeAllowsItsOwnedOutputBeyondThePrimarySourceBudget() {
        byte[] first = BuildPdf();
        byte[] second = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Second merge source"))
            .ToBytes();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = first.Length }
        };

        PdfDocument merged = PdfDocument.Merge(
            PdfDocument.Open(first, readOptions),
            PdfDocument.Open(second));

        Assert.True(merged.ToBytes().LongLength > readOptions.Limits.MaxInputBytes);
        Assert.Equal(2, merged.Inspect().PageCount);
    }

    [Fact]
    public void WebOptimizationAllowsItsOwnedCandidateBeyondTheSourceBudget() {
        byte[] pdf = BuildPdf();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length }
        };

        PdfOptimizationActionResult result = PdfDocument.Open(pdf, readOptions)
            .Optimize(PdfOptimizationProfile.Web);

        Assert.True(result.Bytes.LongLength > readOptions.Limits.MaxInputBytes);
        Assert.Single(result.ToDocument().Inspect().Pages);
    }

    [Fact]
    public void PersistedExternalSignatureCompletionEnforcesStoredPageLimitsBeforeMutation() {
        byte[] twoPages = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfDocument.Open(twoPages)
            .PrepareExternalSignature(new PdfExternalSignatureOptions {
                FieldName = "PersistedSignature",
                ReservedSignatureContentsBytes = 512
            });
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPages = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => PdfDocument.Open(preparation.PreparedPdf, readOptions)
                .CompleteExternalSignature(new byte[] { 0x30, 0x01, 0x00 }));

        Assert.Equal(PdfReadLimitKind.Pages, exception.Kind);
    }

    [Fact]
    public void InputByteBudgetStopsBeforeObjectScanning() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
        Assert.Equal(pdf.Length, exception.Actual);
    }

    [Fact]
    public void SeekableStreamBudgetStopsBeforeBuffering() {
        byte[] pdf = BuildPdf();
        using var stream = new MemoryStream(pdf);
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(stream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void PdfDocumentOpenAppliesTheSameInputBudgetToBytesPathAndStream() {
        byte[] pdf = BuildPdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        try {
            File.WriteAllBytes(path, pdf);
            using var stream = new MemoryStream(pdf);
            stream.Position = stream.Length;
            long originalPosition = stream.Position;

            PdfReadLimitException byteException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Open(pdf, options));
            PdfReadLimitException pathException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Open(path, options));
            PdfReadLimitException streamException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Open(stream, options));

            Assert.Equal(PdfReadLimitKind.InputBytes, byteException.Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes, pathException.Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes, streamException.Kind);
            Assert.Equal(originalPosition, stream.Position);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task PdfDocumentOpenAsyncAppliesTheInputBudgetAndRestoresSeekableStreams() {
        byte[] pdf = BuildPdf();
        using var stream = new MemoryStream(pdf);
        stream.Position = stream.Length;
        long originalPosition = stream.Position;
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(
            () => PdfDocument.OpenAsync(stream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(originalPosition, stream.Position);
    }

    [Fact]
    public void PdfDocumentOpenBoundsNonSeekableStreamsWithoutReadingUnboundedInput() {
        byte[] pdf = BuildPdf();
        using var stream = new ChunkedNonSeekableStream(pdf, maximumChunkSize: 3);
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => PdfDocument.Open(stream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.InRange(stream.BytesRead, 17, 19);
    }

    [Fact]
    public void PdfDocumentPreflightBoundsPathsAndStreamsBeforeUnboundedBuffering() {
        byte[] pdf = BuildPdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-preflight-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        var restrictive = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        try {
            File.WriteAllBytes(path, pdf);
            PdfReadLimitException pathException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Preflight(path, restrictive));

            using var seekable = new MemoryStream(pdf);
            PdfReadLimitException streamException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Preflight(seekable, restrictive));

            using var nonSeekable = new ChunkedNonSeekableStream(pdf, maximumChunkSize: 3);
            PdfReadLimitException nonSeekableException = Assert.Throws<PdfReadLimitException>(
                () => PdfDocument.Preflight(nonSeekable, restrictive));

            Assert.Equal(PdfReadLimitKind.InputBytes, pathException.Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes, streamException.Kind);
            Assert.Equal(0, seekable.Position);
            Assert.Equal(PdfReadLimitKind.InputBytes, nonSeekableException.Kind);
            Assert.InRange(nonSeekable.BytesRead, 17, 19);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PdfDocumentPreflightConsumesSeekableStreamsFromTheirCurrentPosition() {
        byte[] pdf = BuildPdf();
        byte[] prefixed = new byte[pdf.Length + 7];
        Buffer.BlockCopy(pdf, 0, prefixed, 7, pdf.Length);
        using var stream = new MemoryStream(prefixed);
        stream.Position = 7;
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length }
        };

        PdfDocumentPreflight preflight = PdfDocument.Preflight(stream, options);

        Assert.True(preflight.CanRead);
        Assert.Equal(stream.Length, stream.Position);
    }

    [Fact]
    public void IndirectObjectBudgetStopsExcessiveDeclarations() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.IndirectObjects, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RawStreamBudgetStopsAllocation() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxRawStreamBytes = 4 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.RawStreamBytes, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void FlateDecoderStopsWhileOutputExceedsBudget() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        byte[] encoded;
        using (var buffer = new MemoryStream()) {
            using (var compressor = new DeflateStream(buffer, CompressionLevel.Optimal, leaveOpen: true)) {
                compressor.Write(new byte[4096], 0, 4096);
            }

            encoded = buffer.ToArray();
        }

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => StreamDecoder.Decode(dictionary, encoded, maxOutputBytes: 64));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
    }

    [Fact]
    public void RunLengthAndLzwDecodersStopWhileOutputExceedsBudget() {
        var runLengthDictionary = new PdfDictionary();
        runLengthDictionary.Items["Filter"] = new PdfName("RunLengthDecode");
        var lzwDictionary = new PdfDictionary();
        lzwDictionary.Items["Filter"] = new PdfName("LZWDecode");
        byte[] lzw = PackNineBitCodes(
            new[] { 256 }.Concat(Enumerable.Repeat(65, 64)).Concat(new[] { 257 }));

        PdfReadLimitException runLengthException = Assert.Throws<PdfReadLimitException>(
            () => StreamDecoder.Decode(runLengthDictionary, new byte[] { 129, (byte)'A', 128 }, maxOutputBytes: 32));
        PdfReadLimitException lzwException = Assert.Throws<PdfReadLimitException>(
            () => StreamDecoder.Decode(lzwDictionary, lzw, maxOutputBytes: 32));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, runLengthException.Kind);
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, lzwException.Kind);
    }

    [Fact]
    public void AsciiDecodersStopBeforeMaterializingOutputBeyondBudget() {
        var hexDictionary = new PdfDictionary();
        hexDictionary.Items["Filter"] = new PdfName("ASCIIHexDecode");
        var ascii85Dictionary = new PdfDictionary();
        ascii85Dictionary.Items["Filter"] = new PdfName("ASCII85Decode");

        PdfReadLimitException hexException = Assert.Throws<PdfReadLimitException>(() =>
            StreamDecoder.Decode(hexDictionary, Enumerable.Repeat((byte)'A', 64).ToArray(), maxOutputBytes: 8));
        PdfReadLimitException ascii85Exception = Assert.Throws<PdfReadLimitException>(() =>
            StreamDecoder.Decode(ascii85Dictionary, Enumerable.Repeat((byte)'z', 4).ToArray(), maxOutputBytes: 8));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, hexException.Kind);
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, ascii85Exception.Kind);
        Assert.Equal(8, hexException.Limit);
        Assert.Equal(8, ascii85Exception.Limit);
    }

    [Fact]
    public void IndirectStreamLengthCannotBypassRawStreamBudget() {
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxRawStreamBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfSyntax.ParseObjects(BuildIndirectLengthBudgetPdf(), options));

        Assert.Equal(PdfReadLimitKind.RawStreamBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
        Assert.Equal(128, exception.Actual);
    }

    [Fact]
    public void XrefStreamUsesCallerDecodedStreamBudget() {
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfSyntax.ParseObjects(BuildCompressedXrefStreamPdf(), options));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
    }

    [Fact]
    public void ReviewedRedactionPlanUsesCallerReadBudgetDuringApply() {
        byte[] pdf = BuildPdf();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(pdf, new[] {
            new PdfRedactionArea(1, 0, 0, 20, 20, "reviewed")
        });
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfRedactionApplier.Apply(pdf, plan, readOptions: options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(pdf.Length - 1, exception.Limit);
    }

    [Fact]
    public void PageContentUsesCallerDecodedStreamBudget() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 8 }
        };
        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(8, exception.Limit);
    }

    [Fact]
    public void InvalidReadBudgetsAreRejectedExplicitly() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = 0 }
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfReadDocument.Open(pdf, options));
    }

    [Fact]
    public void ObjectCharacterAndTokenBudgetsFailWithoutSilentTruncation() {
        byte[] characterHeavy = BuildObjectPdf("<< /LongValue (abcdefghijklmnopqrstuvwxyz) >>");
        byte[] tokenHeavy = BuildObjectPdf("[/A 1 /B 2 /C 3 /D 4]");
        var characterOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectCharacters = 24 }
        };
        var tokenOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTokensPerObject = 4 }
        };

        PdfReadLimitException characterException = Assert.Throws<PdfReadLimitException>(
            () => PdfSyntax.ParseObjects(characterHeavy, characterOptions));
        PdfReadLimitException tokenException = Assert.Throws<PdfReadLimitException>(
            () => PdfSyntax.ParseObjects(tokenHeavy, tokenOptions));

        Assert.Equal(PdfReadLimitKind.ObjectCharacters, characterException.Kind);
        Assert.Equal(PdfReadLimitKind.ObjectTokens, tokenException.Kind);
    }

    [Fact]
    public void DeclaredStreamDictionaryIsBoundedBeforeSubstringParsing() {
        byte[] pdf = BuildObjectPdf(
            "<< /Length 0 /Padding (" + new string('x', 256) + ") >>\nstream\n\nendstream");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectCharacters = 64 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => PdfSyntax.ParseObjects(pdf, options));

        Assert.Equal(PdfReadLimitKind.ObjectCharacters, exception.Kind);
    }

    [Fact]
    public void NestedObjectBudgetStopsRecursiveArrayParsing() {
        byte[] pdf = BuildObjectPdf("[[[[1]]]]");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectNestingDepth = 2 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfSyntax.ParseObjects(pdf, options));

        Assert.Equal(PdfReadLimitKind.ObjectNestingDepth, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RevisionBudgetStopsIncrementalMarkerScanning() {
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(
            System.Text.Encoding.ASCII.GetString(BuildPdf()) + "\nstartxref\n0\n%%EOF\n");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxRevisions = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.Revisions, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void PageAndPageTreeBudgetsStopTraversal() {
        byte[] twoPages = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"))
            .ToBytes();
        byte[] nestedTree = BuildNestedPageTreePdf();

        PdfReadLimitException pageException = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(
            twoPages,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxPages = 1 } }));
        PdfReadLimitException nodeException = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(
            nestedTree,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxPageTreeNodes = 1 } }));
        PdfReadLimitException depthException = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(
            nestedTree,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxPageTreeDepth = 1 } }));

        Assert.Equal(PdfReadLimitKind.Pages, pageException.Kind);
        Assert.Equal(PdfReadLimitKind.PageTreeNodes, nodeException.Kind);
        Assert.Equal(PdfReadLimitKind.PageTreeDepth, depthException.Kind);
    }

    [Fact]
    public void FormFieldCountAndDepthBudgetsStopFieldTreeTraversal() {
        byte[] twoFields = PdfDocument.Create()
            .TextField("First", width: 100, height: 20)
            .TextField("Second", width: 100, height: 20)
            .ToBytes();
        byte[] nestedFields = BuildNestedFormFieldPdf();

        PdfReadLimitException countException = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(
            twoFields,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxFormFields = 1 } }));
        PdfReadLimitException depthException = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(
            nestedFields,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxFormFieldDepth = 1 } }));

        Assert.Equal(PdfReadLimitKind.FormFields, countException.Kind);
        Assert.Equal(PdfReadLimitKind.FormFieldDepth, depthException.Kind);
    }

    [Fact]
    public void AnnotationAndContentOperationBudgetsStopPageParsing() {
        byte[] annotations = BuildAnnotatedPagePdf();
        byte[] content = BuildPdf();
        PdfReadDocument annotationDocument = PdfReadDocument.Open(
            annotations,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxAnnotationsPerPage = 1 } });
        PdfReadDocument contentDocument = PdfReadDocument.Open(
            content,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxContentOperations = 1 } });
        PdfReadDocument operandDocument = PdfReadDocument.Open(
            content,
            new PdfReadOptions { Limits = new PdfReadLimits { MaxContentOperands = 1 } });

        PdfReadLimitException annotationException = Assert.Throws<PdfReadLimitException>(() => annotationDocument.Pages[0].GetAnnotations());
        PdfReadLimitException annotationDrawingException = Assert.Throws<PdfReadLimitException>(() => annotationDocument.Pages[0].ToDrawing());
        PdfReadLimitException contentException = Assert.Throws<PdfReadLimitException>(() => contentDocument.Pages[0].ExtractText());
        PdfReadLimitException drawingException = Assert.Throws<PdfReadLimitException>(() => contentDocument.Pages[0].ToDrawing());
        PdfReadLimitException operandException = Assert.Throws<PdfReadLimitException>(() => operandDocument.Pages[0].ExtractText());
        PdfReadLimitException drawingOperandException = Assert.Throws<PdfReadLimitException>(() => operandDocument.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.AnnotationsPerPage, annotationException.Kind);
        Assert.Equal(PdfReadLimitKind.AnnotationsPerPage, annotationDrawingException.Kind);
        Assert.Equal(PdfReadLimitKind.ContentOperations, contentException.Kind);
        Assert.Equal(PdfReadLimitKind.ContentOperations, drawingException.Kind);
        Assert.Equal(PdfReadLimitKind.ContentOperands, operandException.Kind);
        Assert.Equal(PdfReadLimitKind.ContentOperands, drawingOperandException.Kind);
    }

    [Fact]
    public void ContentNestingBudgetStopsDeepFormXObjects() {
        PdfReadDocument document = PdfReadDocument.Open(
            BuildNestedFormXObjectPdf(),
            new PdfReadOptions { Limits = new PdfReadLimits { MaxContentNestingDepth = 1 } });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void DeterministicSubsystemFuzzingRemainsWithinDeclaredBudgets() {
        byte[][] seeds = {
            BuildPdf(),
            BuildNestedPageTreePdf(),
            BuildNestedFormFieldPdf(),
            BuildAnnotatedPagePdf(),
            BuildNestedFormXObjectPdf()
        };
        var random = new Random(0x50F1);
        var timer = Stopwatch.StartNew();
        var options = new PdfReadOptions {
            ParsingMode = PdfParsingMode.Strict,
            Limits = new PdfReadLimits {
                MaxInputBytes = 1024 * 1024,
                MaxIndirectObjects = 1_000,
                MaxRawStreamBytes = 256 * 1024,
                MaxDecodedStreamBytes = 256 * 1024,
                MaxObjectCharacters = 64 * 1024,
                MaxTokensPerObject = 8_000,
                MaxObjectNestingDepth = 32,
                MaxObjectParsingTime = TimeSpan.FromMilliseconds(500),
                MaxRevisions = 32,
                MaxPageTreeNodes = 256,
                MaxPageTreeDepth = 32,
                MaxPages = 128,
                MaxFormFields = 128,
                MaxFormFieldDepth = 32,
                MaxAnnotationsPerPage = 128,
                MaxContentOperations = 2_000,
                MaxContentOperands = 4_000,
                MaxContentNestingDepth = 16
            }
        };

        for (int caseNumber = 0; caseNumber < 64; caseNumber++) {
            byte[] seed = seeds[caseNumber % seeds.Length];
            int length = random.Next(1, seed.Length + 33);
            var candidate = new byte[length];
            Buffer.BlockCopy(seed, 0, candidate, 0, Math.Min(seed.Length, candidate.Length));
            for (int mutation = 0; mutation < 12; mutation++) {
                candidate[random.Next(candidate.Length)] = (byte)random.Next(256);
            }

            ExerciseCandidate(candidate, options);
        }

        string[] filters = { "FlateDecode", "RunLengthDecode", "LZWDecode" };
        for (int caseNumber = 0; caseNumber < 32; caseNumber++) {
            var encoded = new byte[random.Next(1, 129)];
            random.NextBytes(encoded);
            var dictionary = new PdfDictionary();
            dictionary.Items["Filter"] = new PdfName(filters[caseNumber % filters.Length]);
            try {
                _ = StreamDecoder.Decode(dictionary, encoded, maxOutputBytes: 256);
            } catch (Exception exception) when (IsExpectedMalformedInputException(exception)) {
                // Decoder failures are expected for random payloads; resource use remains bounded.
            }
        }

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), "Subsystem fuzz pass exceeded the test budget: " + timer.Elapsed + ".");
    }

    [Fact]
    public void DeterministicHostileInputMutationsRemainBounded() {
        byte[] source = BuildPdf();
        var random = new Random(0x2062);
        var timer = Stopwatch.StartNew();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = 2 * 1024 * 1024,
                MaxIndirectObjects = 2_000,
                MaxRawStreamBytes = 512 * 1024,
                MaxObjectParsingTime = TimeSpan.FromSeconds(1)
            }
        };

        for (int caseNumber = 0; caseNumber < 32; caseNumber++) {
            int length = random.Next(1, source.Length + 65);
            var candidate = new byte[length];
            Buffer.BlockCopy(source, 0, candidate, 0, Math.Min(source.Length, candidate.Length));
            for (int mutation = 0; mutation < 8; mutation++) {
                candidate[random.Next(candidate.Length)] = (byte)random.Next(256);
            }

            try {
                _ = PdfReadDocument.Open(candidate, options);
            } catch (Exception exception) when (
                exception is ArgumentException ||
                exception is FormatException ||
                exception is InvalidOperationException ||
                exception is IOException) {
                // Malformed candidates may fail, but must stay within the declared parser contract.
            }
        }

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), "Hostile-input parser pass exceeded the test budget: " + timer.Elapsed + ".");
    }

    private static byte[] BuildPdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Bounded parser source"))
        .ToBytes();

    private static byte[] BuildObjectPdf(string body) => System.Text.Encoding.ASCII.GetBytes(
        "%PDF-1.7\n1 0 obj\n" + body + "\nendobj\ntrailer\n<< /Root 1 0 R /Size 2 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildIndirectLengthBudgetPdf() => System.Text.Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>\nendobj\n" +
        "4 0 obj\n<< /Length 6 0 R >>\nstream\nABCD\nendstream\nendobj\n" +
        "%" + new string('P', 192) + "\n" +
        "6 0 obj\n128\nendobj\ntrailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildCompressedXrefStreamPdf() {
        byte[] decoded = new byte[70];
        byte[] encoded;
        using (var compressed = new MemoryStream()) {
            using (var compressor = new DeflateStream(compressed, CompressionLevel.Optimal, leaveOpen: true)) {
                compressor.Write(decoded, 0, decoded.Length);
            }

            encoded = compressed.ToArray();
        }

        using var output = new MemoryStream();
        void Write(string value) {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        Write("%PDF-1.5\n");
        Write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>\nendobj\n");
        int xrefOffset = checked((int)output.Position);
        Write("5 0 obj\n<< /Type /XRef /Size 10 /Root 1 0 R /W [1 4 2] /Index [0 10] /Filter /FlateDecode /Length " + encoded.Length + " >>\nstream\n");
        output.Write(encoded, 0, encoded.Length);
        Write("\nendstream\nendobj\nstartxref\n" + xrefOffset + "\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildNestedPageTreePdf() => BuildPdfObjects(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Pages /Parent 2 0 R /Count 1 /Kids [4 0 R] >>",
        "<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 200] /Contents 5 0 R >>",
        "<< /Length 0 >>\nstream\n\nendstream");

    private static byte[] BuildNestedFormFieldPdf() => BuildPdfObjects(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< >>",
        "<< /Fields [7 0 R] >>",
        "<< /T (Parent) /Kids [8 0 R] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /T (Child) /FT /Tx /Rect [10 10 100 30] >>");

    private static byte[] BuildAnnotatedPagePdf() => BuildPdfObjects(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Type /Annot /Subtype /Text /Rect [10 10 20 20] /Contents (First) >>",
        "<< /Type /Annot /Subtype /Text /Rect [30 30 40 40] /Contents (Second) >>");

    private static byte[] BuildNestedFormXObjectPdf() => BuildPdfObjects(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 4 0 R >>",
        "<< /Length 7 >>\nstream\n/Fm1 Do\nendstream",
        "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Fm2 6 0 R >> >> /Length 7 >>\nstream\n/Fm2 Do\nendstream",
        "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Length 0 >>\nstream\n\nendstream");

    private static void ExerciseCandidate(byte[] candidate, PdfReadOptions options) {
        try {
            _ = PdfSyntax.ParseObjects(candidate, options);
            PdfReadDocument document = PdfReadDocument.Open(candidate, options);
            for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
                PdfReadPage page = document.Pages[pageIndex];
                _ = page.GetTextSpans();
                _ = page.GetAnnotations();
                _ = page.GetImagePlacements();
            }
        } catch (Exception exception) when (IsExpectedMalformedInputException(exception)) {
            // Strict parsing rejects malformed mutations; typed limits stop hostile resource growth.
        }
    }

    private static bool IsExpectedMalformedInputException(Exception exception) =>
        exception is ArgumentException ||
        exception is FormatException ||
        exception is InvalidOperationException ||
        exception is IOException ||
        exception is System.Text.RegularExpressions.RegexMatchTimeoutException;

    private static byte[] BuildPdfObjects(params string[] bodies) {
        var builder = new System.Text.StringBuilder("%PDF-1.7\n");
        for (int i = 0; i < bodies.Length; i++) {
            builder.Append(i + 1).Append(" 0 obj\n").Append(bodies[i]).Append("\nendobj\n");
        }

        builder.Append("trailer\n<< /Root 1 0 R /Size ")
            .Append(bodies.Length + 1)
            .Append(" >>\nstartxref\n0\n%%EOF\n");
        return System.Text.Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static byte[] PackNineBitCodes(IEnumerable<int> codes) {
        var bits = new List<int>();
        foreach (int code in codes) {
            for (int bit = 8; bit >= 0; bit--) {
                bits.Add((code >> bit) & 1);
            }
        }

        var bytes = new byte[(bits.Count + 7) / 8];
        for (int i = 0; i < bits.Count; i++) {
            bytes[i / 8] |= (byte)(bits[i] << (7 - (i % 8)));
        }

        return bytes;
    }

    private sealed class ChunkedNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private readonly int _maximumChunkSize;
        private int _position;

        internal ChunkedNonSeekableStream(byte[] bytes, int maximumChunkSize) {
            _bytes = bytes;
            _maximumChunkSize = maximumChunkSize;
        }

        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int available = _bytes.Length - _position;
            if (available <= 0) {
                return 0;
            }

            int read = Math.Min(Math.Min(count, _maximumChunkSize), available);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, read);
            _position += read;
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class LengthOnlyReadStream : Stream {
        private long _position;

        internal LengthOnlyReadStream(long length) {
            Length = length;
        }

        internal bool WasRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get; }
        public override long Position {
            get => _position;
            set => _position = value;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            WasRead = true;
            return 0;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) {
            _position = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            return _position;
        }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class RecordingExternalSigner : IPdfExternalSigner {
        public string Name => "Recording test signer";

        internal bool WasCalled { get; private set; }

        public byte[] Sign(PdfExternalSignatureRequest request) {
            WasCalled = true;
            return new byte[] { 0x30, 0x01, 0x00 };
        }
    }
}
