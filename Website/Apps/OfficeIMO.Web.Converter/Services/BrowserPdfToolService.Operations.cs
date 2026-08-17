using System.IO.Compression;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal sealed partial class BrowserPdfToolService {
    private static PdfToolExecution Inspect(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        PdfDocumentPreflight preflight = PdfDocument.Preflight(file.Bytes, BrowserPdfPolicy.CreateReadOptions());
        var details = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["canRead"] = preflight.CanRead.ToString(),
            ["canRewrite"] = preflight.CanRewrite.ToString(),
            ["canExtractText"] = preflight.CanExtractText.ToString(),
            ["canManipulatePages"] = preflight.CanManipulatePages.ToString(),
            ["encrypted"] = preflight.Probe.Security.HasEncryption.ToString(),
            ["signatureMarkers"] = preflight.Probe.HasSignatures.ToString(),
            ["readBlockers"] = preflight.ReadBlockers.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["rewriteBlockers"] = preflight.RewriteBlockers.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };
        int? pageCount = preflight.DocumentInfo?.PageCount;
        if (preflight.DocumentInfo is PdfDocumentInfo info) {
            details["pdfVersion"] = info.EffectiveVersion ?? "unknown";
            details["pages"] = info.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            details["forms"] = info.FormFieldCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            details["annotations"] = info.AnnotationCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            details["attachments"] = info.AttachmentCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            details["activeContent"] = info.HasActiveContent.ToString();
            details["tagged"] = info.HasTaggedContent.ToString();
        }

        var messages = new List<PdfToolMessage> {
            new("Read readiness", preflight.CanRead ? "OfficeIMO.Pdf can inspect this document." : "The document cannot be fully read with the current authentication context.", preflight.CanRead ? "ocx-dot--good" : "ocx-dot--warn"),
            new("Rewrite readiness", preflight.CanRewrite ? "Full-rewrite operations are available." : "One or more security or structure blockers prevent ordinary full rewrites.", preflight.CanRewrite ? "ocx-dot--good" : "ocx-dot--warn")
        };
        foreach (string diagnostic in preflight.Diagnostics.Take(12)) {
            messages.Add(new("Diagnostic", diagnostic, "ocx-dot--warn"));
        }

        string summary = pageCount.HasValue
            ? $"Inspected {pageCount.Value} PDF page{(pageCount.Value == 1 ? string.Empty : "s")}."
            : "Created a bounded PDF preflight report.";
        byte[] inspection = JsonSerializer.SerializeToUtf8Bytes(new {
            schemaVersion = 1,
            tool = request.Tool.Id,
            engine = "OfficeIMO.Pdf",
            browserLocal = true,
            source = new { fileName = file.Name, bytes = file.Bytes.LongLength },
            summary,
            details,
            messages = messages.Select(static message => new { title = message.Title, message = message.Message })
        }, new JsonSerializerOptions { WriteIndented = true });
        return new PdfToolExecution(
            new BrowserConversionArtifact(inspection, OutputName(file, "inspection", ".json"), "application/json"),
            summary,
            pageCount,
            messages,
            details,
            PreviewInBrowser: false);
    }

    private static PdfToolExecution Merge(PdfToolRequest request) {
        PdfDocument[] documents = request.Files.Select(static file => Open(file)).ToArray();
        PdfMergeResult merge = PdfDocument.MergeWithReport(new PdfMergeOptions(), documents);
        byte[] bytes = merge.ToBytes();
        const string fileName = "officeimo-merged.pdf";
        var details = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["sourceCount"] = request.Files.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["outputPages"] = merge.Report.OutputPageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["policyDecisions"] = merge.Report.Decisions.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["outputEncrypted"] = merge.Report.OutputHasEncryption.ToString(),
            ["outputSignatureMarkers"] = merge.Report.OutputHasSignatures.ToString()
        };
        return new PdfToolExecution(
            PdfArtifact(bytes, fileName),
            $"Merged {request.Files.Count} PDFs into {merge.Report.OutputPageCount} pages.",
            merge.Report.OutputPageCount,
            [new("Merge complete", "The selected PDFs were merged in the displayed order through one OfficeIMO.Pdf merge pass.", "ocx-dot--good")],
            details);
    }

    private static PdfToolExecution Split(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        PdfDocument source = Open(file);
        int sourcePages = source.Inspect().PageCount;
        int outputDocuments = (int)Math.Ceiling(sourcePages / (double)request.PagesPerDocument);
        if (outputDocuments > BrowserPdfPolicy.MaxSplitDocuments) {
            throw new InvalidDataException($"Split would create {outputDocuments} files; the browser limit is {BrowserPdfPolicy.MaxSplitDocuments}.");
        }
        long serializedBytes = 0;
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            for (int index = 0; index < outputDocuments; index++) {
                int firstPage = checked(index * request.PagesPerDocument + 1);
                int lastPage = Math.Min(sourcePages, checked(firstPage + request.PagesPerDocument - 1));
                PdfPageSelector selector = PdfPageSelector.Parse(firstPage == lastPage ? firstPage.ToString(System.Globalization.CultureInfo.InvariantCulture) : $"{firstPage}-{lastPage}");
                byte[] bytes = source.Pages.Extract(selector).ToBytes();
                serializedBytes = checked(serializedBytes + bytes.LongLength);
                if (serializedBytes > BrowserPdfPolicy.MaxSplitSerializedBytes) {
                    throw new InvalidDataException($"Split outputs exceed the browser serialization limit of {FormatBytes(BrowserPdfPolicy.MaxSplitSerializedBytes)}.");
                }
                ZipArchiveEntry entry = archive.CreateEntry($"{Path.GetFileNameWithoutExtension(file.Name)}.part-{index + 1:000}.pdf", CompressionLevel.Optimal);
                using Stream destination = entry.Open();
                destination.Write(bytes, 0, bytes.Length);
            }
        }
        if (buffer.Length > BrowserPdfPolicy.MaxSplitSerializedBytes) {
            throw new InvalidDataException($"The split archive exceeds the browser output limit of {FormatBytes(BrowserPdfPolicy.MaxSplitSerializedBytes)}.");
        }
        var artifact = new BrowserConversionArtifact(buffer.ToArray(), OutputName(file, "split", ".zip"), "application/zip");
        return new PdfToolExecution(
            artifact,
            $"Split {sourcePages} pages into {outputDocuments} PDF files.",
            sourcePages,
            [new("Split complete", $"The ZIP contains {outputDocuments} PDFs with up to {request.PagesPerDocument} pages each.", "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["sourcePages"] = sourcePages.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["pagesPerDocument"] = request.PagesPerDocument.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["outputDocuments"] = outputDocuments.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["serializedPdfBytes"] = serializedBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)
            },
            PreviewInBrowser: false);
    }

    private static PdfToolExecution TransformPages(
        PdfToolRequest request,
        string suffix,
        Func<PdfDocument, PdfPageSelector, PdfToolRequest, PdfDocument> transform) {
        SelectedDocument file = request.Files[0];
        PdfDocument source = Open(file);
        int sourcePages = source.Inspect().PageCount;
        PdfPageSelector selector = Selector(request);
        PdfDocument output = transform(source, selector, request);
        int outputPages = output.Inspect().PageCount;
        byte[] bytes = output.ToBytes();
        return new PdfToolExecution(
            PdfArtifact(bytes, OutputName(file, suffix)),
            $"Created a {outputPages}-page PDF from a {sourcePages}-page source.",
            outputPages,
            [new("Page operation complete", $"Selector '{selector.Expression}' was resolved against {sourcePages} source pages.", "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["selector"] = selector.Expression,
                ["sourcePages"] = sourcePages.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["outputPages"] = outputPages.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["rotationDegrees"] = request.Tool.Kind == PdfToolKind.Rotate ? request.RotationDegrees.ToString(System.Globalization.CultureInfo.InvariantCulture) : "0"
            });
    }

    private static PdfToolExecution Optimize(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        PdfOptimizationActionResult result = Open(file).Optimization.Apply(request.OptimizationProfile);
        int pages = result.ToDocument().Inspect().PageCount;
        string sizeSummary = result.SavedBytes > 0
            ? $"Saved {FormatBytes(result.SavedBytes)} without rasterizing pages."
            : "The original was retained because the candidate was not smaller.";
        return new PdfToolExecution(
            PdfArtifact(result.Bytes, OutputName(file, "optimized")),
            sizeSummary,
            pages,
            [new("Lossless optimization", $"Profile {result.RequestedProfile} applied {result.ActionCount} actions and recorded {result.SkippedActionCount} skipped opportunities.", "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["profile"] = result.RequestedProfile.ToString(),
                ["originalBytes"] = result.OriginalLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["candidateBytes"] = result.CandidateLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["returnedBytes"] = result.OptimizedLengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["returnedOriginal"] = result.ReturnedOriginal.ToString(),
                ["actionCount"] = result.ActionCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["linearized"] = result.CandidateLinearized.ToString()
            });
    }

    private static PdfToolExecution Protect(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        var encryption = new PdfStandardEncryptionOptions(request.UserPassword) {
            OwnerPassword = request.OwnerPassword,
            Algorithm = PdfStandardEncryptionAlgorithm.Aes256
        };
        PdfSecurityMutationResult result = Open(file).Security.Encrypt(encryption);
        return new PdfToolExecution(
            PdfArtifact(result.Pdf, OutputName(file, "protected")),
            "Created an AES-256 password-protected PDF.",
            result.ToDocument().Inspect().PageCount,
            [new("AES-256 protection", "The output was read back as encrypted and includes rewrite-preservation evidence in the operation report.", "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["algorithm"] = encryption.Algorithm.ToString(),
                ["outputEncrypted"] = result.IsEncrypted.ToString(),
                ["mutation"] = result.Kind.ToString(),
                ["preservationVerified"] = result.PreservationReport.IsPreserved.ToString(),
                ["preservationIssueCount"] = result.PreservationReport.Issues.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["preservationSummary"] = result.PreservationReport.Summary,
                ["sourcePages"] = result.PreservationReport.Original.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["outputPages"] = result.PreservationReport.Rewritten.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
            },
            PreviewInBrowser: false);
    }

    private static PdfToolExecution Unlock(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        PdfDocument document = Open(file, request.OwnerPassword);
        PdfSecurityMutationResult result = document.Security.Decrypt(request.OwnerPassword);
        return new PdfToolExecution(
            PdfArtifact(result.Pdf, OutputName(file, "unlocked")),
            "Created a separate PDF without Standard password security.",
            result.ToDocument().Inspect().PageCount,
            [new("Protection removed", "The source file was not changed; the downloaded copy was read back as unencrypted.", "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["outputEncrypted"] = result.IsEncrypted.ToString(),
                ["mutation"] = result.Kind.ToString()
            });
    }

    private static PdfToolExecution Redact(PdfToolRequest request) {
        SelectedDocument file = request.Files[0];
        PdfDocument source = Open(file);
        var search = new PdfRedactionSearchOptions { MatchCase = false };
        search.AddLiteral(request.RedactionText);
        PdfRedactionPlan plan = source.Redactions.Search(search);
        if (!plan.HasMatches) {
            throw new InvalidOperationException("No matching text was found; no PDF was produced.");
        }
        PdfDocument redacted = source.Redactions.Apply(plan);
        var verificationOptions = new PdfRedactionVerificationOptions { MatchCase = false };
        verificationOptions.RequireRemovedText(FindConcreteRedactionMarkers(plan, request.RedactionText));
        PdfRedactionVerificationReport verification = redacted.Redactions.Verify(verificationOptions);
        verification.ThrowIfFailed();
        byte[] bytes = redacted.ToBytes();
        int pages = redacted.Inspect().PageCount;
        return new PdfToolExecution(
            PdfArtifact(bytes, OutputName(file, "redacted")),
            $"Redacted {plan.Matches.Count} matching object{(plan.Matches.Count == 1 ? string.Empty : "s")} and verified the rewritten PDF.",
            pages,
            [new("Redaction verified", verification.Summary, "ocx-dot--good")],
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["searchMode"] = "literal-case-insensitive",
                ["areas"] = plan.Areas.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["matches"] = plan.Matches.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["verifiedMarkerVariants"] = verificationOptions.RemovedTextMarkers.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["verified"] = verification.IsVerified.ToString(),
                ["rawBytesChecked"] = verification.RawPdfBytesChecked.ToString(),
                ["encodedStringsChecked"] = verification.EncodedPdfStringsChecked.ToString(),
                ["decodedStreamsChecked"] = verification.DecodedPdfStreamsChecked.ToString()
            });
    }

    private static PdfToolExecution Compare(PdfToolRequest request) {
        SelectedDocument expected = request.Files[0];
        SelectedDocument actual = request.Files[1];
        var options = new PdfVisualComparisonOptions {
            MaxPages = MaxComparisonPages,
            MaxPixelsPerImage = 12_000_000,
            MaxTotalPixels = 50_000_000,
            MaxTotalOutputBytes = 64L * 1024L * 1024L
        };
        PdfVisualComparisonReport comparison = PdfVisualComparer.Compare(
            expected.Bytes,
            actual.Bytes,
            options: options,
            expectedReadOptions: BrowserPdfPolicy.CreateReadOptions(),
            actualReadOptions: BrowserPdfPolicy.CreateReadOptions());
        string html = comparison.ToHtmlGallery($"{expected.Name} compared with {actual.Name}");
        var artifact = new BrowserConversionArtifact(
            Encoding.UTF8.GetBytes(html),
            $"{Path.GetFileNameWithoutExtension(expected.Name)}-vs-{Path.GetFileNameWithoutExtension(actual.Name)}.html",
            "text/html;charset=utf-8");
        var messages = new List<PdfToolMessage> {
            new(comparison.IsMatch ? "Visual match" : "Differences found", comparison.IsMatch ? "Every compared page satisfied the exact comparison threshold." : "Open the gallery to review expected, actual, and highlighted difference images.", comparison.IsMatch ? "ocx-dot--good" : "ocx-dot--warn")
        };
        foreach (string difference in comparison.StructuralDifferences.Take(10)) messages.Add(new("Structural difference", difference, "ocx-dot--warn"));
        return new PdfToolExecution(
            artifact,
            comparison.IsMatch ? "The PDFs match at the configured exact visual threshold." : $"Compared {comparison.Pages.Count} pages and found visual or structural differences.",
            comparison.Pages.Count,
            messages,
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["isMatch"] = comparison.IsMatch.ToString(),
                ["comparedPages"] = comparison.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["structuralDifferences"] = comparison.StructuralDifferences.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["channelTolerance"] = options.ChannelTolerance.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["allowedDifferenceRatio"] = options.AllowedDifferenceRatio.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
    }

    private static string[] FindConcreteRedactionMarkers(PdfRedactionPlan plan, string requestedMarker) {
        var markers = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfRedactionMatch match in plan.Matches) {
            string? text = match.Text;
            if (string.IsNullOrEmpty(text)) continue;
            int start = 0;
            while (start <= text.Length - requestedMarker.Length) {
                int index = text.IndexOf(requestedMarker, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0) break;
                markers.Add(text.Substring(index, requestedMarker.Length));
                start = index + requestedMarker.Length;
            }
        }
        if (markers.Count == 0) markers.Add(requestedMarker);
        return markers.ToArray();
    }
}
