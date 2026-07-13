using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>
/// Recognized text returned by an external OCR provider for a candidate emitted by an OfficeIMO read result.
/// </summary>
public sealed class OfficeDocumentOcrTextResult {
    /// <summary>
    /// Identifier of the <see cref="OfficeDocumentOcrCandidate"/> this OCR result enriches.
    /// </summary>
    public string CandidateId { get; set; } = string.Empty;

    /// <summary>
    /// Plain recognized text for the candidate.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional confidence reported by the external OCR provider.
    /// </summary>
    public double? Confidence { get; set; }

    /// <summary>
    /// Optional BCP-47 language tag for the recognized text.
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    /// Optional provider or service name that produced the recognized text.
    /// </summary>
    public string? Provider { get; set; }

    /// <summary>
    /// Optional provider model or engine version.
    /// </summary>
    public string? Model { get; set; }
}

/// <summary>
/// Options controlling how OCR text is merged into a document read result.
/// </summary>
public sealed class OfficeDocumentOcrEnrichmentOptions {
    /// <summary>
    /// Removes candidates that received OCR text from the enriched result's pending candidate list.
    /// </summary>
    public bool RemoveResolvedCandidates { get; set; } = true;

    /// <summary>
    /// Removes matching <c>ocr-needed</c> diagnostics for candidates that received OCR text.
    /// </summary>
    public bool RemoveResolvedOcrNeededDiagnostics { get; set; } = true;

    /// <summary>
    /// Appends recognized OCR text to the result-level Markdown payload.
    /// </summary>
    public bool AppendRecognizedTextToMarkdown { get; set; } = true;

    /// <summary>
    /// Logical block kind assigned to OCR-enriched text blocks.
    /// </summary>
    public string BlockKind { get; set; } = "ocr-text";
}

/// <summary>
/// Result returned after applying external OCR text to a document read result.
/// </summary>
public sealed class OfficeDocumentOcrEnrichmentResult {
    /// <summary>
    /// Enriched document read result.
    /// </summary>
    public OfficeDocumentReadResult Document { get; set; } = new OfficeDocumentReadResult();

    /// <summary>
    /// Observable counters describing the OCR merge operation.
    /// </summary>
    public OfficeDocumentOcrEnrichmentReport Report { get; set; } = new OfficeDocumentOcrEnrichmentReport();
}

/// <summary>
/// Observable counters describing an OCR enrichment operation.
/// </summary>
public sealed class OfficeDocumentOcrEnrichmentReport {
    /// <summary>Total number of OCR candidates in the source read result.</summary>
    public int CandidateCount { get; set; }

    /// <summary>Number of external OCR results supplied by the caller.</summary>
    public int ResultCount { get; set; }

    /// <summary>Number of OCR results applied to matching candidates.</summary>
    public int AppliedResultCount { get; set; }

    /// <summary>Number of OCR candidates that still do not have recognized text.</summary>
    public int UnresolvedCandidateCount { get; set; }

    /// <summary>Number of supplied OCR results that did not match a candidate.</summary>
    public int UnmatchedResultCount { get; set; }

    /// <summary>Number of OCR text blocks added to the enriched result.</summary>
    public int EnrichedBlockCount { get; set; }

    /// <summary>Number of OCR text chunks added to the enriched result.</summary>
    public int EnrichedChunkCount { get; set; }

    /// <summary>Candidate identifiers that were enriched.</summary>
    public IReadOnlyList<string> AppliedCandidateIds { get; set; } = Array.Empty<string>();

    /// <summary>Supplied OCR result candidate identifiers that did not match a pending candidate.</summary>
    public IReadOnlyList<string> UnmatchedCandidateIds { get; set; } = Array.Empty<string>();
}

/// <summary>
/// Helpers for merging external OCR provider output into OfficeIMO read results.
/// </summary>
public static class OfficeDocumentOcrEnrichmentExtensions {
    /// <summary>
    /// Applies recognized OCR text to a document read result without requiring the core reader to run an OCR engine.
    /// </summary>
    /// <param name="result">Source read result containing OCR candidates.</param>
    /// <param name="ocrResults">Recognized OCR text keyed by candidate identifier.</param>
    /// <param name="options">Optional enrichment behavior.</param>
    /// <returns>An enriched read result plus observable merge counters.</returns>
    public static OfficeDocumentOcrEnrichmentResult ApplyOcrResults(this OfficeDocumentReadResult result, IEnumerable<OfficeDocumentOcrTextResult> ocrResults, OfficeDocumentOcrEnrichmentOptions? options = null) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (ocrResults == null) throw new ArgumentNullException(nameof(ocrResults));

        OfficeDocumentOcrEnrichmentOptions effectiveOptions = options ?? new OfficeDocumentOcrEnrichmentOptions();
        OfficeDocumentOcrTextResult[] suppliedResults = ocrResults
            .Where(static item => item != null)
            .Where(static item => !string.IsNullOrWhiteSpace(item.CandidateId) && !string.IsNullOrWhiteSpace(item.Text))
            .ToArray();
        Dictionary<string, OfficeDocumentOcrTextResult> suppliedByCandidate = suppliedResults
            .GroupBy(static item => item.CandidateId, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key, static group => group.Last(), StringComparer.Ordinal);
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates = result.OcrCandidates ?? Array.Empty<OfficeDocumentOcrCandidate>();
        var applied = new List<AppliedOcrText>();
        var unresolved = new List<OfficeDocumentOcrCandidate>();

        for (int i = 0; i < candidates.Count; i++) {
            OfficeDocumentOcrCandidate candidate = candidates[i];
            if (!string.IsNullOrWhiteSpace(candidate.Id) && suppliedByCandidate.TryGetValue(candidate.Id, out OfficeDocumentOcrTextResult? recognized)) {
                applied.Add(new AppliedOcrText(candidate, recognized, applied.Count));
            } else {
                unresolved.Add(candidate);
            }
        }

        HashSet<string> appliedIds = new HashSet<string>(applied.Select(static item => item.Candidate.Id), StringComparer.Ordinal);
        string[] unmatchedIds = suppliedByCandidate.Keys
            .Where(candidateId => !appliedIds.Contains(candidateId))
            .OrderBy(static candidateId => candidateId, StringComparer.Ordinal)
            .ToArray();
        OfficeDocumentBlock[] enrichedBlocks = applied.Select(item => BuildOcrBlock(result.Source, item, effectiveOptions)).ToArray();
        ReaderChunk[] enrichedChunks = applied.Select(item => BuildOcrChunk(result, item)).ToArray();
        OfficeDocumentReadResult enriched = new OfficeDocumentReadResult {
            SchemaId = string.IsNullOrWhiteSpace(result.SchemaId) ? OfficeDocumentReadResultSchema.Id : result.SchemaId,
            SchemaVersion = result.SchemaVersion == 0 ? OfficeDocumentReadResultSchema.CurrentVersion : result.SchemaVersion,
            Kind = result.Kind,
            Source = result.Source ?? new OfficeDocumentSource(),
            CapabilitiesUsed = AppendCapability(result.CapabilitiesUsed, "officeimo.reader.ocr-enrichment"),
            Markdown = BuildMarkdown(result.Markdown, applied, effectiveOptions),
            Html = result.Html,
            Json = result.Json,
            Chunks = Append(result.Chunks, enrichedChunks),
            Metadata = Append(result.Metadata, BuildOcrMetadata(applied, unresolved.Count, unmatchedIds.Length)),
            Pages = BuildPages(result.Pages, enrichedBlocks, effectiveOptions.RemoveResolvedCandidates ? unresolved : candidates),
            Blocks = Append(result.Blocks, enrichedBlocks),
            Tables = result.Tables ?? Array.Empty<ReaderTable>(),
            Assets = result.Assets ?? Array.Empty<OfficeDocumentAsset>(),
            Links = result.Links ?? Array.Empty<OfficeDocumentLink>(),
            Forms = result.Forms ?? Array.Empty<OfficeDocumentFormField>(),
            OcrCandidates = effectiveOptions.RemoveResolvedCandidates ? unresolved.ToArray() : candidates,
            Visuals = result.Visuals ?? Array.Empty<ReaderVisual>(),
            Diagnostics = BuildDiagnostics(result.Diagnostics, unresolved, effectiveOptions)
        };

        return new OfficeDocumentOcrEnrichmentResult {
            Document = enriched,
            Report = new OfficeDocumentOcrEnrichmentReport {
                CandidateCount = candidates.Count,
                ResultCount = suppliedResults.Length,
                AppliedResultCount = applied.Count,
                UnresolvedCandidateCount = effectiveOptions.RemoveResolvedCandidates ? unresolved.Count : candidates.Count - applied.Count,
                UnmatchedResultCount = unmatchedIds.Length,
                EnrichedBlockCount = enrichedBlocks.Length,
                EnrichedChunkCount = enrichedChunks.Length,
                AppliedCandidateIds = applied.Select(static item => item.Candidate.Id).ToArray(),
                UnmatchedCandidateIds = unmatchedIds
            }
        };
    }

    private sealed class AppliedOcrText {
        public AppliedOcrText(OfficeDocumentOcrCandidate candidate, OfficeDocumentOcrTextResult result, int index) {
            Candidate = candidate;
            Result = result;
            Index = index;
        }

        public OfficeDocumentOcrCandidate Candidate { get; }

        public OfficeDocumentOcrTextResult Result { get; }

        public int Index { get; }
    }

    private static OfficeDocumentBlock BuildOcrBlock(OfficeDocumentSource source, AppliedOcrText applied, OfficeDocumentOcrEnrichmentOptions options) {
        return new OfficeDocumentBlock {
            Id = BuildOcrObjectId(applied.Candidate.Id, "block"),
            Kind = string.IsNullOrWhiteSpace(options.BlockKind) ? "ocr-text" : options.BlockKind.Trim(),
            Text = applied.Result.Text.Trim(),
            Location = BuildOcrLocation(source, applied, "ocr-text"),
            Region = CloneRegion(applied.Candidate.Region)
        };
    }

    private static ReaderChunk BuildOcrChunk(OfficeDocumentReadResult result, AppliedOcrText applied) {
        string text = applied.Result.Text.Trim();
        return new ReaderChunk {
            Id = BuildOcrObjectId(applied.Candidate.Id, "chunk"),
            Kind = result.Kind,
            Location = BuildOcrLocation(result.Source, applied, "ocr-text"),
            SourceId = result.Source?.SourceId,
            SourceHash = result.Source?.SourceHash,
            SourceLastWriteUtc = result.Source?.LastWriteUtc,
            SourceLengthBytes = result.Source?.LengthBytes,
            TokenEstimate = Math.Max(1, text.Length / 4),
            Text = text,
            Markdown = text
        };
    }

    private static ReaderLocation BuildOcrLocation(OfficeDocumentSource source, AppliedOcrText applied, string sourceBlockKind) {
        ReaderLocation location = CloneLocation(applied.Candidate.Location);
        if (location.Path == null) {
            location.Path = source.Path;
        }

        location.SourceBlockKind = sourceBlockKind;
        location.BlockAnchor = BuildOcrObjectId(applied.Candidate.Id, "text");
        return location;
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildPages(IReadOnlyList<OfficeDocumentPage>? pages, IReadOnlyList<OfficeDocumentBlock> enrichedBlocks, IReadOnlyList<OfficeDocumentOcrCandidate> remainingCandidates) {
        if (pages == null || pages.Count == 0) {
            return Array.Empty<OfficeDocumentPage>();
        }

        OfficeDocumentPage[] enrichedPages = new OfficeDocumentPage[pages.Count];
        for (int i = 0; i < pages.Count; i++) {
            OfficeDocumentPage page = pages[i];
            OfficeDocumentBlock[] pageBlocks = enrichedBlocks
                .Where(block => IsSameContainer(block.Location, page.Location))
                .ToArray();
            OfficeDocumentOcrCandidate[] pageCandidates = remainingCandidates
                .Where(candidate => IsSameContainer(candidate.Location, page.Location))
                .ToArray();
            enrichedPages[i] = new OfficeDocumentPage {
                Number = page.Number,
                Name = page.Name,
                Width = page.Width,
                Height = page.Height,
                RotationDegrees = page.RotationDegrees,
                Location = page.Location,
                Blocks = Append(page.Blocks, pageBlocks),
                Tables = page.Tables ?? Array.Empty<ReaderTable>(),
                Assets = page.Assets ?? Array.Empty<OfficeDocumentAsset>(),
                Links = page.Links ?? Array.Empty<OfficeDocumentLink>(),
                Forms = page.Forms ?? Array.Empty<OfficeDocumentFormField>(),
                OcrCandidates = pageCandidates
            };
        }

        return enrichedPages;
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildOcrMetadata(IReadOnlyList<AppliedOcrText> applied, int unresolvedCount, int unmatchedCount) {
        var entries = new List<OfficeDocumentMetadataEntry>();
        AddCountMetadata(entries, "reader-ocr-applied-count", "reader.ocr", "AppliedCount", applied.Count);
        AddCountMetadata(entries, "reader-ocr-unresolved-candidate-count", "reader.ocr", "UnresolvedCandidateCount", unresolvedCount);
        AddCountMetadata(entries, "reader-ocr-unmatched-result-count", "reader.ocr", "UnmatchedResultCount", unmatchedCount);

        for (int i = 0; i < applied.Count; i++) {
            AppliedOcrText item = applied[i];
            entries.Add(new OfficeDocumentMetadataEntry {
                Id = "reader-ocr-applied-" + (i + 1).ToString("D4", CultureInfo.InvariantCulture),
                Category = "reader.ocr",
                Name = "AppliedCandidateId",
                Value = item.Candidate.Id,
                ValueType = "string",
                SourceObjectId = item.Candidate.Id,
                Location = item.Candidate.Location,
                Attributes = BuildOcrAttributes(item.Result)
            });
        }

        return entries;
    }

    private static IReadOnlyDictionary<string, string> BuildOcrAttributes(OfficeDocumentOcrTextResult result) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!string.IsNullOrWhiteSpace(result.Provider)) {
            attributes["provider"] = result.Provider!.Trim();
        }

        if (!string.IsNullOrWhiteSpace(result.Model)) {
            attributes["model"] = result.Model!.Trim();
        }

        if (!string.IsNullOrWhiteSpace(result.Language)) {
            attributes["language"] = result.Language!.Trim();
        }

        if (result.Confidence.HasValue) {
            attributes["confidence"] = result.Confidence.Value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        return attributes;
    }

    private static IReadOnlyList<OfficeDocumentDiagnostic> BuildDiagnostics(IReadOnlyList<OfficeDocumentDiagnostic>? diagnostics, IReadOnlyList<OfficeDocumentOcrCandidate> unresolved, OfficeDocumentOcrEnrichmentOptions options) {
        IReadOnlyList<OfficeDocumentDiagnostic> existing = diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>();
        if (!options.RemoveResolvedOcrNeededDiagnostics) {
            return existing;
        }

        var rebuilt = new List<OfficeDocumentDiagnostic>();
        for (int i = 0; i < existing.Count; i++) {
            if (!IsOcrNeededDiagnostic(existing[i])) {
                rebuilt.Add(existing[i]);
            }
        }

        for (int i = 0; i < unresolved.Count; i++) {
            rebuilt.Add(BuildOcrNeededDiagnostic(unresolved[i]));
        }

        return rebuilt.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : rebuilt.ToArray();
    }

    private static bool IsOcrNeededDiagnostic(OfficeDocumentDiagnostic diagnostic) =>
        string.Equals(diagnostic.Code, "ocr-needed", StringComparison.Ordinal);

    private static OfficeDocumentDiagnostic BuildOcrNeededDiagnostic(OfficeDocumentOcrCandidate candidate) =>
        new OfficeDocumentDiagnostic {
            Severity = OfficeDocumentDiagnosticSeverity.Warning,
            Category = OfficeDocumentDiagnosticCategory.Ocr,
            Code = "ocr-needed",
            Message = candidate.Reason ?? "OCR should be considered for this source region.",
            Source = "officeimo.reader.ocr",
            IsRecoverable = true,
            Location = candidate.Location
        };

    private static string? BuildMarkdown(string? markdown, IReadOnlyList<AppliedOcrText> applied, OfficeDocumentOcrEnrichmentOptions options) {
        if (!options.AppendRecognizedTextToMarkdown || applied.Count == 0) {
            return markdown;
        }

        var builder = new System.Text.StringBuilder();
        if (!string.IsNullOrWhiteSpace(markdown)) {
            builder.Append(markdown!.TrimEnd());
            builder.AppendLine();
            builder.AppendLine();
        }

        for (int i = 0; i < applied.Count; i++) {
            if (i > 0) {
                builder.AppendLine();
            }

            builder.Append(applied[i].Result.Text.Trim());
            builder.AppendLine();
        }

        return builder.ToString().TrimEnd();
    }

    private static IReadOnlyList<string> AppendCapability(IReadOnlyList<string>? capabilities, string capability) {
        string[] existing = capabilities == null || capabilities.Count == 0 ? Array.Empty<string>() : capabilities.ToArray();
        if (existing.Contains(capability, StringComparer.Ordinal)) {
            return existing;
        }

        return existing.Concat(new[] { capability }).ToArray();
    }

    private static IReadOnlyList<T> Append<T>(IReadOnlyList<T>? existing, IReadOnlyList<T> appended) {
        if ((existing == null || existing.Count == 0) && appended.Count == 0) {
            return Array.Empty<T>();
        }

        if (existing == null || existing.Count == 0) {
            return appended.ToArray();
        }

        if (appended.Count == 0) {
            return existing;
        }

        return existing.Concat(appended).ToArray();
    }

    private static void AddCountMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string category, string name, int count) {
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        });
    }

    private static bool IsSameContainer(ReaderLocation? left, ReaderLocation? right) {
        if (left == null || right == null) {
            return false;
        }

        if (left.Page.HasValue || right.Page.HasValue) {
            return left.Page == right.Page;
        }

        if (left.Slide.HasValue || right.Slide.HasValue) {
            return left.Slide == right.Slide;
        }

        if (!string.IsNullOrWhiteSpace(left.Sheet) || !string.IsNullOrWhiteSpace(right.Sheet)) {
            return string.Equals(left.Sheet, right.Sheet, StringComparison.Ordinal);
        }

        return string.Equals(left.Path, right.Path, StringComparison.Ordinal);
    }

    private static string BuildOcrObjectId(string candidateId, string suffix) {
        string id = string.IsNullOrWhiteSpace(candidateId) ? "ocr-candidate" : candidateId.Trim();
        return id + "-" + suffix;
    }

    private static ReaderLocation CloneLocation(ReaderLocation? location) {
        if (location == null) {
            return new ReaderLocation();
        }

        return new ReaderLocation {
            Path = location.Path,
            BlockIndex = location.BlockIndex,
            SourceBlockIndex = location.SourceBlockIndex,
            StartLine = location.StartLine,
            EndLine = location.EndLine,
            NormalizedStartLine = location.NormalizedStartLine,
            NormalizedEndLine = location.NormalizedEndLine,
            HeadingPath = location.HeadingPath,
            HeadingSlug = location.HeadingSlug,
            SourceBlockKind = location.SourceBlockKind,
            BlockAnchor = location.BlockAnchor,
            Sheet = location.Sheet,
            A1Range = location.A1Range,
            Slide = location.Slide,
            Page = location.Page,
            TableIndex = location.TableIndex
        };
    }

    private static OfficeDocumentRegion? CloneRegion(OfficeDocumentRegion? region) {
        if (region == null) {
            return null;
        }

        return new OfficeDocumentRegion {
            X = region.X,
            Y = region.Y,
            Width = region.Width,
            Height = region.Height
        };
    }
}
