using System.Globalization;
using System.IO.Compression;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Reader;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.RealWorldCorpus;

/// <summary>
/// Reads one Open XML package through OfficeIMO's own layer and through the Open XML SDK (or
/// <see cref="System.IO.Packaging"/> for package structure), and classifies each lane as a match, a mismatch with its
/// first difference, or an error on one or both sides. The SDK is the reference.
/// </summary>
internal static class OpenXmlDifferentialProbe {
    internal const string Package = "package";
    internal const string WordBodyParagraphs = "word-body-paragraphs";
    internal const string WordTextProjection = "word-text-projection";
    internal const string SlideParagraphs = "slide-paragraphs";
    internal const string PowerPointTextProjection = "powerpoint-text-projection";

    internal static readonly string[] Lanes = { Package, WordBodyParagraphs, WordTextProjection, SlideParagraphs, PowerPointTextProjection };

    private const int SnippetRadius = 40;
    private const int MaxMessageLength = 300;

    // Exception messages describe the package (part names, ids), not its text, but they are recorded only with the
    // difference text, which a run opts in to.
    private static bool s_includeText;

    /// <summary>Written to standard error as each lane starts, so the coordinator can name the lane a timeout stopped.</summary>
    internal const string LaneStartPrefix = "openxml-lane-start:";

    private static void Start(string lane) {
        Console.Error.WriteLine(LaneStartPrefix + lane);
        Console.Error.Flush();
    }

    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder()
        .AddWordHandler()
        .AddPowerPointHandler()
        .Build();

    public static CorpusOpenXmlEvidence Probe(byte[] snapshot, string sourceName, ReaderInputKind family, long maxFileBytes, bool includeText) {
        s_includeText = includeText;
        var evidence = new CorpusOpenXmlEvidence { Family = family };
        evidence.Lanes.Add(ComparePackage(snapshot));
        if (family == ReaderInputKind.Word) {
            evidence.Lanes.Add(CompareLines(WordBodyParagraphs, () => ReadSdkWordBodyParagraphs(snapshot),
                OpenXmlOwnedLayer.IsAvailable ? () => OpenXmlOwnedLayer.ReadWordBodyParagraphs(snapshot) : null, includeText));
            CorpusOpenXmlLane projection = CompareProjection(WordTextProjection, snapshot, sourceName, maxFileBytes, includeText);
            if (OpenXmlOwnedLayer.WordStreamedTextAvailable) {
                Start("word-streamed-coverage");
                ReaderOptions options = CorpusReaderPolicy.CreateReadOptions(maxFileBytes);
                projection.StreamedDeclineReason = OpenXmlOwnedLayer.GetWordStreamedDeclineReason(snapshot, options);
                projection.Streamed = projection.StreamedDeclineReason == null;
            }
            evidence.Lanes.Add(projection);
        } else if (family == ReaderInputKind.PowerPoint) {
            evidence.Lanes.Add(CompareLines(SlideParagraphs, () => ReadSdkSlideParagraphs(snapshot),
                OpenXmlOwnedLayer.IsAvailable ? () => OpenXmlOwnedLayer.ReadSlideParagraphs(snapshot) : null, includeText));
            evidence.Lanes.Add(CompareProjection(PowerPointTextProjection, snapshot, sourceName, maxFileBytes, includeText));
        }
        evidence.Outcome = CorpusOpenXmlOutcomes.Worst(evidence.Lanes.Select(lane => lane.Outcome));
        return evidence;
    }

    private static CorpusOpenXmlLane ComparePackage(byte[] snapshot) {
        Start(Package);
        var lane = new CorpusOpenXmlLane { Name = Package };
        OpenXmlPackageView? expected = null;
        try {
            expected = ReadSystemPackagingView(snapshot);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.SdkException = TypeName(exception);
            lane.SdkExceptionMessage = Message(exception);
        }
        if (!OpenXmlOwnedLayer.IsAvailable) {
            lane.Outcome = CorpusOpenXmlOutcomes.Unavailable;
            return lane;
        }
        OpenXmlPackageView? actual = null;
        try {
            actual = OpenXmlOwnedLayer.ReadPackage(snapshot, CandidatePartNames(snapshot, expected));
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.OurException = TypeName(exception);
            lane.OurExceptionMessage = Message(exception);
        }
        if (expected == null || actual == null) {
            lane.Outcome = ClassifyErrors(expected != null, actual != null);
            return lane;
        }
        List<CorpusOpenXmlDifference> differences = expected.Compare(actual);
        lane.Outcome = differences.Count == 0 ? CorpusOpenXmlOutcomes.Match : CorpusOpenXmlOutcomes.Mismatch;
        lane.DifferenceCount = differences.Count;
        lane.Classes = differences.Select(difference => difference.Class).Distinct(StringComparer.Ordinal).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        lane.Differences = differences.Take(5).ToArray();
        lane.ExpectedCount = expected.Count;
        lane.ActualCount = actual.Count;
        return lane;
    }

    // The owned layer cannot enumerate parts, so it is asked about every ZIP member and every part System.IO.Packaging
    // reported; a part only the owned layer accepts then shows up as an extra part.
    private static IEnumerable<string> CandidatePartNames(byte[] snapshot, OpenXmlPackageView? expected) {
        var names = new List<string>();
        if (expected != null) names.AddRange(expected.PartNames);
        try {
            using var stream = new MemoryStream(snapshot, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            foreach (ZipArchiveEntry entry in archive.Entries) {
                string name = entry.FullName.Replace('\\', '/');
                if (name.Length == 0 || name.EndsWith('/') || string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                names.Add("/" + name);
            }
        } catch (Exception exception) when (exception is InvalidDataException or IOException or ArgumentException) {
        }
        return names;
    }

    private static OpenXmlPackageView ReadSystemPackagingView(byte[] snapshot) {
        using var stream = new MemoryStream(snapshot, writable: false);
        using System.IO.Packaging.Package package = System.IO.Packaging.Package.Open(stream, FileMode.Open, FileAccess.Read);
        var view = new OpenXmlPackageView();
        AddRelationships(view, null, package.GetRelationships());
        foreach (PackagePart part in package.GetParts()) {
            string name = part.Uri.OriginalString;
            if (OpenXmlPackageView.IsRelationshipsPart(name)) continue;
            view.AddPart(name, part.ContentType);
            AddRelationships(view, part.Uri, part.GetRelationships());
        }
        return view;
    }

    private static void AddRelationships(OpenXmlPackageView view, Uri? source, PackageRelationshipCollection relationships) {
        Uri sourceUri = source ?? new Uri("/", UriKind.Relative);
        foreach (PackageRelationship relationship in relationships) {
            bool external = relationship.TargetMode == TargetMode.External;
            string? target;
            if (external) {
                target = relationship.TargetUri.OriginalString;
            } else if (relationship.TargetUri.OriginalString.StartsWith('#')) {
                // A fragment-only target (an in-document hyperlink) names no part; System.IO.Packaging resolves it to
                // the source part, the owned layer to nothing.
                target = null;
            } else {
                try {
                    target = PackUriHelper.ResolvePartUri(sourceUri, relationship.TargetUri).OriginalString;
                } catch (Exception exception) when (exception is ArgumentException or UriFormatException) {
                    target = null;
                }
            }
            view.AddRelationship(source?.OriginalString, relationship.Id, NormalizeRelationshipType(relationship.RelationshipType), external, target);
        }
    }

    // The owned layer reports Strict relationship types with their transitional spelling, as the SDK's typed parts do.
    private static string NormalizeRelationshipType(string type) {
        const string strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
        const string transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        return type.StartsWith(strict, StringComparison.Ordinal) ? transitional + type.Substring(strict.Length) : type;
    }

    private static OpenSettings Office2007Settings() => new() {
        MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
            MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
    };

    // The body paragraphs' InnerText with markup compatibility processed as an Office 2007 consumer, the reference the
    // owned paragraph read is specified against. Undeclared mc:Ignorable prefixes stop the SDK's processing, so it then
    // reads the markup unprocessed.
    private static List<string> ReadSdkWordBodyParagraphs(byte[] snapshot) {
        try {
            return ReadSdkWordBodyParagraphs(snapshot, Office2007Settings());
        } catch (InvalidMCContentException) {
            return ReadSdkWordBodyParagraphs(snapshot, new OpenSettings());
        }
    }

    private static List<string> ReadSdkWordBodyParagraphs(byte[] snapshot, OpenSettings settings) {
        using var stream = new MemoryStream(snapshot, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false, settings);
        W.Body body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidDataException("The package has no document body.");
        return body.Elements<W.Paragraph>().Select(paragraph => paragraph.InnerText).ToList();
    }

    // Every DrawingML paragraph's InnerText, slide by slide in presentation order.
    private static List<string> ReadSdkSlideParagraphs(byte[] snapshot) {
        using var stream = new MemoryStream(snapshot, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false, Office2007Settings());
        PresentationPart presentation = document.PresentationPart ?? throw new InvalidDataException("The package has no presentation part.");
        var paragraphs = new List<string>();
        int index = 0;
        foreach (P.SlideId slideId in presentation.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? Enumerable.Empty<P.SlideId>()) {
            var slide = (SlidePart)presentation.GetPartById(slideId.RelationshipId?.Value ?? throw new InvalidDataException("A slide id has no relationship."));
            foreach (A.Paragraph paragraph in slide.Slide?.Descendants<A.Paragraph>() ?? Enumerable.Empty<A.Paragraph>()) {
                paragraphs.Add(index.ToString(CultureInfo.InvariantCulture) + ":" + paragraph.InnerText);
            }
            index++;
        }
        return paragraphs;
    }

    // The Reader's text projection (the streamed path, falling back to the model when it declines) against the
    // text-and-metadata projection, which always reads through the SDK-backed model.
    private static CorpusOpenXmlLane CompareProjection(string name, byte[] snapshot, string sourceName, long maxFileBytes, bool includeText) {
        return CompareLines(name,
            () => ReadProjection(snapshot, sourceName, maxFileBytes, ReaderProjection.TextAndMetadata),
            () => ReadProjection(snapshot, sourceName, maxFileBytes, ReaderProjection.Text),
            includeText);
    }

    private static List<string> ReadProjection(byte[] snapshot, string sourceName, long maxFileBytes, ReaderProjection projection) {
        ReaderOptions options = CorpusReaderPolicy.CreateReadOptions(maxFileBytes);
        options.Projection = projection;
        OfficeDocumentReadResult result = Reader.ReadDocument(snapshot, sourceName, options);
        if (result.Diagnostics.Any(diagnostic => diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error)) {
            throw new CorpusReadDiagnosticException(result.Diagnostics
                .Where(diagnostic => diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error)
                .Select(diagnostic => diagnostic.Code).FirstOrDefault() ?? "error");
        }
        return result.GetText().Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').ToList();
    }

    private static CorpusOpenXmlLane CompareLines(string name, Func<List<string>> readExpected, Func<List<string>>? readActual, bool includeText) {
        Start(name);
        var lane = new CorpusOpenXmlLane { Name = name };
        List<string>? expected = null;
        try {
            expected = readExpected();
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.SdkException = TypeName(exception);
            lane.SdkExceptionMessage = Message(exception);
        }
        if (readActual == null) {
            lane.Outcome = CorpusOpenXmlOutcomes.Unavailable;
            return lane;
        }
        List<string>? actual = null;
        try {
            actual = readActual();
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.OurException = TypeName(exception);
            lane.OurExceptionMessage = Message(exception);
        }
        if (expected == null || actual == null) {
            lane.Outcome = ClassifyErrors(expected != null, actual != null);
            return lane;
        }
        lane.ExpectedCount = expected.Count;
        lane.ActualCount = actual.Count;
        CorpusOpenXmlDifference? difference = FirstDifference(expected, actual, includeText);
        if (difference == null) {
            lane.Outcome = CorpusOpenXmlOutcomes.Match;
            return lane;
        }
        lane.Outcome = CorpusOpenXmlOutcomes.Mismatch;
        lane.DifferenceCount = CountDifferences(expected, actual);
        lane.Classes = new[] { difference.Class };
        lane.Differences = new[] { difference };
        return lane;
    }

    private static int CountDifferences(List<string> expected, List<string> actual) {
        int count = Math.Abs(expected.Count - actual.Count);
        for (int index = 0; index < Math.Min(expected.Count, actual.Count); index++) {
            if (!string.Equals(expected[index], actual[index], StringComparison.Ordinal)) count++;
        }
        return count;
    }

    // The first line that differs, classified by how it differs. Lines are compared in order, so a missing or extra
    // line is reported where the sequences first diverge.
    internal static CorpusOpenXmlDifference? FirstDifference(IReadOnlyList<string> expected, IReadOnlyList<string> actual, bool includeText) {
        int count = Math.Min(expected.Count, actual.Count);
        for (int index = 0; index < count; index++) {
            if (string.Equals(expected[index], actual[index], StringComparison.Ordinal)) continue;
            string left = expected[index];
            string right = actual[index];
            int offset = 0;
            while (offset < left.Length && offset < right.Length && left[offset] == right[offset]) offset++;
            var difference = new CorpusOpenXmlDifference {
                Class = ClassifyText(left, right),
                Index = index,
                Offset = offset,
                ExpectedLength = left.Length,
                ActualLength = right.Length
            };
            if (includeText) {
                difference.Expected = Snippet(left, offset);
                difference.Actual = Snippet(right, offset);
            }
            return difference;
        }
        if (expected.Count == actual.Count) return null;
        var countDifference = new CorpusOpenXmlDifference {
            Class = expected.Count > actual.Count ? "missing-lines" : "extra-lines",
            Index = count,
            ExpectedLength = expected.Count,
            ActualLength = actual.Count
        };
        if (includeText) {
            countDifference.Expected = count < expected.Count ? Snippet(expected[count], 0) : null;
            countDifference.Actual = count < actual.Count ? Snippet(actual[count], 0) : null;
        }
        return countDifference;
    }

    private static string ClassifyText(string expected, string actual) {
        if (string.Equals(Collapse(expected), Collapse(actual), StringComparison.Ordinal)) return "whitespace";
        if (actual.Length == 0) return "missing-text";
        if (expected.Length == 0) return "extra-text";
        if (expected.Contains(actual, StringComparison.Ordinal)) return "missing-text";
        if (actual.Contains(expected, StringComparison.Ordinal)) return "extra-text";
        return "different-text";
    }

    private static string Collapse(string value) => new(value.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static string Snippet(string value, int offset) {
        int start = Math.Max(0, offset - SnippetRadius);
        int end = Math.Min(value.Length, offset + SnippetRadius);
        return (start > 0 ? "…" : string.Empty) + value.Substring(start, end - start) + (end < value.Length ? "…" : string.Empty);
    }

    private static string ClassifyErrors(bool sdkRead, bool ourRead) =>
        sdkRead ? CorpusOpenXmlOutcomes.OurError : ourRead ? CorpusOpenXmlOutcomes.SdkError : CorpusOpenXmlOutcomes.BothReject;

    private static string? Message(Exception exception) {
        if (!s_includeText) return null;
        string message = exception.Message.ReplaceLineEndings(" ");
        return message.Length <= MaxMessageLength ? message : message.Substring(0, MaxMessageLength) + "…";
    }

    private static string TypeName(Exception exception) => exception is CorpusReadDiagnosticException diagnostic
        ? "diagnostic:" + diagnostic.Code
        : exception.GetType().FullName ?? exception.GetType().Name;
}

/// <summary>The parts and relationships a package reader reports, keyed by part identity, for comparison.</summary>
internal sealed class OpenXmlPackageView {
    private readonly Dictionary<string, (string Name, string ContentType)> _parts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, (string Type, bool External, string? Target)> _relationships = new(StringComparer.Ordinal);

    internal int Count => _parts.Count + _relationships.Count;

    internal IEnumerable<string> PartNames => _parts.Values.Select(part => part.Name);

    internal static bool IsRelationshipsPart(string partName) =>
        partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) && partName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    internal bool AddPart(string name, string contentType) {
        string key = Identity(name);
        if (_parts.ContainsKey(key)) return false;
        _parts.Add(key, (name, contentType));
        return true;
    }

    internal void AddRelationship(string? source, string id, string type, bool external, string? target) {
        _relationships[(source == null ? "/" : Identity(source)) + "#" + id] = (type, external, external || target == null ? target : Identity(target));
    }

    // Part names compare percent-decoded once and ASCII case-insensitively, as OPC part names do.
    internal static string Identity(string partName) {
        string decoded;
        try {
            decoded = Uri.UnescapeDataString(partName);
        } catch (UriFormatException) {
            decoded = partName;
        }
        return (decoded.StartsWith('/') ? decoded : "/" + decoded).ToLowerInvariant();
    }

    internal List<CorpusOpenXmlDifference> Compare(OpenXmlPackageView actual) {
        var differences = new List<CorpusOpenXmlDifference>();
        foreach ((string key, (string name, string contentType)) in _parts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            if (!actual._parts.TryGetValue(key, out (string Name, string ContentType) other)) {
                differences.Add(Difference("part-missing", name, contentType, null));
            } else if (!ContentTypesEqual(contentType, other.ContentType)) {
                differences.Add(Difference("content-type", name, contentType, other.ContentType));
            }
        }
        foreach ((string key, (string name, string contentType)) in actual._parts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            if (!_parts.ContainsKey(key)) differences.Add(Difference("part-extra", name, null, contentType));
        }
        foreach ((string key, (string type, bool external, string? target)) in _relationships.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            if (!actual._relationships.TryGetValue(key, out (string Type, bool External, string? Target) other)) {
                differences.Add(Difference("relationship-missing", key, type, null));
            } else if (!string.Equals(type, other.Type, StringComparison.Ordinal)) {
                differences.Add(Difference("relationship-type", key, type, other.Type));
            } else if (external != other.External) {
                differences.Add(Difference("relationship-mode", key, external ? "External" : "Internal", other.External ? "External" : "Internal"));
            } else if (!string.Equals(target, other.Target, StringComparison.Ordinal)) {
                differences.Add(Difference("relationship-target", key, target ?? "(unresolved)", other.Target ?? "(unresolved)"));
            }
        }
        foreach ((string key, (string type, _, _)) in actual._relationships.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            if (!_relationships.ContainsKey(key)) differences.Add(Difference("relationship-extra", key, null, type));
        }
        return differences;
    }

    // Content types are MIME types: the type and parameter names compare case-insensitively, and spacing around
    // parameters is not significant.
    private static bool ContentTypesEqual(string expected, string actual) =>
        string.Equals(NormalizeContentType(expected), NormalizeContentType(actual), StringComparison.OrdinalIgnoreCase);

    private static string NormalizeContentType(string value) =>
        string.Join(";", value.Split(';').Select(segment => segment.Trim()));

    private static CorpusOpenXmlDifference Difference(string kind, string subject, string? expected, string? actual) => new() {
        Class = kind,
        Subject = subject,
        Expected = expected,
        Actual = actual
    };
}

internal sealed class CorpusReadDiagnosticException : Exception {
    internal CorpusReadDiagnosticException(string code) : base("The read reported an error diagnostic.") => Code = code;

    internal string Code { get; }
}

internal static class CorpusOpenXmlOutcomes {
    public const string Match = "match";
    public const string Mismatch = "mismatch";
    public const string OurError = "our-error";
    public const string SdkError = "sdk-error";
    public const string BothReject = "both-reject";
    public const string Unavailable = "unavailable";

    // Ordered from most to least in need of attention.
    internal static readonly string[] Severity = { Mismatch, OurError, SdkError, BothReject, Match, Unavailable };

    internal static string Worst(IEnumerable<string> outcomes) {
        int worst = Severity.Length - 1;
        foreach (string outcome in outcomes) worst = Math.Min(worst, Array.IndexOf(Severity, outcome));
        return Severity[worst];
    }
}

internal sealed class CorpusOpenXmlEvidence {
    public ReaderInputKind Family { get; set; }
    public string Outcome { get; set; } = CorpusOpenXmlOutcomes.Unavailable;
    public List<CorpusOpenXmlLane> Lanes { get; set; } = new();
}

internal sealed class CorpusOpenXmlLane {
    public string Name { get; set; } = string.Empty;
    public string Outcome { get; set; } = CorpusOpenXmlOutcomes.Unavailable;
    public int? ExpectedCount { get; set; }
    public int? ActualCount { get; set; }
    public int DifferenceCount { get; set; }
    public IReadOnlyList<string> Classes { get; set; } = Array.Empty<string>();
    public IReadOnlyList<CorpusOpenXmlDifference> Differences { get; set; } = Array.Empty<CorpusOpenXmlDifference>();
    public string? SdkException { get; set; }
    public string? SdkExceptionMessage { get; set; }
    public string? OurException { get; set; }
    public string? OurExceptionMessage { get; set; }
    public bool? Streamed { get; set; }
    public string? StreamedDeclineReason { get; set; }
}

internal sealed class CorpusOpenXmlDifference {
    public string Class { get; set; } = string.Empty;
    public string? Subject { get; set; }
    public int? Index { get; set; }
    public int? Offset { get; set; }
    public int? ExpectedLength { get; set; }
    public int? ActualLength { get; set; }
    public string? Expected { get; set; }
    public string? Actual { get; set; }
}

internal static class CorpusOpenXmlFamilies {
    private static readonly Dictionary<string, ReaderInputKind> Extensions = new(StringComparer.OrdinalIgnoreCase) {
        [".docx"] = ReaderInputKind.Word, [".docm"] = ReaderInputKind.Word, [".dotx"] = ReaderInputKind.Word, [".dotm"] = ReaderInputKind.Word,
        [".xlsx"] = ReaderInputKind.Excel, [".xlsm"] = ReaderInputKind.Excel, [".xltx"] = ReaderInputKind.Excel, [".xltm"] = ReaderInputKind.Excel,
        [".xlam"] = ReaderInputKind.Excel,
        [".pptx"] = ReaderInputKind.PowerPoint, [".pptm"] = ReaderInputKind.PowerPoint, [".potx"] = ReaderInputKind.PowerPoint,
        [".potm"] = ReaderInputKind.PowerPoint, [".ppsx"] = ReaderInputKind.PowerPoint, [".ppsm"] = ReaderInputKind.PowerPoint,
        [".ppam"] = ReaderInputKind.PowerPoint
    };

    /// <summary>The Open XML family an extension claims, or <see cref="ReaderInputKind.Unknown"/>.</summary>
    internal static ReaderInputKind FromExtension(string extension) =>
        Extensions.TryGetValue(extension, out ReaderInputKind kind) ? kind : ReaderInputKind.Unknown;
}
