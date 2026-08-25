using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Common visual positions for Bates numbers on existing PDF pages.</summary>
public enum PdfBatesPosition {
    /// <summary>Bottom-left page corner.</summary>
    BottomLeft,
    /// <summary>Bottom-center page edge.</summary>
    BottomCenter,
    /// <summary>Bottom-right page corner.</summary>
    BottomRight,
    /// <summary>Top-left page corner.</summary>
    TopLeft,
    /// <summary>Top-center page edge.</summary>
    TopCenter,
    /// <summary>Top-right page corner.</summary>
    TopRight
}

/// <summary>One PDF supplied to a Bates-numbering batch.</summary>
public sealed class PdfBatesDocument {
    private readonly byte[] _pdf;

    /// <summary>Creates a batch input from PDF bytes.</summary>
    public PdfBatesDocument(byte[] pdf, string? name = null) {
        Guard.NotNull(pdf, nameof(pdf));
        _pdf = (byte[])pdf.Clone();
        Name = name;
    }

    /// <summary>Optional stable document name carried into the operation report.</summary>
    public string? Name { get; }

    /// <summary>Optional parser, password, permission, and resource-budget settings for this document.</summary>
    public PdfReadOptions? ReadOptions { get; set; }

    /// <summary>Optional document-specific page selection. Overrides <see cref="PdfBatesNumberingOptions.TargetPages"/>.</summary>
    public PdfPageSelector? TargetPages { get; set; }

    internal byte[] GetBytes() => (byte[])_pdf.Clone();
}

/// <summary>Controls continuous Bates numbering across one or more PDFs.</summary>
public sealed class PdfBatesNumberingOptions {
    /// <summary>First numeric value assigned to a selected page.</summary>
    public long StartNumber { get; set; } = 1L;

    /// <summary>Text written before the padded numeric value.</summary>
    public string Prefix { get; set; } = string.Empty;

    /// <summary>Text written after the padded numeric value.</summary>
    public string Suffix { get; set; } = string.Empty;

    /// <summary>Minimum number of decimal digits. Defaults to six.</summary>
    public int MinimumDigits { get; set; } = 6;

    /// <summary>Default page selection applied when an input does not provide its own selector.</summary>
    public PdfPageSelector? TargetPages { get; set; }

    /// <summary>Visual page position used for each number.</summary>
    public PdfBatesPosition Position { get; set; } = PdfBatesPosition.BottomRight;

    /// <summary>Horizontal distance from the selected page edge, in points.</summary>
    public double HorizontalMargin { get; set; } = 36D;

    /// <summary>Vertical distance from the selected page edge, in points.</summary>
    public double VerticalMargin { get; set; } = 24D;

    /// <summary>Height reserved for the number, in points.</summary>
    public double Height { get; set; } = 18D;

    /// <summary>Standard PDF font used by the number.</summary>
    public PdfStandardFont Font { get; set; } = PdfStandardFont.Helvetica;

    /// <summary>Font size in points.</summary>
    public double FontSize { get; set; } = 10D;

    /// <summary>Text color.</summary>
    public PdfColor Color { get; set; } = PdfColor.Black;
}

/// <summary>One assigned Bates number in a batch report.</summary>
public sealed class PdfBatesAssignment {
    internal PdfBatesAssignment(int documentIndex, string documentName, int pageNumber, long number, string text) {
        DocumentIndex = documentIndex;
        DocumentName = documentName;
        PageNumber = pageNumber;
        Number = number;
        Text = text;
    }

    /// <summary>Zero-based input document index.</summary>
    public int DocumentIndex { get; }
    /// <summary>Stable input document name.</summary>
    public string DocumentName { get; }
    /// <summary>One-based page number in the input document.</summary>
    public int PageNumber { get; }
    /// <summary>Numeric Bates value assigned to the page.</summary>
    public long Number { get; }
    /// <summary>Complete rendered text, including prefix, padding, and suffix.</summary>
    public string Text { get; }
}

/// <summary>Numbered output and preservation evidence for one input document.</summary>
public sealed class PdfBatesDocumentResult {
    private readonly byte[] _pdf;
    private readonly PdfReadOptions _readOptions;

    internal PdfBatesDocumentResult(
        int documentIndex,
        string documentName,
        byte[] pdf,
        IReadOnlyList<PdfBatesAssignment> assignments,
        PdfRewritePreservationReport preservation,
        PdfReadOptions readOptions) {
        DocumentIndex = documentIndex;
        DocumentName = documentName;
        _pdf = (byte[])pdf.Clone();
        Assignments = assignments;
        Preservation = preservation;
        _readOptions = readOptions;
    }

    /// <summary>Zero-based input document index.</summary>
    public int DocumentIndex { get; }
    /// <summary>Stable input document name.</summary>
    public string DocumentName { get; }
    /// <summary>Numbers assigned within this document.</summary>
    public IReadOnlyList<PdfBatesAssignment> Assignments { get; }
    /// <summary>Structural preservation comparison between input and output.</summary>
    public PdfRewritePreservationReport Preservation { get; }
    /// <summary>Returns an independent copy of the numbered PDF.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the numbered PDF through the public document API.</summary>
    public PdfDocument ToDocument(PdfReadOptions? readOptions = null) => PdfDocument.Open(_pdf, readOptions ?? _readOptions);
}

/// <summary>Continuous Bates-numbering outputs and assignments for a complete batch.</summary>
public sealed class PdfBatesBatchResult {
    internal PdfBatesBatchResult(IReadOnlyList<PdfBatesDocumentResult> documents, IReadOnlyList<PdfBatesAssignment> assignments, long nextNumber) {
        Documents = documents;
        Assignments = assignments;
        NextNumber = nextNumber;
    }

    /// <summary>Numbered documents in input order.</summary>
    public IReadOnlyList<PdfBatesDocumentResult> Documents { get; }
    /// <summary>All assigned numbers in numbering order.</summary>
    public IReadOnlyList<PdfBatesAssignment> Assignments { get; }
    /// <summary>Next unused number after the batch.</summary>
    public long NextNumber { get; }
}

/// <summary>Applies continuous, report-driven Bates numbering to PDF batches.</summary>
public static class PdfBatesNumberer {
    /// <summary>Numbers PDF byte arrays continuously using default options.</summary>
    public static PdfBatesBatchResult Apply(params byte[][] pdfs) {
        Guard.NotNull(pdfs, nameof(pdfs));
        return Apply(pdfs.Select((pdf, index) => new PdfBatesDocument(pdf, "document-" + (index + 1).ToString(CultureInfo.InvariantCulture))), null);
    }

    /// <summary>Numbers the supplied documents continuously and returns output and preservation evidence.</summary>
    public static PdfBatesBatchResult Apply(IEnumerable<PdfBatesDocument> documents, PdfBatesNumberingOptions? options = null) {
        Guard.NotNull(documents, nameof(documents));
        PdfBatesDocument[] inputs = documents.ToArray();
        if (inputs.Length == 0) throw new ArgumentException("At least one PDF must be supplied.", nameof(documents));
        if (inputs.Any(static document => document is null)) throw new ArgumentException("Bates-numbering inputs cannot contain null documents.", nameof(documents));

        PdfBatesNumberingOptions effective = options ?? new PdfBatesNumberingOptions();
        ValidateOptions(effective);
        long nextNumber = effective.StartNumber;
        var allAssignments = new List<PdfBatesAssignment>();
        var documentResults = new List<PdfBatesDocumentResult>(inputs.Length);

        for (int documentIndex = 0; documentIndex < inputs.Length; documentIndex++) {
            PdfBatesDocument input = inputs[documentIndex];
            byte[] source = input.GetBytes();
            PdfReadDocument read = PdfReadDocument.Open(source, input.ReadOptions);
            if (read.Pages.Count == 0) throw new ArgumentException("PDF input " + documentIndex.ToString(CultureInfo.InvariantCulture) + " does not contain any pages.", nameof(documents));
            IReadOnlyList<int> selectedPages = (input.TargetPages ?? effective.TargetPages)?.Resolve(read.Pages.Count) ??
                Enumerable.Range(1, read.Pages.Count).ToArray();
            if (selectedPages.Distinct().Count() != selectedPages.Count) {
                throw new ArgumentException("Bates page selections cannot contain duplicate pages.", nameof(documents));
            }
            string documentName = string.IsNullOrWhiteSpace(input.Name)
                ? "document-" + (documentIndex + 1).ToString(CultureInfo.InvariantCulture)
                : input.Name!;
            var assignments = new List<PdfBatesAssignment>(selectedPages.Count);
            foreach (int pageNumber in selectedPages) {
                string text = BuildText(nextNumber, effective);
                assignments.Add(new PdfBatesAssignment(documentIndex, documentName, pageNumber, nextNumber, text));
                nextNumber = checked(nextNumber + 1L);
            }

            var assignmentByPage = assignments.ToDictionary(static assignment => assignment.PageNumber);
            PdfGeneratedOutputGrowth generatedGrowth = default;
            byte[] output = assignments.Count == 0
                ? source
                : PdfStamper.StampCanvas(
                    source,
                    (canvas, context) => AddNumber(canvas, context, assignmentByPage[context.PageNumber].Text, effective),
                    out generatedGrowth,
                    new PdfCanvasStampOptions {
                        TargetPages = PdfPageSelector.Parse(string.Join(",", selectedPages.Select(static page => page.ToString(CultureInfo.InvariantCulture))))
                    },
                    input.ReadOptions);
            PdfReadOptions outputReadOptions = PdfReadOptions.ForGeneratedOutput(input.ReadOptions, source, output, generatedGrowth);
            PdfRewritePreservationReport preservation = PdfRewritePreservation.Assess(
                source,
                output,
                options: null,
                originalReadOptions: input.ReadOptions,
                rewrittenReadOptions: outputReadOptions);
            preservation.ThrowIfFailed();
            var readOnlyAssignments = new ReadOnlyCollection<PdfBatesAssignment>(assignments);
            documentResults.Add(new PdfBatesDocumentResult(documentIndex, documentName, output, readOnlyAssignments, preservation, outputReadOptions));
            allAssignments.AddRange(assignments);
        }

        return new PdfBatesBatchResult(
            new ReadOnlyCollection<PdfBatesDocumentResult>(documentResults),
            new ReadOnlyCollection<PdfBatesAssignment>(allAssignments),
            nextNumber);
    }

    private static string BuildText(long number, PdfBatesNumberingOptions options) =>
        options.Prefix + number.ToString("D" + options.MinimumDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture) + options.Suffix;

    private static void AddNumber(PdfPageCanvas canvas, PdfStampPageContext context, string text, PdfBatesNumberingOptions options) {
        double availableWidth = context.Width - (2D * options.HorizontalMargin);
        if (availableWidth <= 0D || options.VerticalMargin + options.Height > context.Height) {
            throw new InvalidOperationException("Bates-number margins leave no drawable page area.");
        }
        if (text.Contains('\r') || text.Contains('\n')) {
            throw new InvalidOperationException("Bates labels must fit on one line and cannot contain line breaks.");
        }
        double requiredWidth = PdfWriter.EstimateSimpleTextWidth(text, options.Font, options.FontSize);
        double requiredHeight = options.FontSize * 1.2D;
        if (requiredWidth > availableWidth + 0.01D || requiredHeight > options.Height + 0.01D) {
            throw new InvalidOperationException("Bates label does not fit the configured page rectangle.");
        }
        bool top = options.Position is PdfBatesPosition.TopLeft or PdfBatesPosition.TopCenter or PdfBatesPosition.TopRight;
        PdfAlign alignment = options.Position is PdfBatesPosition.BottomCenter or PdfBatesPosition.TopCenter
            ? PdfAlign.Center
            : options.Position is PdfBatesPosition.BottomRight or PdfBatesPosition.TopRight ? PdfAlign.Right : PdfAlign.Left;
        double y = top ? options.VerticalMargin : context.Height - options.VerticalMargin - options.Height;
        canvas.Text(text, options.HorizontalMargin, y, availableWidth, options.Height, options.FontSize, options.Color, alignment, options.Font);
    }

    private static void ValidateOptions(PdfBatesNumberingOptions options) {
        if (options.StartNumber < 0L) throw new ArgumentOutOfRangeException(nameof(options), "The first Bates number cannot be negative.");
        if (options.MinimumDigits < 1 || options.MinimumDigits > 18) throw new ArgumentOutOfRangeException(nameof(options), "Minimum digits must be between 1 and 18.");
        if (options.Prefix is null || options.Suffix is null) throw new ArgumentException("Bates prefix and suffix cannot be null.", nameof(options));
        if (options.Prefix.Length + options.Suffix.Length > 512) throw new ArgumentException("Bates prefix and suffix are too long.", nameof(options));
        if (options.Position < PdfBatesPosition.BottomLeft || options.Position > PdfBatesPosition.TopRight) throw new ArgumentOutOfRangeException(nameof(options), "Bates position must be a defined value.");
        if (!IsPositiveFinite(options.HorizontalMargin) || !IsPositiveFinite(options.VerticalMargin) || !IsPositiveFinite(options.Height) || !IsPositiveFinite(options.FontSize)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Bates geometry and font size must be positive finite values.");
        }
        Guard.StandardFont(options.Font, nameof(options), "Bates font must be a supported standard PDF font.");
    }

    private static bool IsPositiveFinite(double value) => value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);
}
