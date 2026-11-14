using System.Globalization;
using System.Text;
using OfficeIMO.Diff;

namespace OfficeIMO.Pdf;

/// <summary>
/// Compares the text of two PDFs word by word and marks the differences on the new PDF with review annotations.
/// </summary>
/// <remarks>
/// Words are read in reading order from every page, so edits are found across line and page breaks and reflowed text
/// is not reported as changed. Insertions and replacements get highlights; deletions get a caret at the point where the text was removed.
/// Replacements, deletions and moves carry the old text in their pop-up note. Page content is not changed.
/// This compares extracted text only; use <see cref="PdfVisualComparer"/> for rendered appearance.
/// </remarks>
public static class PdfTextRedline {
    /// <summary>Compares two PDFs and returns the annotated new PDF and the change list.</summary>
    /// <param name="oldPdf">Original PDF.</param>
    /// <param name="newPdf">Revised PDF; the annotations are added to a copy of it.</param>
    /// <param name="options">Comparison and annotation settings.</param>
    /// <param name="cancellationToken">Cancels the comparison.</param>
    /// <exception cref="ArgumentException">A PDF has no extractable text, for example a scan without a text layer.</exception>
    public static PdfTextRedlineResult Compare(byte[] oldPdf, byte[] newPdf, PdfTextRedlineOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(oldPdf, nameof(oldPdf));
        Guard.NotNull(newPdf, nameof(newPdf));
        options ??= new PdfTextRedlineOptions();
        options.Validate();

        PdfRedlineText oldText = PdfRedlineText.Read(oldPdf, options.OldReadOptions, nameof(oldPdf), "old", cancellationToken);
        PdfRedlineText newText = PdfRedlineText.Read(newPdf, options.NewReadOptions, nameof(newPdf), "new", cancellationToken);

        var diffOptions = new OfficeDiffOptions {
            Algorithm = options.Algorithm,
            SemanticCleanup = options.SemanticCleanup,
            IgnoreCase = options.IgnoreCase,
            IgnoreWhitespace = true,
            MaxWorkUnits = options.MaxWorkUnits,
            Timeout = options.Timeout
        };
        OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareTokens(oldText.Tokens, newText.Tokens, diffOptions, cancellationToken);
        OfficeChangeSet draft = OfficeChangeSet.FromTextDiff(diff, new OfficeChangeSetOptions {
            DetectMoves = options.DetectMoves,
            MinimumMoveLength = options.MinimumMoveLength
        });

        var changes = new List<OfficeChange>(draft.Changes.Count);
        foreach (OfficeChange change in draft.Changes) {
            changes.Add(new OfficeChange(change.Id, change.Kind, change.OldText, change.NewText,
                oldText.Anchor(change.Old!), newText.Anchor(change.New!)));
        }

        var changeSet = new OfficeChangeSet(changes, options.OldName, options.NewName, "word", draft.IsComplete);
        List<PdfAnnotationCreateOptions> annotations = BuildAnnotations(changeSet, newText, options);
        byte[] output;
        if (annotations.Count == 0) {
            output = (byte[])newPdf.Clone();
        } else {
            output = PdfAnnotationEditor.AddAnnotations(newPdf, annotations, options.NewReadOptions, PdfMutationExecutionPreference.Automatic, cancellationToken).Bytes;
        }

        return new PdfTextRedlineResult(output, changeSet, annotations.Count,
            annotations.Count < CountAnnotatable(changeSet), oldText.PlainText, newText.PlainText);
    }

    /// <summary>Compares two PDF files and returns the annotated new PDF and the change list.</summary>
    public static PdfTextRedlineResult CompareFiles(string oldPath, string newPath, PdfTextRedlineOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(oldPath, nameof(oldPath));
        Guard.NotNullOrWhiteSpace(newPath, nameof(newPath));
        options ??= new PdfTextRedlineOptions();
        if (options.OldName is null || options.NewName is null) {
            options = options.Clone();
            options.OldName ??= Path.GetFileName(oldPath);
            options.NewName ??= Path.GetFileName(newPath);
        }

        return Compare(File.ReadAllBytes(oldPath), File.ReadAllBytes(newPath), options, cancellationToken);
    }

    private static int CountAnnotatable(OfficeChangeSet changes) =>
        changes.Changes.Count(static change => change.New != null && (change.New.Regions.Count > 0 || change.New.PageNumber.HasValue));

    private static List<PdfAnnotationCreateOptions> BuildAnnotations(OfficeChangeSet changes, PdfRedlineText newText, PdfTextRedlineOptions options) {
        var annotations = new List<PdfAnnotationCreateOptions>();
        int annotatedChanges = 0;
        foreach (OfficeChange change in changes.Changes) {
            if (annotatedChanges >= options.MaxAnnotatedChanges) break;
            OfficeChangeAnchor? anchor = change.New;
            if (anchor is null) continue;
            int before = annotations.Count;
            string note = Note(change, options);
            if (change.Kind == OfficeChangeKind.Deleted) {
                PdfRedlineCaret? caret = newText.CaretAt(anchor.Start);
                if (caret is { } point) annotations.Add(Caret(point, note, change, options));
            } else {
                IReadOnlyList<double> color = change.Kind switch {
                    OfficeChangeKind.Inserted => options.InsertionColor,
                    OfficeChangeKind.Moved => options.MoveColor,
                    _ => options.ReplacementColor
                };
                bool first = true;
                foreach (IGrouping<int, OfficeChangeRegion> page in anchor.Regions.GroupBy(static region => region.PageNumber)) {
                    annotations.Add(Highlight(page.Key, page.ToArray(), color, first ? note : note + " (continued)", change, options));
                    first = false;
                }
            }

            if (annotations.Count > before) annotatedChanges++;
        }

        return annotations;
    }

    private static string Note(OfficeChange change, PdfTextRedlineOptions options) => change.Kind switch {
        OfficeChangeKind.Inserted => "Inserted: " + Clip(change.NewText, options),
        OfficeChangeKind.Deleted => "Deleted: " + Clip(change.OldText, options),
        OfficeChangeKind.Moved => "Moved" + (change.Old?.PageNumber is int page ? " from page " + page.ToString(CultureInfo.InvariantCulture) : string.Empty) + ": " + Clip(change.NewText, options),
        _ => "Replaced: " + Clip(change.OldText, options) + "\nWith: " + Clip(change.NewText, options)
    };

    private static string Clip(string text, PdfTextRedlineOptions options) =>
        text.Length <= options.MaxNoteLength ? text : new StringBuilder(options.MaxNoteLength + 1).Append(text, 0, options.MaxNoteLength).Append('\u2026').ToString();

    private static PdfAnnotationCreateOptions Highlight(int pageNumber, OfficeChangeRegion[] regions, IReadOnlyList<double> color, string note, OfficeChange change, PdfTextRedlineOptions options) {
        var quads = new List<double>(regions.Length * 8);
        double left = double.MaxValue, bottom = double.MaxValue, right = double.MinValue, top = double.MinValue;
        foreach (OfficeChangeRegion region in regions) {
            quads.AddRange(new[] { region.Left, region.Top, region.Right, region.Top, region.Left, region.Bottom, region.Right, region.Bottom });
            left = Math.Min(left, region.Left);
            bottom = Math.Min(bottom, region.Bottom);
            right = Math.Max(right, region.Right);
            top = Math.Max(top, region.Top);
        }

        return new PdfAnnotationCreateOptions {
            PageNumber = pageNumber,
            Subtype = "Highlight",
            Rectangle = new[] { left, bottom, right, top },
            QuadPoints = quads,
            Color = color,
            Contents = note,
            Title = options.Author,
            Subject = SubjectOf(change.Kind),
            Name = "officeimo-redline-" + change.Id.ToString(CultureInfo.InvariantCulture) + "-" + pageNumber.ToString(CultureInfo.InvariantCulture),
            CreatePopup = options.CreatePopups,
            PopupRectangle = PopupRectangle(right, top),
            GenerateAppearance = true
        };
    }

    private static PdfAnnotationCreateOptions Caret(PdfRedlineCaret point, string note, OfficeChange change, PdfTextRedlineOptions options) {
        double size = Math.Max(4D, point.FontSize);
        double half = size * 0.3D;
        double bottom = point.Baseline - (size * 0.25D);
        double top = point.Baseline + (size * 0.35D);
        return new PdfAnnotationCreateOptions {
            PageNumber = point.PageNumber,
            Subtype = "Caret",
            Rectangle = new[] { point.X - half, bottom, point.X + half, top },
            Color = options.DeletionColor,
            Contents = note,
            Title = options.Author,
            Subject = SubjectOf(change.Kind),
            Name = "officeimo-redline-" + change.Id.ToString(CultureInfo.InvariantCulture),
            CreatePopup = options.CreatePopups,
            PopupRectangle = PopupRectangle(point.X + half, top),
            GenerateAppearance = true
        };
    }

    private static double[] PopupRectangle(double right, double top) => new[] { right + 8D, top - 96D, right + 208D, top };

    private static string SubjectOf(OfficeChangeKind kind) => kind switch {
        OfficeChangeKind.Inserted => "Insertion",
        OfficeChangeKind.Deleted => "Deletion",
        OfficeChangeKind.Moved => "Move",
        _ => "Replacement"
    };
}

/// <summary>Settings for <see cref="PdfTextRedline"/>.</summary>
public sealed class PdfTextRedlineOptions {
    /// <summary>Word alignment strategy. The default is <see cref="OfficeDiffAlgorithm.Myers"/>.</summary>
    public OfficeDiffAlgorithm Algorithm { get; set; } = OfficeDiffAlgorithm.Myers;

    /// <summary>Merges short unchanged fragments between changes so a reworded phrase is one replacement. The default is <c>true</c>.</summary>
    public bool SemanticCleanup { get; set; } = true;

    /// <summary>Ignores letter case. The default is <c>false</c>.</summary>
    public bool IgnoreCase { get; set; }

    /// <summary>Reports text deleted in one place and inserted unchanged in another as a move. The default is <c>true</c>.</summary>
    public bool DetectMoves { get; set; } = true;

    /// <summary>Shortest moved text, in characters, reported as a move. The default is 12.</summary>
    public int MinimumMoveLength { get; set; } = 12;

    /// <summary>Maximum comparison steps before the remaining text is reported as deleted and inserted (see <see cref="OfficeChangeSet.IsComplete"/>). The default is 50,000,000.</summary>
    public long MaxWorkUnits { get; set; } = 50_000_000;

    /// <summary>Maximum alignment time. Null means no limit.</summary>
    public TimeSpan? Timeout { get; set; }

    /// <summary>Most changes that get annotations; the change list always has every change. The default is 2,000.</summary>
    public int MaxAnnotatedChanges { get; set; } = 2_000;

    /// <summary>Longest text quoted in a note before it is shortened. The default is 500 characters.</summary>
    public int MaxNoteLength { get; set; } = 500;

    /// <summary>Adds a pop-up to each annotation so the note opens next to it. The default is <c>true</c>.</summary>
    public bool CreatePopups { get; set; } = true;

    /// <summary>Author written to each annotation. The default is "OfficeIMO".</summary>
    public string Author { get; set; } = "OfficeIMO";

    /// <summary>Name of the old document in the change list.</summary>
    public string? OldName { get; set; }

    /// <summary>Name of the new document in the change list.</summary>
    public string? NewName { get; set; }

    /// <summary>Highlight colour for inserted text, as RGB components from 0 to 1. The default is green.</summary>
    public IReadOnlyList<double> InsertionColor { get; set; } = new[] { 0.55D, 0.9D, 0.55D };

    /// <summary>Highlight colour for replacement text. The default is yellow.</summary>
    public IReadOnlyList<double> ReplacementColor { get; set; } = new[] { 1D, 0.85D, 0.3D };

    /// <summary>Highlight colour for moved text. The default is light blue.</summary>
    public IReadOnlyList<double> MoveColor { get; set; } = new[] { 0.6D, 0.8D, 1D };

    /// <summary>Caret colour for deletions. The default is red.</summary>
    public IReadOnlyList<double> DeletionColor { get; set; } = new[] { 0.85D, 0.1D, 0.1D };

    /// <summary>Read limits or credentials for the old PDF.</summary>
    public PdfLoadOptions? OldReadOptions { get; set; }

    /// <summary>Read limits or credentials for the new PDF.</summary>
    public PdfLoadOptions? NewReadOptions { get; set; }

    internal PdfTextRedlineOptions Clone() => (PdfTextRedlineOptions)MemberwiseClone();

    internal void Validate() {
        if (MinimumMoveLength < 1) throw new ArgumentOutOfRangeException(nameof(MinimumMoveLength));
        if (MaxAnnotatedChanges < 0) throw new ArgumentOutOfRangeException(nameof(MaxAnnotatedChanges));
        if (MaxNoteLength < 1) throw new ArgumentOutOfRangeException(nameof(MaxNoteLength));
        if (Timeout is TimeSpan timeout && timeout < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(Timeout));
        Guard.NotNull(Author, nameof(Author));
        ValidateColor(InsertionColor, nameof(InsertionColor));
        ValidateColor(ReplacementColor, nameof(ReplacementColor));
        ValidateColor(MoveColor, nameof(MoveColor));
        ValidateColor(DeletionColor, nameof(DeletionColor));
    }

    private static void ValidateColor(IReadOnlyList<double>? color, string name) {
        if (color is null || color.Count != 3 || color.Any(static part => double.IsNaN(part) || part < 0D || part > 1D)) {
            throw new ArgumentException("A colour needs three RGB components from 0 to 1.", name);
        }
    }
}

/// <summary>Result of <see cref="PdfTextRedline.Compare(byte[], byte[], PdfTextRedlineOptions?, CancellationToken)"/>.</summary>
public sealed class PdfTextRedlineResult {
    internal PdfTextRedlineResult(byte[] pdf, OfficeChangeSet changes, int annotationCount, bool annotationLimitReached, string oldText, string newText) {
        Pdf = pdf;
        Changes = changes;
        AnnotationCount = annotationCount;
        AnnotationLimitReached = annotationLimitReached;
        OldText = oldText;
        NewText = newText;
    }

    /// <summary>The new PDF with review annotations added. When nothing changed it is an unmodified copy of the new PDF.</summary>
    public byte[] Pdf { get; }

    /// <summary>Every change, with page numbers and page rectangles on both sides.</summary>
    public OfficeChangeSet Changes { get; }

    /// <summary>Number of annotations added (pop-ups not counted).</summary>
    public int AnnotationCount { get; }

    /// <summary>True when <see cref="PdfTextRedlineOptions.MaxAnnotatedChanges"/> left some changes without annotations.</summary>
    public bool AnnotationLimitReached { get; }

    /// <summary>Old PDF text in reading order, one extracted line per line.</summary>
    public string OldText { get; }

    /// <summary>New PDF text in reading order, one extracted line per line.</summary>
    public string NewText { get; }

    /// <summary>True when the texts are the same.</summary>
    public bool HasChanges => !Changes.IsEmpty;

    /// <summary>Serialises the change list to JSON.</summary>
    public string ToJson(bool indented = true) => Changes.ToJson(indented);

    /// <summary>Renders a standalone side-by-side HTML report of the two texts with the change list.</summary>
    public string ToHtml(string? title = null, CancellationToken cancellationToken = default) =>
        OfficeDiffHtmlReport.SideBySide(OldText, NewText, new OfficeDiffHtmlOptions {
            Title = title ?? "PDF text comparison",
            OldTitle = Changes.OldSource ?? "Old",
            NewTitle = Changes.NewSource ?? "New",
            DiffOptions = new OfficeDiffOptions { IgnoreWhitespace = true },
            Changes = Changes
        }, cancellationToken);
}

internal readonly struct PdfRedlineCaret {
    internal PdfRedlineCaret(int pageNumber, double x, double baseline, double fontSize) {
        PageNumber = pageNumber;
        X = x;
        Baseline = baseline;
        FontSize = fontSize;
    }

    internal int PageNumber { get; }
    internal double X { get; }
    internal double Baseline { get; }
    internal double FontSize { get; }
}

/// <summary>Reading-order word tokens of one PDF with the page geometry of each token.</summary>
internal sealed class PdfRedlineText {
    private readonly List<OfficeTextToken> _tokens = new List<OfficeTextToken>();
    private readonly List<TokenBox?> _boxes = new List<TokenBox?>();
    private readonly StringBuilder _text = new StringBuilder();

    internal IReadOnlyList<OfficeTextToken> Tokens => _tokens;

    internal string PlainText => _text.ToString();

    internal static PdfRedlineText Read(byte[] pdf, PdfLoadOptions? readOptions, string parameterName, string label, CancellationToken cancellationToken) {
        PdfDocumentReadResult document = PdfDocumentReadResult.Load(pdf, null, readOptions);
        var text = new PdfRedlineText();
        int lineKey = 0;
        foreach (PdfLogicalPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (PdfUnderstandingLine line in OrderedLines(page.Analysis)) {
                text.AddLine(page.Analysis.PageNumber, lineKey++, line.Words);
            }
        }

        if (!text._boxes.Any(static box => box.HasValue)) {
            throw new ArgumentException(
                "The " + label + " PDF has no extractable text, so its text cannot be compared. If it is a scan, add a text layer with OCR first.",
                parameterName);
        }

        return text;
    }

    // Reading order first, then any line the reading order did not place (for example table text),
    // so every word on the page is compared exactly once.
    private static IEnumerable<PdfUnderstandingLine> OrderedLines(PdfUnderstandingPageResult analysis) {
        var seen = new HashSet<PdfUnderstandingLine>();
        foreach (PdfUnderstandingRegion region in analysis.ReadingOrder) {
            foreach (PdfUnderstandingLine line in region.Lines) {
                if (seen.Add(line)) yield return line;
            }
        }

        foreach (PdfUnderstandingLine line in analysis.Lines) {
            if (seen.Add(line)) yield return line;
        }
    }

    private void AddLine(int pageNumber, int lineKey, IReadOnlyList<PdfUnderstandingWord> words) {
        bool firstWord = true;
        foreach (PdfUnderstandingWord word in words) {
            if (string.IsNullOrWhiteSpace(word.Text)) continue;
            if (_tokens.Count > 0) AddSeparator(firstWord ? "\n" : " ");
            firstWord = false;
            AddWord(pageNumber, lineKey, word);
        }
    }

    private void AddSeparator(string separator) {
        _tokens.Add(new OfficeTextToken(separator, _text.Length, OfficeTextTokenKind.Whitespace));
        _boxes.Add(null);
        _text.Append(separator);
    }

    private void AddWord(int pageNumber, int lineKey, PdfUnderstandingWord word) {
        string value = word.Text;
        double width = word.XEnd - word.XStart;
        double fontSize = word.FontSize > 0 ? word.FontSize : 10D;
        double bottom = word.BaselineY - (fontSize * 0.22D);
        double top = word.BaselineY + (fontSize * 0.78D);
        if (word.VisualBounds is null && !(word.RotationDegrees > -1D && word.RotationDegrees < 1D) && Math.Abs(word.RotationDegrees) < 359D) {
            // Rotated baselines: keep the word comparable, approximate its box from the baseline origin.
            width = Math.Max(width, word.Advance ?? width);
        }

        // Split the PDF word into word and punctuation tokens so "fox." -> "fox," marks only the comma.
        foreach (OfficeTextToken part in OfficeTextTokenizer.Words(value)) {
            if (part.Kind == OfficeTextTokenKind.Whitespace) {
                AddSeparator(" ");
                continue;
            }

            double left = word.XStart + (width * part.Start / value.Length);
            double right = word.XStart + (width * (part.Start + part.Length) / value.Length);
            _tokens.Add(new OfficeTextToken(part.Text, _text.Length, part.Kind));
            _boxes.Add(new TokenBox(pageNumber, lineKey, left, bottom, right, top, word.BaselineY, fontSize));
            _text.Append(part.Text);
        }
    }

    internal OfficeChangeAnchor Anchor(OfficeChangeAnchor source) {
        var regions = new List<OfficeChangeRegion>();
        TokenBox? current = null;
        for (int index = source.Start; index < source.Start + source.Length && index < _boxes.Count; index++) {
            if (_boxes[index] is not TokenBox box) continue;
            if (current is TokenBox open && open.PageNumber == box.PageNumber && open.LineKey == box.LineKey) {
                current = open.Union(box);
                continue;
            }

            if (current is TokenBox finished) regions.Add(finished.ToRegion());
            current = box;
        }

        if (current is TokenBox last) regions.Add(last.ToRegion());
        int? page = regions.Count > 0 ? regions[0].PageNumber : Nearest(source.Start)?.PageNumber;
        return new OfficeChangeAnchor(source.Start, source.Length, page, null, regions);
    }

    internal PdfRedlineCaret? CaretAt(int index) {
        for (int cursor = Math.Min(index, _boxes.Count) - 1; cursor >= 0; cursor--) {
            if (_boxes[cursor] is TokenBox box) return new PdfRedlineCaret(box.PageNumber, box.Right, box.Baseline, box.FontSize);
            if (_tokens[cursor].Text == "\n") break;
        }

        for (int cursor = Math.Max(0, index); cursor < _boxes.Count; cursor++) {
            if (_boxes[cursor] is TokenBox box) return new PdfRedlineCaret(box.PageNumber, box.Left, box.Baseline, box.FontSize);
        }

        for (int cursor = Math.Min(index, _boxes.Count) - 1; cursor >= 0; cursor--) {
            if (_boxes[cursor] is TokenBox box) return new PdfRedlineCaret(box.PageNumber, box.Right, box.Baseline, box.FontSize);
        }

        return null;
    }

    private TokenBox? Nearest(int index) {
        for (int cursor = Math.Min(index, _boxes.Count - 1); cursor >= 0; cursor--) {
            if (_boxes[cursor] is TokenBox box) return box;
        }

        for (int cursor = Math.Max(0, index); cursor < _boxes.Count; cursor++) {
            if (_boxes[cursor] is TokenBox box) return box;
        }

        return null;
    }

    private readonly struct TokenBox {
        internal TokenBox(int pageNumber, int lineKey, double left, double bottom, double right, double top, double baseline, double fontSize) {
            PageNumber = pageNumber;
            LineKey = lineKey;
            Left = left;
            Bottom = bottom;
            Right = right;
            Top = top;
            Baseline = baseline;
            FontSize = fontSize;
        }

        internal int PageNumber { get; }
        internal int LineKey { get; }
        internal double Left { get; }
        internal double Bottom { get; }
        internal double Right { get; }
        internal double Top { get; }
        internal double Baseline { get; }
        internal double FontSize { get; }

        internal TokenBox Union(TokenBox other) => new TokenBox(PageNumber, LineKey, Math.Min(Left, other.Left), Math.Min(Bottom, other.Bottom),
            Math.Max(Right, other.Right), Math.Max(Top, other.Top), Baseline, Math.Max(FontSize, other.FontSize));

        internal OfficeChangeRegion ToRegion() => new OfficeChangeRegion(PageNumber, Left, Bottom, Right, Top);
    }
}
