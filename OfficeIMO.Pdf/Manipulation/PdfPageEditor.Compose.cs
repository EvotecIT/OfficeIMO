namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>Creates a PDF from selected one-based pages in caller order, allowing omissions and repetitions.</summary>
    public static byte[] ComposePages(byte[] pdf, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        ValidatePageNumbers(pageNumbers, document.Pages.Count, nameof(pageNumbers), allowDuplicates: true);
        var orderedObjectNumbers = new int[pageNumbers.Length];
        for (int i = 0; i < pageNumbers.Length; i++) {
            orderedObjectNumbers[i] = document.Pages[pageNumbers[i] - 1].ObjectNumber;
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(
            objects,
            document.UncheckedMetadata,
            orderedObjectNumbers,
            catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw),
            fileVersion: fileVersion);
    }

    /// <summary>Creates a PDF from selected pages read from a stream, allowing omissions and repetitions.</summary>
    public static byte[] ComposePages(Stream stream, params int[] pageNumbers) =>
        ComposePages(ReadStream(stream, nameof(stream)), pageNumbers);

    /// <summary>Creates a PDF from selected inclusive page ranges in caller order, allowing omissions and repetitions.</summary>
    public static byte[] ComposePageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        int pageCount = PdfReadDocument.Open(pdf).Pages.Count;
        return ComposePages(pdf, PdfPageRange.ExpandMany(pageRanges, pageCount, nameof(pageRanges)));
    }

    /// <summary>Creates a PDF whose pages are in reverse document order.</summary>
    public static byte[] ReversePages(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        int pageCount = PdfReadDocument.Open(pdf).Pages.Count;
        if (pageCount == 0) {
            throw new ArgumentException("PDF must contain at least one page.", nameof(pdf));
        }

        return ComposePages(pdf, Enumerable.Range(1, pageCount).Reverse().ToArray());
    }

    /// <summary>Creates a PDF whose pages are in reverse document order from a readable stream.</summary>
    public static byte[] ReversePages(Stream stream) => ReversePages(ReadStream(stream, nameof(stream)));

    /// <summary>Writes a PDF whose pages are in reverse document order.</summary>
    public static void ReversePages(string inputPath, string outputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        WriteOutput(ValidateOutputPath(outputPath), ReversePages(File.ReadAllBytes(inputPath)));
    }

    /// <summary>Repeats the selected one-based page sequence the requested number of times.</summary>
    public static byte[] RepeatPages(byte[] pdf, int repetitions, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        if (repetitions < 1) {
            throw new ArgumentOutOfRangeException(nameof(repetitions), repetitions, "Repetitions must be at least one.");
        }

        int pageCount = PdfReadDocument.Open(pdf).Pages.Count;
        ValidatePageNumbers(pageNumbers, pageCount, nameof(pageNumbers), allowDuplicates: true);
        var composed = new int[checked(pageNumbers.Length * repetitions)];
        for (int repetition = 0; repetition < repetitions; repetition++) {
            Array.Copy(pageNumbers, 0, composed, repetition * pageNumbers.Length, pageNumbers.Length);
        }

        return ComposePages(pdf, composed);
    }

    /// <summary>Repeats the selected page sequence from a readable stream.</summary>
    public static byte[] RepeatPages(Stream stream, int repetitions, params int[] pageNumbers) =>
        RepeatPages(ReadStream(stream, nameof(stream)), repetitions, pageNumbers);

    /// <summary>Repeats the selected inclusive page ranges in caller order.</summary>
    public static byte[] RepeatPageRanges(byte[] pdf, int repetitions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        int pageCount = PdfReadDocument.Open(pdf).Pages.Count;
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, pageCount, nameof(pageRanges));
        return RepeatPages(pdf, repetitions, pageNumbers);
    }

    /// <summary>Repeats the selected inclusive page ranges from a readable stream.</summary>
    public static byte[] RepeatPageRanges(Stream stream, int repetitions, params PdfPageRange[] pageRanges) =>
        RepeatPageRanges(ReadStream(stream, nameof(stream)), repetitions, pageRanges);

    /// <summary>
    /// Composes selected ranges in round-robin order: the first page of each range, then the second page of each range, and so on.
    /// Uneven ranges continue until every selected page has been emitted.
    /// </summary>
    public static byte[] InterleavePageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        if (pageRanges.Length < 2) {
            throw new ArgumentException("At least two page ranges are required for interleaving.", nameof(pageRanges));
        }

        int pageCount = PdfReadDocument.Open(pdf).Pages.Count;
        _ = PdfPageRange.ExpandMany(pageRanges, pageCount, nameof(pageRanges));
        int[][] ranges = pageRanges.Select(static range => range.ToPageNumbers()).ToArray();
        int maximumLength = ranges.Max(static range => range.Length);
        var composed = new List<int>(ranges.Sum(static range => range.Length));
        for (int offset = 0; offset < maximumLength; offset++) {
            for (int rangeIndex = 0; rangeIndex < ranges.Length; rangeIndex++) {
                if (offset < ranges[rangeIndex].Length) {
                    composed.Add(ranges[rangeIndex][offset]);
                }
            }
        }

        return ComposePages(pdf, composed.ToArray());
    }

    /// <summary>Interleaves selected inclusive page ranges from a readable stream.</summary>
    public static byte[] InterleavePageRanges(Stream stream, params PdfPageRange[] pageRanges) =>
        InterleavePageRanges(ReadStream(stream, nameof(stream)), pageRanges);
}
