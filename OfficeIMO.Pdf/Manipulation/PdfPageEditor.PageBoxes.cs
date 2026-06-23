namespace OfficeIMO.Pdf;

public static partial class PdfPageEditor {
    /// <summary>
    /// Creates a new PDF with the selected pages updated to the supplied production boundary box.
    /// </summary>
    public static byte[] SetPageBox(byte[] pdf, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        string normalizedBoxName = NormalizePageBoxName(boxName);
        ValidatePageBoxCoordinates(left, bottom, right, top);
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        var selectedPages = pageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : pageNumbers;
        ValidatePageNumbers(selectedPages, document.Pages.Count, nameof(pageNumbers));

        var selected = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(static page => page.ObjectNumber).ToArray();
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selected.Contains(pageNumber)) {
                continue;
            }

            overrides[document.Pages[i].ObjectNumber] = new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                [normalizedBoxName] = CreatePageBoxArray(left, bottom, right, top)
            };
        }

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw));
    }

    /// <summary>Creates a new PDF with the selected pages updated to the supplied production boundary box from a readable stream.</summary>
    public static byte[] SetPageBox(Stream stream, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        return SetPageBox(ReadStream(stream, nameof(stream)), boxName, left, bottom, right, top, pageNumbers);
    }

    /// <summary>Writes a new PDF with the selected pages updated to the supplied production boundary box.</summary>
    public static void SetPageBox(byte[] pdf, Stream outputStream, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        WriteOutput(outputStream, SetPageBox(pdf, boxName, left, bottom, right, top, pageNumbers));
    }

    /// <summary>Writes a new PDF with the selected pages updated to the supplied production boundary box from a readable stream.</summary>
    public static void SetPageBox(Stream inputStream, Stream outputStream, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        WriteOutput(outputStream, SetPageBox(inputStream, boxName, left, bottom, right, top, pageNumbers));
    }

    /// <summary>Writes a new PDF file with the selected pages updated to the supplied production boundary box.</summary>
    public static void SetPageBox(string inputPath, string outputPath, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = SetPageBox(File.ReadAllBytes(inputPath), boxName, left, bottom, right, top, pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>Creates a new PDF file with the selected pages updated to the supplied production boundary box.</summary>
    public static byte[] SetPageBox(string inputPath, string boxName, double left, double bottom, double right, double top, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return SetPageBox(File.ReadAllBytes(inputPath), boxName, left, bottom, right, top, pageNumbers);
    }

    private static PdfArray CreatePageBoxArray(double left, double bottom, double right, double top) {
        var array = new PdfArray();
        array.Items.Add(new PdfNumber(left));
        array.Items.Add(new PdfNumber(bottom));
        array.Items.Add(new PdfNumber(right));
        array.Items.Add(new PdfNumber(top));
        return array;
    }

    private static string NormalizePageBoxName(string boxName) {
        if (string.IsNullOrWhiteSpace(boxName)) {
            throw new ArgumentException("Page box name cannot be empty.", nameof(boxName));
        }

        string normalized = boxName.Trim().TrimStart('/');
        if (string.Equals(normalized, "MediaBox", StringComparison.OrdinalIgnoreCase)) {
            return "MediaBox";
        }

        if (string.Equals(normalized, "CropBox", StringComparison.OrdinalIgnoreCase)) {
            return "CropBox";
        }

        if (string.Equals(normalized, "BleedBox", StringComparison.OrdinalIgnoreCase)) {
            return "BleedBox";
        }

        if (string.Equals(normalized, "TrimBox", StringComparison.OrdinalIgnoreCase)) {
            return "TrimBox";
        }

        if (string.Equals(normalized, "ArtBox", StringComparison.OrdinalIgnoreCase)) {
            return "ArtBox";
        }

        throw new ArgumentOutOfRangeException(nameof(boxName), "Page box name must be MediaBox, CropBox, BleedBox, TrimBox, or ArtBox.");
    }

    private static void ValidatePageBoxCoordinates(double left, double bottom, double right, double top) {
        if (!IsFinite(left) || !IsFinite(bottom) || !IsFinite(right) || !IsFinite(top)) {
            throw new ArgumentOutOfRangeException(nameof(left), "Page box coordinates must be finite numbers.");
        }

        if (right <= left || top <= bottom) {
            throw new ArgumentOutOfRangeException(nameof(right), "Page box right/top coordinates must be greater than left/bottom coordinates.");
        }
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
