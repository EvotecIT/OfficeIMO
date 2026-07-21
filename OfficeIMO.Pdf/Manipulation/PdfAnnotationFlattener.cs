using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

/// <summary>
/// Flattens supported visual PDF annotations into regular page content.
/// </summary>
internal static partial class PdfAnnotationFlattener {
    private const string UnsupportedVisualAnnotationMessage = "Only FreeText, text markup, shape, line, ink, path, stamp, and caret annotations with a normal appearance stream or supported synthesis data can be visually flattened by OfficeIMO.Pdf.";

    /// <summary>
    /// Returns a new PDF with supported visual annotations painted into page content and removed from page annotations.
    /// </summary>
    public static byte[] FlattenVisualAnnotations(byte[] pdf) {
        return FlattenVisualAnnotations(pdf, (PdfAnnotationFlattenOptions?)null);
    }

    /// <summary>Flattens only supported visual annotations matching the supplied selector.</summary>
    public static byte[] FlattenVisualAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options) =>
        FlattenVisualAnnotations(pdf, options, readOptions: null);

    /// <summary>Flattens matching supported visual annotations using explicit read limits or credentials.</summary>
    public static byte[] FlattenVisualAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFlattenOptions(options);
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAnnotations, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        PdfReadDocument read = PdfReadDocument.Open(pdf, readOptions);
        var pageNumbers = new Dictionary<int, int>();
        for (int i = 0; i < read.Pages.Count; i++) pageNumbers[read.Pages[i].ObjectNumber] = i + 1;
        int flattenedCount = FlattenPageVisualAnnotations(objects, ref nextObjectNumber, options, pageNumbers);
        if (flattenedCount == 0) {
            return pdf.ToArray();
        }

        return RewriteAllObjects(objects, catalogObjectNumber, read.UncheckedMetadata);
    }

    private static void ValidateFlattenOptions(PdfAnnotationFlattenOptions? options) {
        if (options?.ObjectNumber <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Annotation object number must be positive.");
        if (options?.PageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Page number must be positive.");
        if (options?.Subtype != null) Guard.NotNullOrWhiteSpace(options.Subtype, nameof(options.Subtype));
    }

    /// <summary>
    /// Returns a new PDF with supported visual annotations flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FlattenVisualAnnotations(Stream stream) {
        return FlattenVisualAnnotations(ReadStream(stream, nameof(stream)));
    }

    /// <summary>
    /// Writes a new PDF with supported visual annotations flattened.
    /// </summary>
    public static void FlattenVisualAnnotations(byte[] pdf, Stream outputStream) {
        WriteOutput(outputStream, FlattenVisualAnnotations(pdf));
    }

    /// <summary>
    /// Writes a new PDF with supported visual annotations flattened from the current position of a readable stream.
    /// </summary>
    public static void FlattenVisualAnnotations(Stream inputStream, Stream outputStream) {
        WriteOutput(outputStream, FlattenVisualAnnotations(inputStream));
    }

    /// <summary>
    /// Writes a new PDF file with supported visual annotations flattened.
    /// </summary>
    public static void FlattenVisualAnnotations(string inputPath, string outputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FlattenVisualAnnotations(File.ReadAllBytes(inputPath));
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with supported visual annotations flattened.
    /// </summary>
    public static void FlattenVisualAnnotations(string inputPath, Stream outputStream) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, FlattenVisualAnnotations(File.ReadAllBytes(inputPath)));
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with supported visual annotations flattened.
    /// </summary>
    public static byte[] FlattenVisualAnnotationsToBytes(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FlattenVisualAnnotations(File.ReadAllBytes(inputPath));
    }
}
