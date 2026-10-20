using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>Deletes pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] DeletePages(Stream inputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        DeletePagesWithReadOptions(ReadStream(inputStream, nameof(inputStream), readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes delete output after bounded input processing; caller streams remain open.</summary>
    internal static void DeletePages(Stream inputStream, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, DeletePages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes delete output after bounded input processing; caller streams remain open.</summary>
    internal static void DeletePages(Stream inputStream, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, DeletePages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Deletes pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] DeletePages(string inputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        DeletePagesWithReadOptions(ReadPath(inputPath, readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes delete output after bounded input processing; caller streams remain open.</summary>
    internal static void DeletePages(string inputPath, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, DeletePages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes delete output after bounded input processing; caller streams remain open.</summary>
    internal static void DeletePages(string inputPath, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, DeletePages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Reorders pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] ReorderPages(Stream inputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        ReorderPagesWithReadOptions(ReadStream(inputStream, nameof(inputStream), readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes reorder output after bounded input processing; caller streams remain open.</summary>
    internal static void ReorderPages(Stream inputStream, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, ReorderPages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes reorder output after bounded input processing; caller streams remain open.</summary>
    internal static void ReorderPages(Stream inputStream, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, ReorderPages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Reorders pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] ReorderPages(string inputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        ReorderPagesWithReadOptions(ReadPath(inputPath, readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes reorder output after bounded input processing; caller streams remain open.</summary>
    internal static void ReorderPages(string inputPath, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, ReorderPages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes reorder output after bounded input processing; caller streams remain open.</summary>
    internal static void ReorderPages(string inputPath, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, ReorderPages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Duplicates pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] DuplicatePages(Stream inputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        DuplicatePagesWithReadOptions(ReadStream(inputStream, nameof(inputStream), readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes duplicate output after bounded input processing; caller streams remain open.</summary>
    internal static void DuplicatePages(Stream inputStream, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, DuplicatePages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes duplicate output after bounded input processing; caller streams remain open.</summary>
    internal static void DuplicatePages(Stream inputStream, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, DuplicatePages(inputStream, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Duplicates pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] DuplicatePages(string inputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        DuplicatePagesWithReadOptions(ReadPath(inputPath, readOptions, cancellationToken), readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes duplicate output after bounded input processing; caller streams remain open.</summary>
    internal static void DuplicatePages(string inputPath, Stream outputStream, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, DuplicatePages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes duplicate output after bounded input processing; caller streams remain open.</summary>
    internal static void DuplicatePages(string inputPath, string outputPath, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, DuplicatePages(inputPath, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Moves pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] MovePages(Stream inputStream, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        MovePagesWithReadOptions(ReadStream(inputStream, nameof(inputStream), readOptions, cancellationToken), insertBeforePageNumber, readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes move output after bounded input processing; caller streams remain open.</summary>
    internal static void MovePages(Stream inputStream, Stream outputStream, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, MovePages(inputStream, insertBeforePageNumber, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes move output after bounded input processing; caller streams remain open.</summary>
    internal static void MovePages(Stream inputStream, string outputPath, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, MovePages(inputStream, insertBeforePageNumber, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Moves pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] MovePages(string inputPath, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        MovePagesWithReadOptions(ReadPath(inputPath, readOptions, cancellationToken), insertBeforePageNumber, readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes move output after bounded input processing; caller streams remain open.</summary>
    internal static void MovePages(string inputPath, Stream outputStream, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, MovePages(inputPath, insertBeforePageNumber, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes move output after bounded input processing; caller streams remain open.</summary>
    internal static void MovePages(string inputPath, string outputPath, int insertBeforePageNumber, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, MovePages(inputPath, insertBeforePageNumber, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Rotates pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] RotatePages(Stream inputStream, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        RotatePagesWithReadOptions(ReadStream(inputStream, nameof(inputStream), readOptions, cancellationToken), rotationDegrees, readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes rotate output after bounded input processing; caller streams remain open.</summary>
    internal static void RotatePages(Stream inputStream, Stream outputStream, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, RotatePages(inputStream, rotationDegrees, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes rotate output after bounded input processing; caller streams remain open.</summary>
    internal static void RotatePages(Stream inputStream, string outputPath, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, RotatePages(inputStream, rotationDegrees, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Rotates pages using bounded input admission and cooperative cancellation.</summary>
    internal static byte[] RotatePages(string inputPath, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        RotatePagesWithReadOptions(ReadPath(inputPath, readOptions, cancellationToken), rotationDegrees, readOptions, cancellationToken, pageNumbers);

    /// <summary>Writes rotate output after bounded input processing; caller streams remain open.</summary>
    internal static void RotatePages(string inputPath, Stream outputStream, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputStream, RotatePages(inputPath, rotationDegrees, readOptions, cancellationToken, pageNumbers), cancellationToken);

    /// <summary>Writes rotate output after bounded input processing; caller streams remain open.</summary>
    internal static void RotatePages(string inputPath, string outputPath, int rotationDegrees, PdfLoadOptions? readOptions, CancellationToken cancellationToken, params int[] pageNumbers) =>
        WriteOutput(outputPath, RotatePages(inputPath, rotationDegrees, readOptions, cancellationToken, pageNumbers), cancellationToken);

}
