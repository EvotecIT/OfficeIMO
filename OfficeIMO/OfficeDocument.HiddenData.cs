using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeDocument {
    /// <summary>
    /// Removes the selected hidden data from a document file and atomically writes the cleaned copy to <paramref name="outputPath"/>.
    /// The input is never changed. See <see cref="RemoveHiddenData(byte[], string, OfficeHiddenDataOptions?, OfficeSignatureMutationPolicy)"/>.
    /// </summary>
#pragma warning disable RS0026 // File and bytes overloads, as RemoveProvenance has.
    public static OfficeHiddenDataReport RemoveHiddenData(string inputPath, string outputPath, OfficeHiddenDataOptions? options = null,
        OfficeSignatureMutationPolicy signatureMutationPolicy = OfficeSignatureMutationPolicy.BlockSave) {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        string input = Path.GetFullPath(inputPath);
        string output = Path.GetFullPath(outputPath);
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("The output must differ from the input.", nameof(outputPath));
        if (!string.Equals(Path.GetExtension(input), Path.GetExtension(output), StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("The output must keep the input's extension.", nameof(outputPath));
        }
        OfficeHiddenDataResult result = RemoveHiddenData(File.ReadAllBytes(input), input, options, signatureMutationPolicy);
        OfficeFileCommit.WriteAllBytes(output, result.ToArray());
        return result.Report;
    }

    /// <summary>
    /// Removes the selected hidden data from encoded Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, JPEG, PNG, WebP, GIF or TIFF bytes
    /// without touching the filesystem. <paramref name="fileName"/> chooses the format by extension. Use the document's own
    /// <c>RemoveHiddenData</c> to clean a document you already have open, and the format's static <c>RemoveProvenance</c> to remove only provenance
    /// while leaving every other byte alone.
    /// </summary>
    /// <param name="data">The encoded document.</param>
    /// <param name="fileName">A file name whose extension identifies the format.</param>
    /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>.</param>
    /// <param name="signatureMutationPolicy">How the save treats digital signatures that cleaning invalidates. Defaults to blocking the save.</param>
    /// <returns>The report and the cleaned bytes.</returns>
    public static OfficeHiddenDataResult RemoveHiddenData(byte[] data, string fileName, OfficeHiddenDataOptions? options = null,
        OfficeSignatureMutationPolicy signatureMutationPolicy = OfficeSignatureMutationPolicy.BlockSave) {
#pragma warning restore RS0026
        ArgumentNullException.ThrowIfNull(data);
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        HiddenDataFormat format = GetHiddenDataFormat(fileName);
        options ??= new OfficeHiddenDataOptions();
        if (format == HiddenDataFormat.Image) {
            var collector = new OfficeHiddenDataCollector("Image", options);
            byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(data, Path.GetFileName(fileName), collector);
            return new OfficeHiddenDataResult(collector.Build(), replacement ?? (byte[])data.Clone());
        }
        using var cleaned = new MemoryStream();
        OfficeHiddenDataReport report;
        using (var stream = new MemoryStream(data, writable: false)) {
            switch (format) {
                case HiddenDataFormat.Word:
                    using (WordDocument document = WordDocument.Load(stream)) {
                        report = document.RemoveHiddenData(options);
                        document.Save(cleaned, new WordSaveOptions { SignatureMutationPolicy = signatureMutationPolicy });
                    }
                    break;
                case HiddenDataFormat.Excel:
                    using (ExcelDocument document = ExcelDocument.Load(stream)) {
                        report = document.RemoveHiddenData(options);
                        document.Save(cleaned, new ExcelSaveOptions { SignatureMutationPolicy = signatureMutationPolicy });
                    }
                    break;
                case HiddenDataFormat.PowerPoint:
                    using (PowerPointDocument document = PowerPointDocument.Load(stream)) {
                        report = document.RemoveHiddenData(options);
                        document.Save(cleaned, new PowerPointSaveOptions { SignatureMutationPolicy = signatureMutationPolicy });
                    }
                    break;
                case HiddenDataFormat.Visio:
                    using (VisioDocument document = VisioDocument.Load(stream)) {
                        report = document.RemoveHiddenData(options);
                        document.Save(cleaned);
                    }
                    break;
                case HiddenDataFormat.OpenDocument:
                    using (OdfDocument document = OdfDocument.Load(stream)) {
                        report = document.RemoveHiddenData(options);
                        document.Save(cleaned);
                    }
                    break;
                default: {
                    EpubPublication publication = EpubPublication.Load(stream);
                    report = publication.RemoveHiddenData(options);
                    publication.Save(cleaned);
                    break;
                }
            }
        }
        return new OfficeHiddenDataResult(report, cleaned.ToArray());
    }

    /// <summary>Whether hidden-data inspection and removal support a file name's extension.</summary>
    public static bool SupportsHiddenData(string path) => TryGetHiddenDataFormat(path, out _);

    internal static OfficeHiddenDataReport InspectHiddenData(byte[] data, string sourceName) {
        HiddenDataFormat format = GetHiddenDataFormat(sourceName);
        if (format == HiddenDataFormat.Image) {
            var collector = new OfficeHiddenDataCollector("Image", null);
            OfficeHiddenDataAssets.ProcessImage(data, Path.GetFileName(sourceName), collector);
            return collector.Build();
        }
        using var stream = new MemoryStream(data, writable: false);
        switch (format) {
            case HiddenDataFormat.Word:
                using (WordDocument document = WordDocument.Load(stream)) return document.InspectHiddenData();
            case HiddenDataFormat.Excel:
                using (ExcelDocument document = ExcelDocument.Load(stream)) return document.InspectHiddenData();
            case HiddenDataFormat.PowerPoint:
                using (PowerPointDocument document = PowerPointDocument.Load(stream)) return document.InspectHiddenData();
            case HiddenDataFormat.Visio:
                using (VisioDocument document = VisioDocument.Load(stream)) return document.InspectHiddenData();
            case HiddenDataFormat.OpenDocument:
                using (OdfDocument document = OdfDocument.Load(stream)) return document.InspectHiddenData();
            default:
                return EpubPublication.Load(stream).InspectHiddenData();
        }
    }

    internal static OfficeHiddenDataReport InspectLoadedHiddenData(object document) => document switch {
        WordDocument word => word.InspectHiddenData(),
        ExcelDocument excel => excel.InspectHiddenData(),
        PowerPointDocument presentation => presentation.InspectHiddenData(),
        VisioDocument visio => visio.InspectHiddenData(),
        OdfDocument odf => odf.InspectHiddenData(),
        EpubPublication epub => epub.InspectHiddenData(),
        _ => throw new NotSupportedException("Hidden-data inspection is not available for " + document.GetType().FullName + ".")
    };

    private enum HiddenDataFormat { Word, Excel, PowerPoint, Visio, OpenDocument, Epub, Image }

    private static HiddenDataFormat GetHiddenDataFormat(string path) => TryGetHiddenDataFormat(path, out HiddenDataFormat format)
        ? format
        : throw new NotSupportedException("Hidden-data inspection supports Word, Excel, PowerPoint, Visio, OpenDocument, EPUB and JPEG, PNG, WebP, GIF and TIFF image files.");

    private static bool TryGetHiddenDataFormat(string path, out HiddenDataFormat format) {
        HiddenDataFormat? resolved = Path.GetExtension(path).ToLowerInvariant() switch {
            ".docx" or ".docm" or ".dotx" or ".dotm" => HiddenDataFormat.Word,
            ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => HiddenDataFormat.Excel,
            ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => HiddenDataFormat.PowerPoint,
            ".vsdx" or ".vsdm" or ".vstx" or ".vstm" or ".vssx" or ".vssm" => HiddenDataFormat.Visio,
            ".odt" or ".ods" or ".odp" or ".odg" or ".ott" or ".ots" or ".otp" or ".otg" => HiddenDataFormat.OpenDocument,
            ".epub" => HiddenDataFormat.Epub,
            ".jpg" or ".jpeg" or ".png" or ".webp" or ".gif" or ".tif" or ".tiff" => HiddenDataFormat.Image,
            _ => null
        };
        format = resolved ?? default;
        return resolved.HasValue;
    }
}
