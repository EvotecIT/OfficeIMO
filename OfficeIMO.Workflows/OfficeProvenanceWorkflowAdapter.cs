using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using System.Text;

namespace OfficeIMO.Workflows;

internal static class OfficeProvenanceWorkflowAdapter {
    internal static ProvenanceOwner ResolveByPath(string? path) =>
        Path.GetExtension(path ?? string.Empty).ToLowerInvariant() switch {
            ".docx" or ".docm" or ".dotx" or ".dotm" => ProvenanceOwner.Word,
            ".xlsx" or ".xlsb" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" => ProvenanceOwner.Excel,
            ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" or ".ppam" => ProvenanceOwner.PowerPoint,
            ".vsdx" or ".vsdm" or ".vstx" or ".vstm" or ".vssx" or ".vssm" => ProvenanceOwner.Visio,
            ".odt" or ".ods" or ".odp" or ".odg" or ".ott" or ".ots" or ".otp" or ".otg" => ProvenanceOwner.OpenDocument,
            ".epub" => ProvenanceOwner.Epub,
            ".pdf" => ProvenanceOwner.Pdf,
            ".html" or ".htm" => ProvenanceOwner.Html,
            ".md" or ".markdown" => ProvenanceOwner.Markdown,
            _ => ProvenanceOwner.Core
        };

    internal static ProvenanceOwner Refine(ProvenanceOwner owner, OfficeProvenanceAssetFormat format) {
        if (owner != ProvenanceOwner.Core) return owner;
        return format switch {
            OfficeProvenanceAssetFormat.Pdf => ProvenanceOwner.Pdf,
            OfficeProvenanceAssetFormat.Html => ProvenanceOwner.Html,
            _ => owner
        };
    }

    internal static string GetPackage(ProvenanceOwner owner) => owner switch {
        ProvenanceOwner.Word => "OfficeIMO.Word",
        ProvenanceOwner.Excel => "OfficeIMO.Excel",
        ProvenanceOwner.PowerPoint => "OfficeIMO.PowerPoint",
        ProvenanceOwner.Visio => "OfficeIMO.Visio",
        ProvenanceOwner.OpenDocument => "OfficeIMO.OpenDocument",
        ProvenanceOwner.Epub => "OfficeIMO.Epub",
        ProvenanceOwner.Pdf => "OfficeIMO.Pdf",
        ProvenanceOwner.Html => "OfficeIMO.Html",
        ProvenanceOwner.Markdown => "OfficeIMO.Markdown",
        _ => "OfficeIMO.Core"
    };

    internal static OfficeProvenanceReport Inspect(
        ProvenanceOwner owner,
        string path,
        OfficeProvenanceOptions options,
        string? logicalFilePath = null,
        CancellationToken cancellationToken = default) {
        options.CancellationToken = cancellationToken;
        return owner switch {
            ProvenanceOwner.Word => WordDocument.InspectProvenance(path, options),
            ProvenanceOwner.Excel => ExcelDocument.InspectProvenance(path, options),
            ProvenanceOwner.PowerPoint => PowerPointPresentation.InspectProvenance(path, options),
            ProvenanceOwner.Visio => VisioDocument.InspectProvenance(path, options),
            ProvenanceOwner.OpenDocument => OdfDocument.InspectProvenance(path, options),
            ProvenanceOwner.Epub => EpubDocument.InspectProvenance(path, options),
            ProvenanceOwner.Pdf => PdfProvenance.InspectFile(path, options),
            ProvenanceOwner.Html => logicalFilePath == null
                ? HtmlProvenance.InspectFile(path, options)
                : HtmlProvenance.InspectFile(path, logicalFilePath, options),
            ProvenanceOwner.Markdown => MarkdownProvenance.InspectFile(path, options),
            _ => OfficeProvenanceInspector.InspectFile(path, options)
        };
    }

    internal static OfficeProvenanceRemovalResult Remove(
        ProvenanceOwner owner,
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions options,
        CancellationToken cancellationToken = default) {
        options.Limits.CancellationToken = cancellationToken;
        return owner switch {
            ProvenanceOwner.Word => WordDocument.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.Excel => ExcelDocument.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.PowerPoint => PowerPointPresentation.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.Visio => VisioDocument.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.OpenDocument => OdfDocument.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.Epub => EpubDocument.RemoveProvenance(inputPath, outputPath, options),
            ProvenanceOwner.Pdf => PdfProvenance.RemoveFile(inputPath, outputPath, options),
            ProvenanceOwner.Html => HtmlProvenance.RemoveFile(inputPath, outputPath, options),
            ProvenanceOwner.Markdown => MarkdownProvenance.RemoveFile(inputPath, outputPath, options),
            _ => OfficeProvenanceRemover.RemoveFile(inputPath, outputPath, options)
        };
    }

    internal static Encoding? ResolveTextEncoding(
        ProvenanceOwner owner,
        OfficeProvenanceAssetFormat format,
        string path,
        long maximumBytes,
        CancellationToken cancellationToken) {
        if (owner == ProvenanceOwner.Html) {
            return HtmlProvenance.ResolveTextEncoding(path, maximumBytes, cancellationToken);
        }
        if (owner == ProvenanceOwner.Core && format == OfficeProvenanceAssetFormat.Svg) {
            return OfficeProvenanceXml.ResolveTextEncoding(path, maximumBytes, cancellationToken);
        }
        return null;
    }

    internal static bool SupportsCoreRemoval(OfficeProvenanceAssetFormat format) => format is
        OfficeProvenanceAssetFormat.Jpeg or
        OfficeProvenanceAssetFormat.Png or
        OfficeProvenanceAssetFormat.Webp or
        OfficeProvenanceAssetFormat.Gif or
        OfficeProvenanceAssetFormat.Tiff or
        OfficeProvenanceAssetFormat.Svg or
        OfficeProvenanceAssetFormat.StructuredText or
        OfficeProvenanceAssetFormat.UnstructuredText;

    internal static bool IsTextLike(OfficeProvenanceAssetFormat format) => format is
        OfficeProvenanceAssetFormat.StructuredText or
        OfficeProvenanceAssetFormat.UnstructuredText or
        OfficeProvenanceAssetFormat.Html or
        OfficeProvenanceAssetFormat.Svg;

    internal enum ProvenanceOwner {
        Core,
        Word,
        Excel,
        PowerPoint,
        Visio,
        OpenDocument,
        Epub,
        Pdf,
        Html,
        Markdown
    }
}
