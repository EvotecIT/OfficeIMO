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

    internal static OfficeTextIntegrityReport InspectDocumentText(
        ProvenanceOwner owner, string path, OfficeProvenanceAssessmentOptions options,
        long expandedBytesAlreadyInspected, CancellationToken cancellationToken) {
        if (owner != ProvenanceOwner.Word) throw new NotSupportedException("Document text integrity is qualified for Word packages only.");
        long remainingExpanded = options.Structural.MaxExpandedContainerBytes - expandedBytesAlreadyInspected;
        if (remainingExpanded <= 0) throw new InvalidDataException("The provenance assessment exhausted its expanded-data budget.");
        var contentOptions = new OfficeIMO.ContentSafety.OfficeContentSafetyOptions {
            MaxInputBytes = Math.Min(options.TextIntegrity.MaxEncodedBytes, options.Structural.MaxAssetBytes),
            MaxPackageEntries = options.Structural.MaxContainerEntries,
            MaxExpandedPackageBytes = remainingExpanded,
            MaxCharacters = options.TextIntegrity.MaxCharacters,
            MaxFindings = options.TextIntegrity.MaxFindings,
            DetectInstructionLikeText = false,
            IncludeTextIntegrityEvidence = true,
            TextIntegrityOnly = true,
            TextIntegrityOptions = new OfficeTextIntegrityOptions {
                IncludeTypographicSpaces = options.TextIntegrity.IncludeTypographicSpaces,
                IncludeVariationSelectors = options.TextIntegrity.IncludeVariationSelectors,
                // These are decoded document text nodes, not the start of an encoded file.
                IgnoreLeadingByteOrderMark = false
            }
        };
        var content = WordDocument.InspectContentSafety(path, contentOptions, cancellationToken);
        return new OfficeTextIntegrityReport(content.TextIntegrityFindings);
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
