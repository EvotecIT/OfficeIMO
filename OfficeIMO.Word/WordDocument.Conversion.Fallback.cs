using OfficeIMO.Drawing;
using OfficeIMO.Word.LegacyDoc.Write;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private const string LegacyVisualFallbackArtifact = "DOC Data stream:inline-page-images";

    private sealed class WordLegacyVisualFallbackPlan {
        internal WordLegacyVisualFallbackPlan(
            IReadOnlyList<OfficeImageExportResult> pages,
            bool embedSource,
            string nativeFailure) {
            Pages = pages;
            EmbedSource = embedSource;
            NativeFailure = nativeFailure;
        }

        internal IReadOnlyList<OfficeImageExportResult> Pages { get; }
        internal bool EmbedSource { get; }
        internal string NativeFailure { get; }
    }

    private static WordLegacyVisualFallbackPlan? PlanLegacyVisualFallback(
        WordDocument document,
        OfficeFormatDescriptor destinationFormat,
        OfficeCompatibilityMode mode,
        WordDocumentConversionOptions options,
        List<OfficeConversionDiagnostic> diagnostics) {
        if (destinationFormat.Generation != OfficeFormatGeneration.Legacy) return null;

        try {
            _ = LegacyDocWriter.WriteDocument(
                document,
                out OfficeConversionReport nativeReport,
                CreateLegacyWriterProbeOptions(options.SaveOptions),
                isTemplate: destinationFormat.DocumentKind == OfficeDocumentKind.Template);
            bool allowsPaletteLoss = AllowsLoss(options, mode);
            foreach (OfficeDiagnostic finding in nativeReport.Diagnostics) {
                diagnostics.Add(new OfficeConversionDiagnostic(finding.Code,
                    OfficeConversionDiagnosticCategory.DataLoss,
                    allowsPaletteLoss ? OfficeConversionDiagnosticSeverity.Warning : OfficeConversionDiagnosticSeverity.Error,
                    finding.Message, representsDataLoss: true,
                    allowsPaletteLoss ? OfficeCompatibilityState.Approximated : OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Visual, sourceLocation: finding.Location));
            }
            return null;
        } catch (NotSupportedException exception) {
            bool permitsVisualFallback = mode is OfficeCompatibilityMode.PreferVisual
                or OfficeCompatibilityMode.BestEffort
                or OfficeCompatibilityMode.PreservationOnly;
            if (!permitsVisualFallback) {
                diagnostics.Add(new OfficeConversionDiagnostic(
                    "word-legacy-writer-unsupported",
                    OfficeConversionDiagnosticCategory.DestinationFormat,
                    OfficeConversionDiagnosticSeverity.Error,
                    exception.Message,
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    GetLegacyWriterFailureImpact(document)));
                return null;
            }

            IReadOnlyList<OfficeImageExportResult> pages;
            try {
                pages = (options.VisualFallbackRenderer ?? throw new NotSupportedException(
                    "Visual DOC fallback requires the Word PDF adapter through OfficeConverter."))(document);
            } catch (OfficeImageExportPolicyException renderException) {
                string codes = string.Join(", ", renderException.Diagnostics
                    .Select(item => item.Code)
                    .Distinct(StringComparer.Ordinal)
                    .Take(8));
                diagnostics.Add(new OfficeConversionDiagnostic(
                    "word-legacy-writer-visual-fallback-unavailable",
                    OfficeConversionDiagnosticCategory.DestinationFormat,
                    OfficeConversionDiagnosticSeverity.Error,
                    $"Native DOC output is unsupported and the visual fallback has renderer omissions ({codes}). Native writer: {exception.Message}",
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Semantic));
                return null;
            }

            if (pages.Count == 0) {
                diagnostics.Add(new OfficeConversionDiagnostic(
                    "word-legacy-writer-visual-fallback-unavailable",
                    OfficeConversionDiagnosticCategory.DestinationFormat,
                    OfficeConversionDiagnosticSeverity.Error,
                    "Native DOC output is unsupported and the visual fallback did not render any pages. Native writer: " + exception.Message,
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Semantic));
                return null;
            }

            diagnostics.Add(new OfficeConversionDiagnostic(
                "word-legacy-writer-visual-fallback",
                OfficeConversionDiagnosticCategory.DataLoss,
                OfficeConversionDiagnosticSeverity.Warning,
                $"The editable document is represented by {pages.Count} static page image(s) because native DOC writing rejected part of the source model. {exception.Message}",
                representsDataLoss: true,
                OfficeCompatibilityState.Rasterized,
                GetLegacyWriterFailureImpact(document) | OfficeCompatibilityImpact.Editability,
                fallbackArtifact: LegacyVisualFallbackArtifact));

            bool embedSource = options.EmbedSourceWhenLossy
                || mode == OfficeCompatibilityMode.PreservationOnly;
            AddWordSourceCarrierDiagnostic(diagnostics, embedSource, document.HasMacros);

            return new WordLegacyVisualFallbackPlan(pages, embedSource, exception.Message);
        }
    }

    private static WordSaveOptions CreateLegacyWriterProbeOptions(WordSaveOptions? source) => new() {
        LossPolicy = OfficeConversionLossPolicy.Allow,
        SignedDocumentPolicy = source?.SignedDocumentPolicy ?? WordSignedDocumentSavePolicy.Block
    };

    private static OfficeCompatibilityImpact GetLegacyWriterFailureImpact(WordDocument document) {
        OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Editability;
        if (document.Charts.Count > 0 || document.SmartArts.Count > 0 || document.Shapes.Count > 0) {
            impact |= OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Semantic;
        }
        if (document.HasMacros) {
            impact |= OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier;
        }
        return impact;
    }

    private static byte[] CreateLegacyVisualFallbackBytes(
        WordLegacyVisualFallbackPlan plan,
        OfficeFormatDescriptor sourceFormat,
        OfficeFormatDescriptor destinationFormat,
        OfficeCompatibilityMode mode,
        string sourcePath,
        byte[]? sourceBytes) {
        using WordDocument fallback = WordDocument.Create();
        OfficeImageExportResult firstPage = plan.Pages[0];
        WordSection section = fallback.Sections[0];
        section.PageSettings.Width = global::OfficeIMO.OfficeLength.Points((checked((uint)Math.Max(1, firstPage.Width * 20))) / 20D);
        section.PageSettings.Height = global::OfficeIMO.OfficeLength.Points((checked((uint)Math.Max(1, firstPage.Height * 20))) / 20D);
        section.Margins.Left = global::OfficeIMO.OfficeLength.Zero;
        section.Margins.Right = global::OfficeIMO.OfficeLength.Zero;
        section.Margins.Top = global::OfficeIMO.OfficeLength.Zero;
        section.Margins.Bottom = global::OfficeIMO.OfficeLength.Zero;
        section.Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Zero;
        section.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Zero;

        for (int index = 0; index < plan.Pages.Count; index++) {
            OfficeImageExportResult page = plan.Pages[index];
            using var image = new MemoryStream(page.Bytes, writable: false);
            fallback.AddParagraph().AddImage(
                image,
                $"compatibility-page-{index + 1:D4}.png",
                OfficeLength.Points(page.Width),
                OfficeLength.Points(page.Height),
                new WordImageOptions { Description = $"Static visual compatibility fallback page {index + 1}" });
            if (index + 1 < plan.Pages.Count) fallback.AddPageBreak();
        }

        byte[] bytes = LegacyDocWriter.WriteDocument(
            fallback,
            new WordSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow },
            isTemplate: destinationFormat.DocumentKind == OfficeDocumentKind.Template);
        if (!plan.EmbedSource) return bytes;
        if (sourceBytes == null) throw new InvalidOperationException("Embedded-source fallback requires source bytes.");
        return AttachWordSourceCarrier(
            bytes,
            destinationFormat,
            sourceFormat.Id,
            Path.GetFileName(sourcePath),
            mode,
            sourceBytes);
    }

    private static byte[] AttachWordSourceCarrier(
        byte[] destinationBytes,
        OfficeFormatDescriptor destinationFormat,
        string sourceFormatId,
        string sourceFileName,
        OfficeCompatibilityMode mode,
        byte[] sourceBytes) => destinationFormat.Encoding == OfficeFormatEncoding.CompoundBinary
        ? OfficeCompatibilitySourceCarrier.AttachToCompound(
            destinationBytes,
            sourceFormatId,
            sourceFileName,
            mode,
            sourceBytes)
        : OfficeCompatibilitySourceCarrier.AttachToPackage(
            destinationBytes,
            sourceFormatId,
            sourceFileName,
            mode,
            sourceBytes);

    private static void AddWordSourceCarrierDiagnostic(
        List<OfficeConversionDiagnostic> diagnostics,
        bool embedded,
        bool hasMacros) {
        diagnostics.Add(new OfficeConversionDiagnostic(
            embedded ? "word-source-carrier-embedded" : "word-source-carrier-not-embedded",
            OfficeConversionDiagnosticCategory.DataLoss,
            OfficeConversionDiagnosticSeverity.Warning,
            embedded
                ? "The complete original source is retained in an inert, hash-verified OfficeIMO compatibility carrier. It is not executable or editable through the converted document model."
                : "The complete original source is not retained. Set EmbedSourceWhenLossy when deliberate byte-level recovery is required.",
            representsDataLoss: !embedded,
            embedded ? OfficeCompatibilityState.EmbeddedSource : OfficeCompatibilityState.Dropped,
            OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability
                | (hasMacros
                    ? OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Behavioral
                    : OfficeCompatibilityImpact.None),
            fallbackArtifact: embedded ? OfficeCompatibilitySourceCarrier.PayloadPath : null));
    }
}
