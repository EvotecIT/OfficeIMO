using OfficeIMO.Drawing;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private const string ExcelCellRasterFallbackArtifact = "worksheet:palette-quantized-cell-raster";

    private sealed class ExcelVisualFallbackSheet {
        internal ExcelVisualFallbackSheet(string name, bool hidden, OfficeImageExportResult image) {
            Name = name;
            Hidden = hidden;
            Image = image;
        }

        internal string Name { get; }
        internal bool Hidden { get; }
        internal OfficeImageExportResult Image { get; }
    }

    private sealed class ExcelVisualFallbackPlan {
        internal ExcelVisualFallbackPlan(
            IReadOnlyList<ExcelVisualFallbackSheet> sheets,
            bool embedSource,
            string nativeFailure) {
            Sheets = sheets;
            EmbedSource = embedSource;
            NativeFailure = nativeFailure;
        }

        internal IReadOnlyList<ExcelVisualFallbackSheet> Sheets { get; }
        internal bool EmbedSource { get; }
        internal string NativeFailure { get; }
    }

    private static ExcelVisualFallbackPlan? PlanExcelVisualFallback(
        ExcelDocument document,
        OfficeFormatDescriptor destinationFormat,
        OfficeCompatibilityMode mode,
        ExcelDocumentConversionOptions options,
        List<ExcelConversionDiagnostic> diagnostics) {
        if (!IsBoundedNativeBinaryDestination(destinationFormat)) return null;
        ValidateVisualFallbackDimensions(options);

        try {
            ProbeNativeBinaryWriter(document, destinationFormat);
            return null;
        } catch (NotSupportedException exception) {
            bool permitsVisualFallback = mode is OfficeCompatibilityMode.PreferVisual
                or OfficeCompatibilityMode.BestEffort
                or OfficeCompatibilityMode.PreservationOnly;
            if (!permitsVisualFallback) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.BinaryWriter.Unsupported",
                    ExcelConversionDiagnosticCategory.DestinationFormat,
                    ExcelConversionDiagnosticSeverity.Error,
                    exception.Message,
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual
                        | OfficeCompatibilityImpact.Editability));
                return null;
            }

            IReadOnlyList<ExcelVisualFallbackSheet> sheets;
            try {
                var rendered = new List<ExcelVisualFallbackSheet>();
                foreach (ExcelSheet sheet in document.Sheets) {
                    OfficeImageExportResult image = sheet.ExportImage(
                        OfficeImageExportFormat.Png,
                        new ExcelWorksheetImageExportOptions {
                            BackgroundColor = OfficeColor.White,
                            ConditionalFormattingDate = new DateTime(2000, 1, 1),
                            HeaderFooterDateTime = new DateTime(2000, 1, 1),
                            Policy = new OfficeImageExportPolicy {
                                RequireNoOmissions = true,
                                RequireNoFailures = true
                            }
                        });
                    rendered.Add(new ExcelVisualFallbackSheet(sheet.Name, sheet.Hidden, image));
                }
                sheets = rendered.AsReadOnly();
            } catch (OfficeImageExportPolicyException renderException) {
                string codes = string.Join(", ", renderException.Diagnostics
                    .Select(item => item.Code)
                    .Distinct(StringComparer.Ordinal)
                    .Take(8));
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.BinaryWriter.VisualFallbackUnavailable",
                    ExcelConversionDiagnosticCategory.DestinationFormat,
                    ExcelConversionDiagnosticSeverity.Error,
                    $"Native binary output is unsupported and the worksheet renderer has omissions ({codes}). Native writer: {exception.Message}",
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual));
                return null;
            }

            if (sheets.Count == 0) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.BinaryWriter.VisualFallbackUnavailable",
                    ExcelConversionDiagnosticCategory.DestinationFormat,
                    ExcelConversionDiagnosticSeverity.Error,
                    "Native binary output is unsupported and the visual fallback found no worksheets to render. Native writer: " + exception.Message,
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual));
                return null;
            }

            diagnostics.Add(new ExcelConversionDiagnostic(
                "Excel.BinaryWriter.CellRasterFallback",
                ExcelConversionDiagnosticCategory.DataLoss,
                ExcelConversionDiagnosticSeverity.Warning,
                $"The workbook is represented by {sheets.Count} palette-quantized cell-raster worksheet(s) because the native binary writer rejected part of the editable model. {exception.Message}",
                representsDataLoss: true,
                OfficeCompatibilityState.Rasterized,
                OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual
                    | OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Editability,
                fallbackArtifact: ExcelCellRasterFallbackArtifact));

            bool embedSource = options.EmbedSourceWhenLossy
                || mode == OfficeCompatibilityMode.PreservationOnly;
            AddExcelSourceCarrierDiagnostic(
                diagnostics,
                embedSource,
                document.HasMacros,
                visualFallback: true);
            return new ExcelVisualFallbackPlan(sheets, embedSource, exception.Message);
        }
    }

    private static bool IsBoundedNativeBinaryDestination(OfficeFormatDescriptor destinationFormat) =>
        string.Equals(destinationFormat.Extension, ".xls", StringComparison.Ordinal)
        || string.Equals(destinationFormat.Extension, ".xlsb", StringComparison.Ordinal);

    private static void ProbeNativeBinaryWriter(
        ExcelDocument document,
        OfficeFormatDescriptor destinationFormat) {
        if (string.Equals(destinationFormat.Extension, ".xls", StringComparison.Ordinal)) {
            _ = LegacyXlsWriter.WriteWorkbook(document);
            return;
        }

        using var output = new MemoryStream();
        XlsbNewPackageWriter.Write(document, output);
    }

    private static void ValidateVisualFallbackDimensions(ExcelDocumentConversionOptions options) {
        if (options.VisualFallbackMaxColumns <= 0 || options.VisualFallbackMaxColumns > 256) {
            throw new ArgumentOutOfRangeException(
                nameof(options.VisualFallbackMaxColumns),
                "The XLS-compatible visual fallback must use between 1 and 256 columns.");
        }
        if (options.VisualFallbackMaxRows <= 0 || options.VisualFallbackMaxRows > 65536) {
            throw new ArgumentOutOfRangeException(
                nameof(options.VisualFallbackMaxRows),
                "The XLS-compatible visual fallback must use between 1 and 65,536 rows.");
        }
    }

    private static byte[] CreateExcelVisualFallbackBytes(
        ExcelVisualFallbackPlan plan,
        OfficeFormatDescriptor sourceFormat,
        OfficeFormatDescriptor destinationFormat,
        OfficeCompatibilityMode mode,
        ExcelDocumentConversionOptions options,
        string sourcePath,
        byte[]? sourceBytes) {
        using ExcelDocument fallback = ExcelDocument.Create();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ExcelVisualFallbackSheet sourceSheet in plan.Sheets) {
            string name = CreateFallbackSheetName(sourceSheet.Name, usedNames);
            ExcelSheet sheet = fallback.AddWorksheet(name);
            sheet.SetHidden(sourceSheet.Hidden);
            if (!OfficePngReader.TryDecode(sourceSheet.Image.Bytes, out OfficeRasterImage? raster)
                || raster == null) {
                throw new InvalidDataException($"The visual fallback could not decode rendered worksheet '{sourceSheet.Name}'.");
            }

            OfficeRasterImage sampled = ResizeForCellRaster(
                raster,
                options.VisualFallbackMaxColumns,
                options.VisualFallbackMaxRows);
            sheet.ApplyCompatibilityCellRaster(QuantizeToBiffPalette(sampled), sampled.Width, sampled.Height);
        }

        ExcelFileFormat target = string.Equals(destinationFormat.Extension, ".xls", StringComparison.Ordinal)
            ? ExcelFileFormat.Xls
            : ExcelFileFormat.Xlsb;
        byte[] bytes = fallback.ToBytes(
            target,
            new ExcelSaveOptions { LossPolicy = ExcelConversionLossPolicy.Allow });
        if (!plan.EmbedSource) return bytes;
        if (sourceBytes == null) throw new InvalidOperationException("Embedded-source fallback requires source bytes.");
        return AttachExcelSourceCarrier(
            bytes,
            destinationFormat,
            sourceFormat.Id,
            Path.GetFileName(sourcePath),
            mode,
            sourceBytes);
    }

    private static byte[] AttachExcelSourceCarrier(
        byte[] destinationBytes,
        OfficeFormatDescriptor destinationFormat,
        string sourceFormatId,
        string sourceFileName,
        OfficeCompatibilityMode mode,
        byte[] sourceBytes) => destinationFormat.Encoding == OfficeFormatEncoding.CompoundBinary
        ? OfficeCompatibilitySourceCarrier.AttachToCompound(
            destinationBytes, sourceFormatId, sourceFileName, mode, sourceBytes)
        : OfficeCompatibilitySourceCarrier.AttachToPackage(
            destinationBytes, sourceFormatId, sourceFileName, mode, sourceBytes);

    private static OfficeRasterImage ResizeForCellRaster(
        OfficeRasterImage source,
        int maxColumns,
        int maxRows) {
        double scale = Math.Min(1D, Math.Min(
            maxColumns / (double)source.Width,
            maxRows / (double)source.Height));
        int width = Math.Max(1, (int)Math.Round(source.Width * scale));
        int height = Math.Max(1, (int)Math.Round(source.Height * scale));
        return OfficeRasterResampler.Resize(source, width, height, OfficeRasterResamplingMode.Bilinear);
    }

    private static IReadOnlyList<string> QuantizeToBiffPalette(OfficeRasterImage image) {
        string[] palette = BiffColorPalette.DefaultPaletteColors
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        OfficeColor[] colors = palette.Select(ParseArgb).ToArray();
        byte[] pixels = image.GetPixels();
        var result = new string[image.Width * image.Height];
        for (int index = 0; index < result.Length; index++) {
            int offset = index * 4;
            int alpha = pixels[offset + 3];
            int red = ((pixels[offset] * alpha) + (255 * (255 - alpha))) / 255;
            int green = ((pixels[offset + 1] * alpha) + (255 * (255 - alpha))) / 255;
            int blue = ((pixels[offset + 2] * alpha) + (255 * (255 - alpha))) / 255;
            int best = 0;
            int bestDistance = int.MaxValue;
            for (int paletteIndex = 0; paletteIndex < colors.Length; paletteIndex++) {
                int dr = red - colors[paletteIndex].R;
                int dg = green - colors[paletteIndex].G;
                int db = blue - colors[paletteIndex].B;
                int distance = (dr * dr * 3) + (dg * dg * 4) + (db * db * 2);
                if (distance >= bestDistance) continue;
                best = paletteIndex;
                bestDistance = distance;
                if (distance == 0) break;
            }
            result[index] = palette[best];
        }
        return result;
    }

    private static OfficeColor ParseArgb(string argb) => OfficeColor.FromRgb(
        System.Convert.ToByte(argb.Substring(2, 2), 16),
        System.Convert.ToByte(argb.Substring(4, 2), 16),
        System.Convert.ToByte(argb.Substring(6, 2), 16));

    private static string CreateFallbackSheetName(string source, HashSet<string> used) {
        char[] invalid = { ':', '\\', '/', '?', '*', '[', ']' };
        string sanitized = new string((source ?? string.Empty)
            .Select(character => invalid.Contains(character) ? '_' : character)
            .ToArray()).Trim('\'');
        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "Sheet";
        if (sanitized.Length > 31) sanitized = sanitized.Substring(0, 31);
        string candidate = sanitized;
        int suffix = 2;
        while (!used.Add(candidate)) {
            string ending = "_" + suffix++;
            candidate = sanitized.Substring(0, Math.Min(sanitized.Length, 31 - ending.Length)) + ending;
        }
        return candidate;
    }

    private static void AddExcelSourceCarrierDiagnostic(
        List<ExcelConversionDiagnostic> diagnostics,
        bool embedded,
        bool hasMacros,
        bool visualFallback) {
        diagnostics.Add(new ExcelConversionDiagnostic(
            embedded ? "Excel.SourceCarrier.Embedded" : "Excel.SourceCarrier.NotEmbedded",
            ExcelConversionDiagnosticCategory.DataLoss,
            ExcelConversionDiagnosticSeverity.Warning,
            embedded
                ? "The complete original source is retained in an inert, hash-verified OfficeIMO compatibility carrier. It is not executable or editable through the fallback workbook model."
                : "The original source carrier is not retained. Set EmbedSourceWhenLossy when deliberate byte-level recovery is required.",
            representsDataLoss: !embedded,
            embedded ? OfficeCompatibilityState.EmbeddedSource : OfficeCompatibilityState.Dropped,
            OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability
                | (visualFallback ? OfficeCompatibilityImpact.Semantic : OfficeCompatibilityImpact.None)
                | (hasMacros
                    ? OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Behavioral
                    : OfficeCompatibilityImpact.None),
            fallbackArtifact: embedded ? OfficeCompatibilitySourceCarrier.PayloadPath : null));
    }
}
