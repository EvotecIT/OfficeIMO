using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static readonly Lazy<byte[]> ChartStyle251Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-style-251.xml"));
        private static readonly Lazy<byte[]> ChartColorStyle10Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-colors-10.xml"));

        private static StringReference CreateStringReference(string formula, IReadOnlyList<string>? values) {
            var reference = new StringReference(new Formula { Text = formula });
            if (values == null) return reference;

            StringCache cache = new();
            cache.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new StringPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i] ?? string.Empty }
                });
            }
            reference.Append(cache);
            return reference;
        }

        private static StringLiteral CreateStringLiteral(IReadOnlyList<string> values) {
            StringLiteral literal = new();
            literal.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                literal.Append(new StringPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i] ?? string.Empty }
                });
            }
            return literal;
        }

        private static NumberReference CreateNumberReference(string formula, IReadOnlyList<double>? values) {
            var reference = new NumberReference(new Formula { Text = formula });
            if (values == null) return reference;

            NumberingCache cache = new();
            cache.Append(new FormatCode { Text = "General" });
            cache.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new NumericPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i].ToString(CultureInfo.InvariantCulture) }
                });
            }
            reference.Append(cache);
            return reference;
        }

        private static Title CreateChartTitle(string text) {
            return new Title(
                new ChartText(CreateChartText(text)),
                new Overlay { Val = false }
            );
        }

        private static RichText CreateChartText(string text) {
            return new RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties(new A.DefaultRunProperties()),
                    new A.Run(new A.Text { Text = text })
                ));
        }

        internal static void ApplyChartStyle(ChartPart chartPart, int styleId, int colorStyleId) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            ApplyChartStyle(chartPart, new ExcelChartStylePreset(styleId, colorStyleId));
        }

        internal static void ApplyChartStyle(ChartPart chartPart, ExcelChartStylePreset preset) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (preset == null) {
                throw new ArgumentNullException(nameof(preset));
            }

            byte[] styleBytes = preset.StyleXmlBytes ?? GetChartStyleBytes(preset.StyleId);
            byte[] colorBytes = preset.ColorXmlBytes ?? GetChartColorStyleBytes(preset.ColorStyleId);

            ChartStylePart stylePart = chartPart.GetPartsOfType<ChartStylePart>().FirstOrDefault()
                ?? chartPart.AddNewPart<ChartStylePart>();
            PopulateChartStyle(stylePart, styleBytes);
            ChartColorStylePart colorStylePart = chartPart.GetPartsOfType<ChartColorStylePart>().FirstOrDefault()
                ?? chartPart.AddNewPart<ChartColorStylePart>();
            PopulateChartColorStyle(colorStylePart, colorBytes);
        }

        internal static void PopulateChartStyle(ChartStylePart stylePart, byte[]? xmlBytes = null) {
            if (stylePart == null) {
                throw new ArgumentNullException(nameof(stylePart));
            }

            xmlBytes ??= ChartStyle251Bytes.Value;
            using var stream = new MemoryStream(xmlBytes);
            stylePart.FeedData(stream);
        }

        internal static void PopulateChartColorStyle(ChartColorStylePart colorStylePart, byte[]? xmlBytes = null) {
            if (colorStylePart == null) {
                throw new ArgumentNullException(nameof(colorStylePart));
            }

            xmlBytes ??= ChartColorStyle10Bytes.Value;
            using var stream = new MemoryStream(xmlBytes);
            colorStylePart.FeedData(stream);
        }

        private static byte[] GetChartStyleBytes(int styleId) {
            if (styleId == 251) return ChartStyle251Bytes.Value;
            return ChartStyle251Bytes.Value;
        }

        private static byte[] GetChartColorStyleBytes(int colorStyleId) {
            if (colorStyleId == 10) return ChartColorStyle10Bytes.Value;
            return ChartColorStyle10Bytes.Value;
        }

        private static byte[] LoadEmbeddedResource(string resourceName) {
            var assembly = typeof(ExcelChartUtils).Assembly;
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new InvalidOperationException($"Missing embedded resource '{resourceName}'.");
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
