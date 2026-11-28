using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static readonly Lazy<byte[]> ChartStyle251Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-style-251.xml"));
        private static readonly Lazy<byte[]> ChartColorStyle10Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-colors-10.xml"));

        // The builders below make detached elements with the element they will go under as the creation context, so
        // each one gets the class that parent expects.

        private static OfficeXmlNode CreateStringReference(OfficeXmlNode parent, string formula, IReadOnlyList<string>? values) {
            OfficeXmlNode reference = parent.New(C, "strRef");
            reference.Append(reference.NewText(C, "f", formula));
            if (values == null) return reference;
            reference.Append(CreateStringCache(reference, values));
            return reference;
        }

        private static OfficeXmlNode CreateStringCache(OfficeXmlNode parent, IReadOnlyList<string> values) {
            OfficeXmlNode cache = parent.New(C, "strCache");
            cache.Append(cache.NewVal(C, "ptCount", (uint)values.Count));
            for (int i = 0; i < values.Count; i++) {
                cache.Append(CreateStringPoint(cache, (uint)i, values[i] ?? string.Empty));
            }
            return cache;
        }

        private static OfficeXmlNode CreateStringPoint(OfficeXmlNode parent, uint index, string text) {
            OfficeXmlNode point = parent.New(C, "pt").Attr("idx", index);
            point.Append(point.NewText(C, "v", text));
            return point;
        }

        private static OfficeXmlNode CreateSingleStringReference(OfficeXmlNode parent, string formula, string? value) {
            OfficeXmlNode reference = parent.New(C, "strRef");
            reference.Append(reference.NewText(C, "f", formula));
            reference.Append(CreateSingleStringCache(reference, value));
            return reference;
        }

        private static OfficeXmlNode CreateStringLiteral(OfficeXmlNode parent, IReadOnlyList<string> values) {
            OfficeXmlNode literal = parent.New(C, "strLit");
            literal.Append(literal.NewVal(C, "ptCount", (uint)values.Count));
            for (int i = 0; i < values.Count; i++) {
                literal.Append(CreateStringPoint(literal, (uint)i, values[i] ?? string.Empty));
            }
            return literal;
        }

        private static OfficeXmlNode CreateSingleStringCache(OfficeXmlNode parent, string? value) {
            OfficeXmlNode cache = parent.New(C, "strCache");
            cache.Append(cache.NewVal(C, "ptCount", 1U));
            cache.Append(CreateStringPoint(cache, 0U, value ?? string.Empty));
            return cache;
        }

        private static OfficeXmlNode CreateNumberReference(OfficeXmlNode parent, string formula, IReadOnlyList<double>? values) {
            OfficeXmlNode reference = parent.New(C, "numRef");
            reference.Append(reference.NewText(C, "f", formula));
            if (values == null) return reference;

            OfficeXmlNode cache = reference.New(C, "numCache");
            AppendNumberPoints(cache, values);
            reference.Append(cache);
            return reference;
        }

        private static OfficeXmlNode CreateNumberLiteral(OfficeXmlNode parent, IReadOnlyList<double> values) {
            OfficeXmlNode literal = parent.New(C, "numLit");
            AppendNumberPoints(literal, values);
            return literal;
        }

        private static void AppendNumberPoints(OfficeXmlNode container, IReadOnlyList<double> values) {
            container.Append(container.NewText(C, "formatCode", "General"));
            container.Append(container.NewVal(C, "ptCount", (uint)values.Count));
            for (int i = 0; i < values.Count; i++) {
                OfficeXmlNode point = container.New(C, "pt").Attr("idx", (uint)i);
                point.Append(point.NewText(C, "v", InvariantNumberText.Get(values[i])));
                container.Append(point);
            }
        }

        private static OfficeXmlNode CreateChartTitle(OfficeXmlNode parent, string text) {
            OfficeXmlNode title = parent.New(C, "title");
            OfficeXmlNode chartText = title.New(C, "tx");
            chartText.Append(CreateChartText(chartText, text));
            title.Append(chartText);
            title.Append(title.NewVal(C, "overlay", false));
            return title;
        }

        private static OfficeXmlNode CreateChartText(OfficeXmlNode parent, string text) {
            OfficeXmlNode rich = parent.New(C, "rich");
            rich.Append(rich.New(A, "bodyPr"));
            rich.Append(rich.New(A, "lstStyle"));
            OfficeXmlNode paragraph = rich.New(A, "p");
            OfficeXmlNode paragraphProperties = paragraph.New(A, "pPr");
            paragraphProperties.Append(paragraphProperties.New(A, "defRPr"));
            paragraph.Append(paragraphProperties);
            OfficeXmlNode run = paragraph.New(A, "r");
            run.Append(run.NewText(A, "t", text));
            paragraph.Append(run);
            rich.Append(paragraph);
            return rich;
        }

        internal static void ApplyChartStyle(ExcelPartRef chartPart, int styleId, int colorStyleId) {
            if (chartPart.IsNone) throw new ArgumentNullException(nameof(chartPart));
            ApplyChartStyle(chartPart, new ExcelChartStylePreset(styleId, colorStyleId));
        }

        internal static void ApplyChartStyle(ExcelPartRef chartPart, ExcelChartStylePreset preset) {
            if (chartPart.IsNone) throw new ArgumentNullException(nameof(chartPart));
            ArgumentNullException.ThrowIfNull(preset);

            byte[] styleBytes = preset.StyleXmlBytes ?? GetChartStyleBytes(preset.StyleId);
            byte[] colorBytes = preset.ColorXmlBytes ?? GetChartColorStyleBytes(preset.ColorStyleId);

            ExcelPartRef stylePart = chartPart.FirstPartOfKind(ExcelPartKind.ChartStyle);
            if (stylePart.IsNone) stylePart = chartPart.AddNewPart(ExcelPartKind.ChartStyle);
            PopulateChartStyle(stylePart, styleBytes);
            ExcelPartRef colorStylePart = chartPart.FirstPartOfKind(ExcelPartKind.ChartColorStyle);
            if (colorStylePart.IsNone) colorStylePart = chartPart.AddNewPart(ExcelPartKind.ChartColorStyle);
            PopulateChartColorStyle(colorStylePart, colorBytes);
        }

        internal static void PopulateChartStyle(ExcelPartRef stylePart, byte[]? xmlBytes = null) {
            if (stylePart.IsNone) throw new ArgumentNullException(nameof(stylePart));

            xmlBytes ??= ChartStyle251Bytes.Value;
            stylePart.Feed(xmlBytes);
        }

        internal static void PopulateChartColorStyle(ExcelPartRef colorStylePart, byte[]? xmlBytes = null) {
            if (colorStylePart.IsNone) throw new ArgumentNullException(nameof(colorStylePart));

            xmlBytes ??= ChartColorStyle10Bytes.Value;
            colorStylePart.Feed(xmlBytes);
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
