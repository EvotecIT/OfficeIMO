using System.Globalization;
using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>Builds the stylesheet and assigns the XF indices used by the tabular writer.</summary>
    internal sealed class ExcelTabularStylesheet {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const int None = OfficeXmlNamespaces.None;
        private const int MaximumCellFormats = 64_000;

        private readonly OfficeTheme? _theme;
        private readonly OfficeXmlTree _tree;
        private readonly int _numberFormats;
        private readonly int _fonts;
        private readonly int _fills;
        private readonly int _styleFormats;
        private readonly int _cellFormats;
        private readonly int _styles;
        private readonly uint[] _temporalFormats;
        private readonly Dictionary<string, uint> _numberFormatIds = new(StringComparer.Ordinal);
        private readonly Dictionary<FontKey, uint> _fontIds = new();
        private readonly Dictionary<(string Attribute, string Value), uint> _fillIds = new();
        private readonly FontKey _normalFont;
        private uint _styleCount = 1;
        private uint _cellFormatCount = 1;

        internal ExcelTabularStylesheet(OfficeTheme? theme) {
            _theme = theme;
            OfficeTheme source = theme ?? OfficeTheme.Default;
            _normalFont = new FontKey(false, false, null, source.TypeScale.Body.ToPoints(), source.Fonts.Body.Family, null, null);
            _fontIds.Add(_normalFont, 0);
            byte[] seed = Encoding.UTF8.GetBytes(ExcelDocument.CreatedStylesXml(theme));
            _tree = OfficeXmlTree.Parse(seed, seed.Length, new OfficeXmlNamespaces.Table());
            _fonts = Section("fonts");
            _fills = Section("fills");
            _styleFormats = Section("cellStyleXfs");
            _cellFormats = Section("cellXfs");
            _styles = Section("cellStyles");
            _numberFormats = _tree.CreateElement(Spreadsheet, "numFmts");
            _tree.InsertBefore(_tree.Root, _numberFormats, _fonts);

            _temporalFormats = new[] { NumberFormat("yyyy-mm-dd hh:mm"), NumberFormat("[h]:mm:ss"), 14U, 46U };
            int normalFormat = _tree.FirstChildOf(_cellFormats);
            foreach (uint numberFormat in _temporalFormats) AddCellFormat(normalFormat, 0, numberFormat);
        }

        internal ExcelTabularStylePlan.DeclaredStyle AddStyle(string name, ExcelStyleDefinition definition) {
            uint styleId = _styleCount++;
            uint numberFormat = definition.NumberFormat == null ? 0U : NumberFormat(definition.NumberFormat);
            int format = Append(_styleFormats, "xf");
            Attribute(format, "numFmtId", numberFormat);
            Attribute(format, "fontId", Font(definition));
            Attribute(format, "fillId", definition.BackgroundColor is { } fill ? Fill(fill) : 0U);
            Attribute(format, "borderId", 0U);
            if (definition.NumberFormat != null) Attribute(format, "applyNumberFormat", "1");
            Attribute(format, "applyFont", "1");
            if (definition.BackgroundColor.HasValue) Attribute(format, "applyFill", "1");
            if (definition.WrapText || definition.HorizontalAlignment.HasValue || definition.VerticalAlignment.HasValue) {
                Attribute(format, "applyAlignment", "1");
                int alignment = Append(format, "alignment");
                if (definition.HorizontalAlignment is { } horizontal) Attribute(alignment, "horizontal", horizontal.ToOpenXmlToken());
                if (definition.VerticalAlignment is { } vertical) Attribute(alignment, "vertical", vertical.ToOpenXmlToken());
                Attribute(alignment, "wrapText", definition.WrapText ? "1" : "0");
            }

            int style = Append(_styles, "cellStyle");
            Attribute(style, "name", name);
            Attribute(style, "xfId", styleId);
            Attribute(style, "hidden", "0");

            string baseAttribute = StyleAttribute(AddCellFormat(format, styleId));
            var temporal = new string[_temporalFormats.Length];
            for (int index = 0; index < temporal.Length; index++) {
                temporal[index] = definition.NumberFormat != null
                    ? baseAttribute : StyleAttribute(AddCellFormat(format, styleId, _temporalFormats[index]));
            }

            if (_cellFormatCount > MaximumCellFormats) {
                throw new ArgumentException("Declared styles and their temporal variants exceed the 64,000 cell-format export limit.", "options");
            }
            return new ExcelTabularStylePlan.DeclaredStyle(baseAttribute, temporal);
        }

        private uint AddCellFormat(int source, uint styleId, uint? numberFormat = null) {
            // Keep the same attribute order in the style and cell XF entries.
            int cellFormat = Append(_cellFormats, "xf");
            Attribute(cellFormat, "numFmtId", numberFormat.HasValue ? XmlConvert.ToString(numberFormat.Value) : _tree.GetAttribute(source, None, "numFmtId")!);
            Attribute(cellFormat, "fontId", _tree.GetAttribute(source, None, "fontId")!);
            Attribute(cellFormat, "fillId", _tree.GetAttribute(source, None, "fillId")!);
            Attribute(cellFormat, "borderId", 0U);
            Attribute(cellFormat, "xfId", styleId);
            if (numberFormat.HasValue || _tree.HasAttribute(source, None, "applyNumberFormat")) Attribute(cellFormat, "applyNumberFormat", "1");
            if (_tree.HasAttribute(source, None, "applyFont")) Attribute(cellFormat, "applyFont", "1");
            if (_tree.HasAttribute(source, None, "applyFill")) Attribute(cellFormat, "applyFill", "1");
            if (_tree.HasAttribute(source, None, "applyAlignment")) Attribute(cellFormat, "applyAlignment", "1");
            int alignment = _tree.FindChild(source, Spreadsheet, "alignment");
            if (alignment != OfficeXmlTree.NoNode) _tree.Append(cellFormat, _tree.CloneElement(alignment));
            return _cellFormatCount++;
        }

        private uint NumberFormat(string code) {
            if (_numberFormatIds.TryGetValue(code, out uint id)) return id;
            id = 164U + (uint)_numberFormatIds.Count;
            _numberFormatIds.Add(code, id);
            int format = Append(_numberFormats, "numFmt");
            Attribute(format, "numFmtId", id);
            Attribute(format, "formatCode", code);
            return id;
        }

        private uint Font(ExcelStyleDefinition definition) {
            ExcelStyleColor? color = definition.Color is { } value ? ExcelStyleColor.From(value, _theme) : null;
            string? underline = definition.Underline is { } style && style != ExcelUnderlineStyle.None
                ? style.ToOpenXmlToken() : null;
            var font = new FontKey(definition.Bold, definition.Italic, underline,
                definition.FontSize?.ToPoints() ?? _normalFont.Size, definition.FontFamily ?? _normalFont.Family,
                color?.Attribute, color?.Value);
            if (_fontIds.TryGetValue(font, out uint id)) return id;
            id = (uint)_fontIds.Count;
            _fontIds.Add(font, id);
            int element = Append(_fonts, "font");
            if (font.Bold) Append(element, "b");
            if (font.Italic) Append(element, "i");
            if (font.Underline != null) Attribute(Append(element, "u"), "val", font.Underline);
            Attribute(Append(element, "sz"), "val", XmlConvert.ToString(font.Size));
            if (font.ColorAttribute != null) Attribute(Append(element, "color"), font.ColorAttribute, font.ColorValue!);
            Attribute(Append(element, "name"), "val", font.Family);
            return id;
        }

        private uint Fill(OfficeColor value) {
            ExcelStyleColor color = ExcelStyleColor.From(value, _theme);
            var key = (color.Attribute, color.Value);
            if (_fillIds.TryGetValue(key, out uint id)) return id;
            id = 2U + (uint)_fillIds.Count;
            _fillIds.Add(key, id);
            int pattern = Append(Append(_fills, "fill"), "patternFill");
            Attribute(pattern, "patternType", "solid");
            Attribute(Append(pattern, "fgColor"), color.Attribute, color.Value);
            Attribute(Append(pattern, "bgColor"), color.Attribute, color.Value);
            return id;
        }

        internal string WriteXml() {
            Attribute(_numberFormats, "count", (uint)_numberFormatIds.Count);
            Attribute(_fonts, "count", (uint)_fontIds.Count);
            Attribute(_fills, "count", 2U + (uint)_fillIds.Count);
            Attribute(_styles, "count", _styleCount);
            Attribute(_styleFormats, "count", _styleCount);
            Attribute(_cellFormats, "count", _cellFormatCount);

            var output = new StringBuilder();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { OmitXmlDeclaration = true })) {
                WriteElement(writer, _tree.Root, depth: 0);
            }
            return output.ToString();
        }

        // The generated style schema is at most five elements deep. Use XmlWriter for the existing
        // lexical form, including the space before an empty element's closing delimiter.
        private void WriteElement(XmlWriter writer, int element, int depth) {
            if (depth >= 5) throw new InvalidOperationException("Unexpected tabular stylesheet nesting.");
            writer.WriteStartElement("x", _tree.LocalNameOf(element), _tree.Namespaces.GetUri(Spreadsheet));
            int count = _tree.AttributeCountOf(element);
            for (int index = 0; index < count; index++) {
                string value = _tree.AttributeAt(element, index, out int ns, out string name);
                if (ns != OfficeXmlNamespaces.Xmlns) writer.WriteAttributeString(name, value);
            }
            for (int child = _tree.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = _tree.NextSiblingOf(child)) {
                WriteElement(writer, child, depth + 1);
            }
            writer.WriteEndElement();
        }

        private int Section(string name) => _tree.FindChild(_tree.Root, Spreadsheet, name);

        private int Append(int parent, string name) {
            int child = _tree.CreateElement(Spreadsheet, name);
            _tree.Append(parent, child);
            return child;
        }

        private void Attribute(int element, string name, string value) => _tree.SetAttribute(element, None, name, value);

        private void Attribute(int element, string name, uint value) => Attribute(element, name, XmlConvert.ToString(value));

        private static string StyleAttribute(uint index) => " s=\"" + index.ToString(CultureInfo.InvariantCulture) + "\"";

        private readonly record struct FontKey(bool Bold, bool Italic, string? Underline, double Size, string Family,
            string? ColorAttribute, string? ColorValue);
    }
}
