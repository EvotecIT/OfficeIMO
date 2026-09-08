using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reuses immutable schema names across XML readers. Unknown names belong to
    /// the individual reader, so arbitrary document names are never retained by
    /// a process-wide cache or shared between concurrent parsers.
    /// </summary>
    internal sealed class OpenXmlReadNameTable : XmlNameTable {
        private static readonly NameTable SchemaNames = CreateSchemaNames();
        private NameTable? _documentNames;

        public override string Add(string value) => SchemaNames.Get(value)
            ?? (_documentNames ??= new NameTable()).Add(value);

        public override string Add(char[] value, int start, int length) => SchemaNames.Get(value, start, length)
            ?? (_documentNames ??= new NameTable()).Add(value, start, length);

        public override string? Get(string value) => SchemaNames.Get(value) ?? _documentNames?.Get(value);

        public override string? Get(char[] value, int start, int length) => SchemaNames.Get(value, start, length)
            ?? _documentNames?.Get(value, start, length);

        internal static XmlReaderSettings WithSchemaNames(XmlReaderSettings template) {
            XmlReaderSettings settings = template.Clone();
            settings.NameTable = new OpenXmlReadNameTable();
            return settings;
        }

        private static NameTable CreateSchemaNames() {
            var names = new NameTable();
            foreach (string name in new[] {
                "", "xml", "xmlns", "version", "encoding", "standalone", "r", "mc", "x14ac", "x15", "xr", "xr2", "xr3",
                "http://www.w3.org/XML/1998/namespace", "http://www.w3.org/2000/xmlns/",
                "http://schemas.openxmlformats.org/package/2006/relationships",
                "http://schemas.openxmlformats.org/package/2006/content-types",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "http://purl.oclc.org/ooxml/spreadsheetml/main",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
                "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
                "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
                "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
                "Types", "Default", "Override", "Extension", "ContentType", "PartName",
                "Relationships", "Relationship", "Id", "Type", "Target", "TargetMode",
                "workbook", "workbookPr", "date1904", "sheets", "sheet", "sheetId", "name", "state", "id",
                "fileVersion", "appName", "lastEdited", "lowestEdited", "rupBuild", "codeName", "defaultThemeVersion",
                "bookViews", "workbookView", "xWindow", "yWindow", "windowWidth", "windowHeight", "activeTab",
                "definedNames", "definedName", "localSheetId", "hidden", "calcPr", "calcId", "fullCalcOnLoad",
                "workbookProtection", "fileSharing", "extLst", "ext", "uri", "Ignorable", "AlternateContent",
                "Choice", "Fallback", "Requires", "uid", "revisionPtr", "lastSaveId", "documentId",
                "styleSheet", "numFmts", "numFmt", "numFmtId", "formatCode", "count",
                "fonts", "font", "sz", "val", "color", "rgb", "indexed", "theme", "tint", "family", "scheme",
                "b", "i", "u", "strike", "outline", "shadow", "condense", "extend", "vertAlign", "charset",
                "fills", "fill", "patternFill", "patternType", "fgColor", "bgColor", "gradientFill", "stop", "position",
                "borders", "border", "left", "right", "top", "bottom", "diagonal", "style", "diagonalUp", "diagonalDown",
                "cellStyleXfs", "cellXfs", "xf", "fontId", "fillId", "borderId", "xfId", "applyNumberFormat",
                "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection", "quotePrefix", "pivotButton",
                "alignment", "horizontal", "vertical", "wrapText", "textRotation", "indent", "shrinkToFit", "readingOrder",
                "protection", "locked", "cellStyles", "cellStyle", "builtinId", "customBuiltin", "dxfs",
                "tableStyles", "defaultTableStyle", "defaultPivotStyle", "colors", "indexedColors", "rgbColor",
                "sst", "si", "t", "uniqueCount", "space", "rPr", "rFont", "rPh", "phoneticPr", "sb", "eb", "type"
            }) names.Add(name);
            return names;
        }
    }
}
