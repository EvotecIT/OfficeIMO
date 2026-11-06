using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private int _remainingGeneratedElements;
    private readonly Dictionary<OpenXmlPart, long> _storyGridMarkupBytes = new();

    private void InitializeReconstructionLimits(XElement root, OfficePackageSecurityOptions policy) {
        _remainingGeneratedElements = WordXmlInterchange.MaximumXmlElements;
        foreach (XElement unused in root.DescendantsAndSelf()) {
            _token.ThrowIfCancellationRequested();
            if (_remainingGeneratedElements == 0)
                throw new InvalidDataException("WordML 2003 exceeds the XML element reconstruction limit.");
            _remainingGeneratedElements--;
        }
        _remainingReconstructionBytes = Math.Min(WordXmlInterchange.MaximumXmlBytes, policy.MaxTotalUncompressedBytes);
    }

    private int AdmitTableGrid(XElement table, XElement? grid) {
        RequireSingle(table, "tblPr", "tblGrid");
        int explicitColumns = grid?.Elements(Word + "gridCol").Count() ?? 0;
        int columns = 1;
        foreach (XElement row in table.Elements(Word + "tr")) {
            _token.ThrowIfCancellationRequested();
            RequireSingle(row, "trPr", "tblPrEx");
            int occupied = 0;
            XElement? properties = row.Element(Word + "trPr");
            if (properties != null) RequireSingle(properties, "gridBefore", "gridAfter");
            Add(ReadGridUnits(properties?.Element(Word + "gridBefore"), span: false));
            foreach (XElement cell in row.Elements(Word + "tc")) {
                RequireSingle(cell, "tcPr");
                if (cell.Element(Word + "tcPr") is { } cellProperties) RequireSingle(cellProperties, "gridSpan");
                Add(ReadGridUnits(cell.Element(Word + "tcPr")?.Element(Word + "gridSpan"), span: true));
            }
            Add(ReadGridUnits(properties?.Element(Word + "gridAfter"), span: false));
            columns = Math.Max(columns, occupied);

            void Add(int units) {
                if (units > WordXmlInterchange.MaximumXmlElements - occupied) throw TableGridLimit();
                occupied += units;
            }
        }
        bool generated = explicitColumns == 0;
        int reserved = generated ? columns : explicitColumns;
        if (reserved > WordXmlInterchange.MaximumXmlElements || generated && reserved > _remainingGeneratedElements)
            throw TableGridLimit();

        // Even namespace-free <gridCol/> requires ten characters and UTF-8 bytes.
        // Reserve this lower bound before building DOM nodes; exact package admission still follows.
        long minimumMarkup = (long)reserved * 10;
        long partLimit = Math.Min(WordXmlInterchange.MaximumXmlBytes,
            Math.Min(_resourcePolicy.MaxPartUncompressedBytes, _resourcePolicy.MaxXmlCharactersInPart));
        _storyGridMarkupBytes.TryGetValue(_story, out long used);
        if (minimumMarkup > partLimit || used > partLimit - minimumMarkup || minimumMarkup > _remainingReconstructionBytes)
            throw TableGridLimit();
        if (generated) _remainingGeneratedElements -= reserved;
        _remainingReconstructionBytes -= minimumMarkup;
        _storyGridMarkupBytes[_story] = used + minimumMarkup;
        return columns;
    }

    private static int ReadGridUnits(XElement? element, bool span) {
        if (element == null) return span ? 1 : 0;
        if (!int.TryParse((string?)element.Attribute(Word + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture,
                out int value) || value < 0)
            throw TableGridLimit();
        return span ? Math.Max(1, value) : value;
    }

    private static void RequireSingle(XElement container, params string[] names) {
        foreach (string name in names) {
            if (container.Elements(Word + name).Skip(1).Any())
                throw new InvalidDataException("WordML 2003 table grid contains a duplicate " + name + " declaration.");
        }
    }

    private static InvalidDataException TableGridLimit() =>
        new("WordML 2003 table grid exceeds the supported reconstruction limits.");
}
