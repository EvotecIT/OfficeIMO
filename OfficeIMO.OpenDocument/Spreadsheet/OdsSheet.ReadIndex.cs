namespace OfficeIMO.OpenDocument;

public sealed partial class OdsSheet {
    private int _readContentVersion = -1;
    private int _readExternalVersion = -1;
    private ReadRun[]? _readRows;
    private readonly Dictionary<XElement, ReadRun[]> _readCells = new Dictionary<XElement, ReadRun[]>();

    private XElement? FindPrototypeRow(long rowIndex) {
        if (_readRows == null || _readContentVersion != _document.Package.ContentEditVersion ||
            _readExternalVersion != _document.Package.ExternalXmlEditVersion) {
            _readRows = IndexRuns(RowElements(), OdfNamespaces.Table + "number-rows-repeated");
            _readCells.Clear();
            _readContentVersion = _document.Package.ContentEditVersion;
            _readExternalVersion = _document.Package.ExternalXmlEditVersion;
        }
        return FindRun(_readRows, rowIndex);
    }

    private XElement? FindPrototypeCell(long rowIndex, long columnIndex) {
        XElement? row = FindPrototypeRow(rowIndex);
        if (row == null) return null;
        if (!_readCells.TryGetValue(row, out ReadRun[]? cells)) {
            cells = IndexRuns(CellElements(row), OdfNamespaces.Table + "number-columns-repeated");
            _readCells.Add(row, cells);
        }
        return FindRun(cells, columnIndex);
    }

    private static ReadRun[] IndexRuns(IEnumerable<XElement> elements, XName repeatAttribute) {
        var runs = new List<ReadRun>();
        long end = 0;
        foreach (XElement element in elements) {
            end = checked(end + OdsRepeatModel.Read(element, repeatAttribute));
            runs.Add(new ReadRun(end, element));
        }
        return runs.ToArray();
    }

    private static XElement? FindRun(ReadRun[] runs, long index) {
        int low = 0, high = runs.Length - 1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            if (index >= runs[middle].End) low = middle + 1;
            else high = middle - 1;
        }
        return low < runs.Length ? runs[low].Element : null;
    }

    private readonly struct ReadRun {
        internal ReadRun(long end, XElement element) { End = end; Element = element; }
        internal long End { get; }
        internal XElement Element { get; }
    }
}
