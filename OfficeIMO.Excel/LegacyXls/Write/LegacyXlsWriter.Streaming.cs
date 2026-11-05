using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Stages BIFF records and the format-required SST, without retaining source rows.</summary>
    internal sealed class StreamingWorkbook : IDisposable {
        private readonly ExcelDocument _document;
        private readonly LegacyXlsFontTable _fonts;
        private readonly LegacyXlsStyleTable _styles;
        private readonly ushort[] _cellStyles;
        private readonly LegacyXlsDirectSharedStringBuilder _strings = new();
        private readonly List<StreamingSheet> _sheets = new();
        private readonly List<ExcelSheet> _metadata = new();

        internal StreamingWorkbook(ExcelDateSystem dateSystem, Stylesheet styles, OfficeTheme theme) {
            _document = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
            try {
                _document.DateSystem = dateSystem;
                // The generated font already names the resolved theme family. BIFF8 stores
                // that literal font, not the OOXML minor/major theme selection.
                foreach (var font in styles.Fonts!.Elements<Font>()) font.RemoveAllChildren<FontScheme>();
                _document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet = styles;
                _fonts = LegacyXlsFontTable.Create(_document);
                _styles = LegacyXlsStyleTable.Create(_document, Array.Empty<ExcelSheet>(), _fonts, new uint[] { 0, 1, 2, 3, 4 });
                _cellStyles = Enumerable.Range(0, 5).Select(index => _styles.GetBiffStyleIndex((uint)index)).ToArray();
            } catch { _document.Dispose(); throw; }
        }

        internal ExcelTabularRowSink AddSheet(string name, int columns) {
            // The facade has already resolved and registered this exact name.
            _metadata.Add(_document.AddSheet(name, ExcelSheetNameValidationMode.None));
            var sheet = new StreamingSheet(columns, _cellStyles, _strings);
            _sheets.Add(sheet);
            return sheet;
        }

        internal void Complete(Stream destination) {
            long capacity = 16_384L + _strings.EstimatedSerializedBytes + _sheets.Sum(sheet => sheet.Length + 16_384L);
            if (capacity > int.MaxValue) throw new NotSupportedException("The native XLS workbook exceeds 2 GiB.");
            using var workbook = new DirectTabularWorkbookStream((int)capacity);
            var positions = new List<long>(_sheets.Count);
            WriteDirectWorkbookGlobals(workbook, _document, _metadata[0], _fonts, _styles,
                LegacyXlsExternSheetTable.Create(_document, _metadata), _strings.Build(), _document.DateSystem,
                CancellationToken.None, _metadata, positions);
            for (int index = 0; index < _sheets.Count; index++) {
                long sheetPosition = workbook.Position;
                _sheets[index].WriteTo(workbook);
                long end = workbook.Position;
                workbook.Position = positions[index] + 4;
                WriteUInt32(workbook, checked((uint)sheetPosition));
                workbook.Position = end;
            }
            WriteDirectCompoundFile(_document, workbook, destination);
        }

        public void Dispose() {
            foreach (var sheet in _sheets) sheet.Dispose();
            _document.Dispose();
        }
    }
}
