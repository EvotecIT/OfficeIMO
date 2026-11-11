using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>The BIFF8 font, format, XF and style records derived once from generated workbook styles.</summary>
    internal sealed class StreamingStyles {
        private StreamingStyles(byte[][] fonts, byte[][] formats, byte[][] cellFormats, byte[][] styles, ushort[] cellStyles) {
            Fonts = fonts; Formats = formats; CellFormats = cellFormats; Styles = styles; CellStyles = cellStyles;
        }

        internal byte[][] Fonts { get; }
        internal byte[][] Formats { get; }
        internal byte[][] CellFormats { get; }
        internal byte[][] Styles { get; }
        /// <summary>BIFF XF indexes for the five generated tabular cell styles.</summary>
        internal ushort[] CellStyles { get; }

        /// <summary>Converts the generated stylesheet through the workbook style tables used by every XLS save.</summary>
        internal static StreamingStyles Create(string stylesXml, OfficeTheme theme) {
            using var document = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
            var styles = new Stylesheet(stylesXml);
            // The generated font already names the resolved theme family. BIFF8 stores
            // that literal font, not the OOXML minor/major theme selection.
            foreach (var font in styles.Fonts!.Elements<Font>()) font.RemoveAllChildren<FontScheme>();
            document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet = styles;
            var fonts = LegacyXlsFontTable.Create(document);
            var table = LegacyXlsStyleTable.Create(document, Array.Empty<ExcelSheet>(), fonts, new uint[] { 0, 1, 2, 3, 4 });
            return new StreamingStyles(fonts.FontRecords.ToArray(), table.FormatRecords.ToArray(), table.CellFormatRecords.ToArray(),
                table.StyleRecords.ToArray(), Enumerable.Range(0, 5).Select(index => table.GetBiffStyleIndex((uint)index)).ToArray());
        }
    }

    /// <summary>Stages BIFF records without retaining source rows. Shared strings are used when requested or for text over 255 characters.</summary>
    internal sealed class StreamingWorkbook : IDisposable {
        private readonly ExcelDateSystem _dateSystem;
        private readonly StreamingStyles _styles;
        private readonly bool _inlineText;
        private static LegacyXlsDirectSharedStringBuilder? _cachedStrings;
        private LegacyXlsDirectSharedStringBuilder? _leasedStrings;
        private readonly LegacyXlsDirectSharedStringBuilder _strings;
        private readonly List<StreamingSheet> _sheets = new();
        private readonly List<string> _names = new();

        /// <summary>Exclusively leases a cleared shared-string workspace.</summary>
        internal StreamingWorkbook(ExcelDateSystem dateSystem, StreamingStyles styles, bool useSharedStrings) {
            _dateSystem = dateSystem; _styles = styles; _inlineText = !useSharedStrings;
            _leasedStrings = Interlocked.Exchange(ref _cachedStrings, null) ?? new LegacyXlsDirectSharedStringBuilder();
            _strings = _leasedStrings;
        }

        internal ExcelTabularRowSink AddSheet(string name, int columns) {
            // The facade has already resolved and registered this exact name.
            if (name.Length > 31) throw new NotSupportedException("Native XLS saving supports worksheet names up to 31 characters.");
            _names.Add(name);
            var sheet = new StreamingSheet(columns, _styles.CellStyles, _strings, _inlineText);
            _sheets.Add(sheet);
            return sheet;
        }

        /// <summary>Writes the globals, then publishes them and each sheet's owned buffers straight into the compound file.</summary>
        internal void Complete(Stream destination) {
            // The SST estimate counts string bytes only. CONTINUE headers and split-string option bytes add
            // 5 bytes per 8,224-byte record, and EXTSST up to about 1 KiB; without that room the buffer would
            // double, which on .NET Framework allocates a second full-size array on every save.
            long estimate = _strings.EstimatedSerializedBytes;
            long capacity = 16_384L + estimate + estimate / 1_024L + 8_192L;
            if (capacity > int.MaxValue) throw new NotSupportedException("The native XLS workbook exceeds 2 GiB.");
            using var globals = new DirectTabularWorkbookStream((int)capacity);
            var boundSheets = new List<long>(_sheets.Count);
            byte[] tabIds = new byte[_names.Count * 2];
            for (int index = 0; index < _names.Count; index++) WriteUInt16(tabIds, index * 2, checked((ushort)(index + 1)));
            WriteWorkbookGlobals(globals, tabIds, _names.Select(name => BuildBoundSheetPayload(0, name, 0)).ToList(),
                _styles.Fonts, _styles.Formats, _styles.CellFormats, _styles.Styles,
                LegacyXlsExternSheetTable.CreateStreaming(_names), _strings.Build(reuseEntries: true), _dateSystem,
                CancellationToken.None, boundSheets);
            byte[] globalsBuffer = globals.Buffer;
            var segments = new List<ArraySegment<byte>> { new(globalsBuffer, 0, checked((int)globals.Length)) };
            long position = globals.Length;
            for (int index = 0; index < _sheets.Count; index++) {
                if (position > uint.MaxValue) throw new NotSupportedException("The native XLS workbook exceeds 4 GiB.");
                WriteUInt32(globalsBuffer, checked((int)boundSheets[index] + 4), checked((uint)position));
                position = _sheets[index].AddSegments(segments, position);
            }
            var streams = new List<OfficeCompoundStream>(1) { new(GetWorkbookStreamName(null), segments) };
            OfficeCompoundWriterLayout layout = OfficeCompoundWriterLayout.Create(streams);
            long length = OfficeCompoundFileWriter.GetSerializedLength(layout);
            ExcelDocument.PrepareDestinationStreamForWrite(destination);
            if (destination.GetType() == typeof(MemoryStream) && ((MemoryStream)destination).Capacity < length) {
                if (length > int.MaxValue) throw new NotSupportedException("A compound file returned as a byte array cannot exceed 2 GiB.");
                ((MemoryStream)destination).Capacity = checked((int)length);
            }
            OfficeCompoundFileWriter.Write(destination, layout, rootClassId: null);
        }

        /// <summary>Disposes sheet buffers and releases the shared-string lease once.</summary>
        public void Dispose() {
            try {
                foreach (var sheet in _sheets) sheet.Dispose();
            } finally {
                ReleaseSharedStrings();
            }
        }

        /// <summary>Clears this workbook's text before returning its exclusive shared-string workspace.</summary>
        private void ReleaseSharedStrings() {
            LegacyXlsDirectSharedStringBuilder? strings = Interlocked.Exchange(ref _leasedStrings, null);
            if (strings != null && strings.ClearForReuse()) {
                Interlocked.CompareExchange(ref _cachedStrings, strings, null);
            }
        }
    }
}
