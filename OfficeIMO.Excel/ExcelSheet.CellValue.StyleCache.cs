using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Named styles can replace cell formats without changing their indexes or count.
        // The revision belongs to the shared format table, including uncached sheet wrappers.
        private static readonly ConditionalWeakTable<CellFormats, CellFormatCacheRevision> CellFormatCacheRevisions = new();
        private CellFormatCacheRevision? _cellValueStyleRevision;
        private long _cellValueStyleRevisionNumber;
        private WorkbookPart? _knownStylesOwner;
        private WorkbookStylesPart? _knownStylesPart;

        private sealed class CellFormatCacheRevision {
            internal long Number;
        }

        /// <summary>Reuses automatic number formats only while their shared base table is unchanged.</summary>
        private uint GetAutomaticNumberFormatStyleIndex(uint baseStyleIndex, bool duration) {
            RefreshAutomaticCellFormatCaches();
            if (baseStyleIndex == 0U) {
                uint? cached = duration ? _cellValueDefaultDurationStyleIndex : _cellValueDefaultDateStyleIndex;
                if (cached.HasValue) return cached.Value;
            } else {
                Dictionary<uint, uint>? indexes = duration ? _cellValueDurationStyleIndexes : _cellValueDateStyleIndexes;
                if (indexes != null && indexes.TryGetValue(baseStyleIndex, out uint cached)) return cached;
            }

            uint index = baseStyleIndex != 0U && HasMatchingTemporalNumberFormat(baseStyleIndex, duration)
                ? baseStyleIndex
                : GetOrCreateBuiltInNumberFormatStyleIndex(baseStyleIndex, duration ? 46U : 14U);
            // Creating the first automatic format can also create the shared format table.
            RefreshAutomaticCellFormatCaches();
            if (duration) {
                if (baseStyleIndex == 0U) _cellValueDefaultDurationStyleIndex = index;
                else (_cellValueDurationStyleIndexes ??= new Dictionary<uint, uint>())[baseStyleIndex] = index;
            } else {
                if (baseStyleIndex == 0U) _cellValueDefaultDateStyleIndex = index;
                else (_cellValueDateStyleIndexes ??= new Dictionary<uint, uint>())[baseStyleIndex] = index;
            }
            return index;
        }

        /// <summary>
        /// Excel keeps a cell's own date format when a date is entered, and its time format for a time, so a
        /// template's <c>yyyy-mm-dd</c> column is not replaced by the built-in short date.
        /// </summary>
        private bool HasMatchingTemporalNumberFormat(uint styleIndex, bool duration) {
            Stylesheet? stylesheet = GetKnownStylesPart()?.Stylesheet;
            if (stylesheet?.CellFormats == null || styleIndex >= (uint)stylesheet.CellFormats.ChildElements.Count) return false;
            if (stylesheet.CellFormats.ChildElements[(int)styleIndex] is not CellFormat format) return false;
            uint numberFormatId = format.NumberFormatId?.Value ?? 0U;
            if (numberFormatId == 0U) return false;
            ExcelSerialDateStyle style = ExcelNumberFormatClassifier.ClassifySerialDateStyle(numberFormatId, GetNumberFormatCode(stylesheet, numberFormatId));
            return style == (duration ? ExcelSerialDateStyle.Time : ExcelSerialDateStyle.Calendar);
        }

        private static void InvalidateAutomaticCellFormatCaches(CellFormats formats) =>
            CellFormatCacheRevisions.GetValue(formats, static _ => new CellFormatCacheRevision()).Number++;

        /// <summary>Discards automatic date/time formats after a shared format replacement or rollback.</summary>
        private void RefreshAutomaticCellFormatCaches() {
            CellFormatCacheRevision? revision = GetCurrentCellFormatCacheRevision();
            if (ReferenceEquals(_cellValueStyleRevision, revision)
                && _cellValueStyleRevisionNumber == (revision?.Number ?? 0)) return;

            _cellValueStyleRevision = revision;
            _cellValueStyleRevisionNumber = revision?.Number ?? 0;
            _cellValueDateStyleIndexes = null;
            _cellValueDurationStyleIndexes = null;
            _cellValueDefaultDateStyleIndex = null;
            _cellValueDefaultDurationStyleIndex = null;
        }

        /// <summary>Finding a part walks the relationships of the workbook, which is too slow to repeat per cell.</summary>
        private WorkbookStylesPart? GetKnownStylesPart() {
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot;
            if (_knownStylesPart == null || !ReferenceEquals(_knownStylesOwner, workbookPart)) {
                _knownStylesOwner = workbookPart;
                _knownStylesPart = workbookPart.WorkbookStylesPart;
            }
            return _knownStylesPart;
        }

        private CellFormatCacheRevision? GetCurrentCellFormatCacheRevision() {
            // The typed property walks the schema particle, which allocates, and this runs per date cell.
            for (OpenXmlElement? child = GetKnownStylesPart()?.Stylesheet?.FirstChild; child != null; child = child.NextSibling()) {
                if (child is CellFormats formats) return CellFormatCacheRevisions.GetValue(formats, static _ => new CellFormatCacheRevision());
            }
            return null;
        }
    }
}
