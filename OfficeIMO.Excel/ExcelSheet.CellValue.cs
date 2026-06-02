using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal readonly struct DirectFormulaCellValue {
        internal DirectFormulaCellValue(string formula, string? formulaXml = null, string? cachedValue = null) {
            Formula = formula;
            FormulaXml = formulaXml;
            CachedValue = cachedValue;
        }

        internal string Formula { get; }

        internal string? FormulaXml { get; }

        internal string? CachedValue { get; }

        public override string ToString() => Formula;
    }

    internal readonly struct DirectTypedCellValue {
        internal DirectTypedCellValue(string dataType, string? value, string? inlineStringXml = null) {
            DataType = dataType;
            Value = value;
            InlineStringXml = inlineStringXml;
        }

        internal string DataType { get; }

        internal string? Value { get; }

        internal string? InlineStringXml { get; }
    }

    public partial class ExcelSheet {
        internal const int CellValuePlainStringPromotionSharedStringCount = 4096;
        private const int CellValueSharedStringIndexCacheLimit = 256;
        private const int PendingDirectCellValueMinimumCellCount = 128;
        private static readonly bool EnablePendingDirectCellValueBuffer = true;
        private static readonly bool MirrorPendingDirectCellValueBufferToWorksheet = false;
        private static readonly DateTime CellValueExcelMinimumSupportedDate = DateTime.FromOADate(2d);
        private Dictionary<uint, uint>? _cellValueDateStyleIndexes;
        private Dictionary<uint, uint>? _cellValueDurationStyleIndexes;
        private Dictionary<string, CellValueSharedStringIndexCacheEntry>? _cellValueSharedStringIndexCache;
        private uint? _cellValueDefaultDateStyleIndex;
        private uint? _cellValueDefaultDurationStyleIndex;
        private CellValueDirectSaveBuffer? _pendingCellValueDirectSaveBuffer;
        private int _pendingCellValueDirectSaveThreadId;
        private bool _disablePendingCellValueDirectSaveBuffer;
        private bool _materializingPendingCellValueDirectSaveBuffer;
        private bool _hasCellValueDomWrites;
    }
}
