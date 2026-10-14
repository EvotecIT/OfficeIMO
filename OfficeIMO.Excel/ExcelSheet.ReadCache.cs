namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Header and search results belong to the worksheet part, including uncached wrappers.</summary>
        private sealed class WorksheetReadCache {
            internal Dictionary<string, int>? HeaderMap;
            internal string? HeaderMapSourceA1;
            internal int HeaderRowIndex;
            internal int HeaderCellCount;
            internal int FirstColumnIndex;
            internal int LastColumnIndex;
            internal bool NormalizeHeaders;
            internal bool TreatHeaderDates;
            internal bool HeaderOptionsReusable;
            internal CellFormatCacheRevision? HeaderStyleRevision;
            internal long HeaderStyleRevisionNumber;
            internal bool HeaderMapPopulated;
            internal readonly object HeaderMapLock = new object();
            internal string? FindFirstText;
            internal string? FindFirstAddress;
            internal bool FindFirstHasValue;
            internal readonly object FindFirstLock = new object();
        }
    }
}
