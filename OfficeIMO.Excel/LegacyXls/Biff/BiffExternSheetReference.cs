namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal readonly struct BiffExternSheetReference {
        internal BiffExternSheetReference(ushort supBookIndex, short firstSheetIndex, short lastSheetIndex) {
            SupBookIndex = supBookIndex;
            FirstSheetIndex = firstSheetIndex;
            LastSheetIndex = lastSheetIndex;
        }

        internal ushort SupBookIndex { get; }

        internal short FirstSheetIndex { get; }

        internal short LastSheetIndex { get; }
    }
}
