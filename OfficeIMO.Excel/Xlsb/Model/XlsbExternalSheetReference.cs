namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Maps one BIFF12 external-sheet index to a supporting link and sheet span.</summary>
    internal sealed class XlsbExternalSheetReference {
        internal XlsbExternalSheetReference(uint supportingLinkIndex, int firstSheetIndex, int lastSheetIndex) {
            SupportingLinkIndex = supportingLinkIndex;
            FirstSheetIndex = firstSheetIndex;
            LastSheetIndex = lastSheetIndex;
        }

        internal uint SupportingLinkIndex { get; }

        internal int FirstSheetIndex { get; }

        internal int LastSheetIndex { get; }
    }
}
