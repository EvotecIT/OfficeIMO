namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents the standard worksheet properties stored in BrtWsProp.</summary>
    internal sealed class XlsbWorksheetProperties {
        internal bool ShowAutomaticPageBreaks { get; set; }
        internal bool Published { get; set; }
        internal bool ApplyOutlineStyles { get; set; }
        internal bool SummaryRowsBelow { get; set; }
        internal bool SummaryColumnsRight { get; set; }
        internal bool FitToPage { get; set; }
        internal bool ShowOutlineSymbols { get; set; }
        internal bool SynchronizeHorizontal { get; set; }
        internal bool SynchronizeVertical { get; set; }
        internal bool TransitionEvaluation { get; set; }
        internal bool TransitionEntry { get; set; }
        internal bool FilterMode { get; set; }
        internal bool CalculateConditionalFormatting { get; set; }
        internal XlsbColor? TabColor { get; set; }
        internal uint SynchronizedRow { get; set; } = uint.MaxValue;
        internal uint SynchronizedColumn { get; set; } = uint.MaxValue;
        internal string CodeName { get; set; } = string.Empty;
    }
}
