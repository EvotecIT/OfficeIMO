using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Resolves the existing insertion contract's date format without materializing tabular cells.</summary>
        internal uint GetOrCreateDirectTabularDateStyle(bool useCellValueNumberFormats) {
            if (useCellValueNumberFormats) {
                return GetOrCreateBuiltInNumberFormatStyleIndex(0U, 14U);
            }

            WorkbookStylesPart stylesPart = _excelDocument.WorkbookPartRoot.WorkbookStylesPart
                ?? _excelDocument.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            CellFormat format = GetBaseCellFormat(stylesheet, 0U);
            format.NumberFormatId = GetOrCreateNumberFormatId(stylesheet, DataTableDateTimeNumberFormat);
            format.ApplyNumberFormat = true;
            uint index = AppendOrReuseCellFormat(stylesheet, format);
            SaveStylesheet(stylesPart);
            return index;
        }
    }
}
