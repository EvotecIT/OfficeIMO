namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets the separate authored worksheet print-area references in their stored order.
        /// References may include an absolute cell address and a quoted worksheet prefix.
        /// Returns an empty list when no print area is configured.
        /// </summary>
        public IReadOnlyList<string> GetPrintAreas() {
            string? text = GetPrintArea();
            return string.IsNullOrWhiteSpace(text)
                ? Array.Empty<string>()
                : SplitDefinedNameParts(text!).ToArray();
        }

        /// <summary>
        /// Parses a print-area reference belonging to this worksheet. External, other-sheet,
        /// malformed and out-of-bounds references are rejected before export.
        /// </summary>
        internal bool TryParsePrintAreaReference(string area, out ExcelReference? reference) {
            reference = null;
            if (!ExcelReference.TryParse(area, out ExcelReference? parsed) || parsed == null) return false;
            if (parsed.IsQualified &&
                (!SheetNameLookup.TryParseSheetQualifiedReference(area, out string sheetName, out _, allowExternalWorkbookReferences: false)
                 || !string.Equals(sheetName, Name, StringComparison.OrdinalIgnoreCase))) return false;
            reference = parsed;
            return true;
        }
    }
}
