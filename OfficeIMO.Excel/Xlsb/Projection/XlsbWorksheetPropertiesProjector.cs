using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares worksheet-level properties stored in BrtWsProp.</summary>
    internal static class XlsbWorksheetPropertiesProjector {
        internal const string ElementName = "sheetPr";

        internal static void Apply(ExcelSheet sheet, XlsbWorksheetProperties? source) {
            if (source != null) Apply(XlsbWorksheetMarkup.Root(sheet), source);
        }

        internal static void Apply(OfficeXmlNode worksheet, XlsbWorksheetProperties source) =>
            worksheet.InsertBefore(Create(worksheet, source), worksheet.FirstChild);

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheetProperties? source) {
            OfficeXmlNode worksheet = XlsbWorksheetMarkup.Root(sheet);
            OfficeXmlNode expected = source == null ? default : Create(worksheet, source);
            if (!XlsbWorksheetMarkup.MatchesChild(worksheet, ElementName, expected)) {
                throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify worksheet properties on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
            }
        }

        private static OfficeXmlNode Create(OfficeXmlNode worksheet, XlsbWorksheetProperties source) {
            OfficeXmlNode properties = XlsbWorksheetMarkup.Create(worksheet, ElementName);
            properties.SetBooleanAttribute("published", source.Published);
            properties.SetBooleanAttribute("syncHorizontal", source.SynchronizeHorizontal);
            properties.SetBooleanAttribute("syncVertical", source.SynchronizeVertical);
            properties.SetBooleanAttribute("transitionEvaluation", source.TransitionEvaluation);
            properties.SetBooleanAttribute("transitionEntry", source.TransitionEntry);
            properties.SetBooleanAttribute("filterMode", source.FilterMode);
            properties.SetBooleanAttribute("enableFormatConditionsCalculation", source.CalculateConditionalFormatting);
            if (!string.IsNullOrEmpty(source.CodeName)) properties.SetAttribute("codeName", source.CodeName);
            if ((source.SynchronizeHorizontal || source.SynchronizeVertical)
                && source.SynchronizedRow != uint.MaxValue
                && source.SynchronizedColumn != uint.MaxValue) {
                properties.SetAttribute("syncRef", ExcelA1.CellReference(
                    checked((int)source.SynchronizedRow + 1),
                    checked((int)source.SynchronizedColumn + 1)));
            }
            if (source.TabColor?.Type is > 0 and <= 3) {
                OfficeXmlNode color = XlsbWorksheetMarkup.Append(properties, "tabColor");
                ApplyColor(color, source.TabColor);
            }
            OfficeXmlNode outline = XlsbWorksheetMarkup.Append(properties, "outlinePr");
            outline.SetBooleanAttribute("applyStyles", source.ApplyOutlineStyles);
            outline.SetBooleanAttribute("summaryBelow", source.SummaryRowsBelow);
            outline.SetBooleanAttribute("summaryRight", source.SummaryColumnsRight);
            outline.SetBooleanAttribute("showOutlineSymbols", source.ShowOutlineSymbols);
            OfficeXmlNode pageSetup = XlsbWorksheetMarkup.Append(properties, "pageSetUpPr");
            pageSetup.SetBooleanAttribute("autoPageBreaks", source.ShowAutomaticPageBreaks);
            pageSetup.SetBooleanAttribute("fitToPage", source.FitToPage);
            return properties;
        }

        private static void ApplyColor(OfficeXmlNode target, XlsbColor source) {
            if (source.Type == 1) target.SetUInt32Attribute("indexed", source.Index);
            else if (source.Type == 2) target.SetAttribute("rgb", $"{source.Alpha:X2}{source.Red:X2}{source.Green:X2}{source.Blue:X2}");
            else if (source.Type == 3) target.SetUInt32Attribute("theme", source.Index);
            if (source.Tint != 0) {
                target.SetDoubleAttribute("tint", source.Tint / (source.Tint < 0 ? 32768D : 32767D));
            }
        }
    }
}
