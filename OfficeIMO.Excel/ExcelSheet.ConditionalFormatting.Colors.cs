using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Sets supplied text and solid-fill colors on an attached conditional rule, preserving its other native styling.
        /// Returns a fresh editable snapshot. If both colors are omitted, applies the creation theme's Danger text and light Danger fill.
        /// When one color is supplied, the other remains unchanged. Native theme roles stay linked to the workbook.
        /// </summary>
        public ExcelConditionalFormattingInfo SetConditionalFormattingColors(ExcelConditionalFormattingInfo rule,
            OfficeColor? color = null, OfficeColor? fillColor = null) {
            ArgumentNullException.ThrowIfNull(rule);
            if (!color.HasValue && !fillColor.HasValue) {
                color = OfficeColor.Danger;
                fillColor = _excelDocument.EffectiveCreationTheme.Surfaces.DangerFill;
            }
            ExcelConditionalFormattingInfo? updated = null;
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLock(() => {
                OpenXmlElement backing = RequireAttachedConditionalFormattingRule(rule);
                WorkbookPart workbook = _excelDocument.WorkbookPartRoot;
                Color? font = color.HasValue ? ExcelColor.ToSpreadsheetColor(color.Value, _excelDocument.EffectiveCreationTheme) : null;
                Color? fill = fillColor.HasValue ? ExcelColor.ToSpreadsheetColor(fillColor.Value, _excelDocument.EffectiveCreationTheme) : null;
                if (backing is ConditionalFormattingRule standard) {
                    WorkbookStylesPart stylesPart = workbook.WorkbookStylesPart ?? workbook.AddNewPart<WorkbookStylesPart>();
                    Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                    EnsureDefaultStylePrimitives(stylesheet);
                    var candidate = GetDifferentialFormat(stylesheet, standard.FormatId?.Value)?.CloneNode(true) as DifferentialFormat
                        ?? new DifferentialFormat();
                    ApplyConditionalColors(candidate, font, fill);
                    standard.FormatId = GetOrAddConditionalDifferentialFormat(stylesPart, stylesheet, candidate);
                    string range = (standard.Parent as ConditionalFormatting)?.SequenceOfReferences?.InnerText ?? rule.Range;
                    updated = ReadConditionalFormattingInfo(standard, range, stylesheet, workbook);
                } else if (backing is X14.ConditionalFormattingRule extension) {
                    X14.DifferentialType differential = extension.GetFirstChild<X14.DifferentialType>() ?? new X14.DifferentialType();
                    ApplyConditionalColors(differential, font, fill);
                    if (differential.Parent == null) InsertBeforeRuleExtension(extension, differential);
                    string range = extension.Parent?.GetFirstChild<Xm.ReferenceSequence>()?.Text ?? rule.Range;
                    updated = ReadOffice2010ConditionalFormattingInfo(extension, range, workbook);
                } else {
                    throw new InvalidOperationException("The conditional-formatting rule type is unsupported.");
                }
            });
            return updated!;
        }

        private static void ApplyConditionalColors(OpenXmlCompositeElement differential, Color? fontColor, Color? fillColor) {
            if (fontColor != null) {
                Font font = differential.GetFirstChild<Font>() ?? new Font();
                SetFontColorElement(font, fontColor);
                if (font.Parent == null) InsertDifferentialStyleChild(differential, font);
            }
            if (fillColor != null) {
                differential.GetFirstChild<Fill>()?.Remove();
                InsertDifferentialStyleChild(differential, new Fill(new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = fillColor.Rgb, Theme = fillColor.Theme },
                    BackgroundColor = new BackgroundColor { Rgb = fillColor.Rgb, Theme = fillColor.Theme }
                }));
            }
        }
    }
}
