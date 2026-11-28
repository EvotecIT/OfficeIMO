using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void SetOrientationCore(OfficePageOrientation orientation) {
            Persistence.OwnsTablePrintDefaults = false;
            SetPageSetupAttribute("orientation", orientation == OfficePageOrientation.Landscape ? "landscape" : "portrait");
        }

        internal void SetPaperSizeCore(ExcelPaperSize paperSize) {
            ValidatePaperSize(paperSize);
            SetPageSetupAttribute("paperSize", ((uint)paperSize).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private void SetPageSetupAttribute(string name, string value) =>
            EditWorksheetObject(worksheet => {
                GetOrCreatePageSetup(worksheet).SetAttribute(name, value);
                return true;
            });

        private static OfficeXmlNode GetOrCreatePageSetup(OfficeXmlNode worksheet) =>
            AddWorksheetObject(worksheet, "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties",
                "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
                "oleObjects", "controls", "webPublishItems", "tableParts");

        private static void ValidatePaperSize(ExcelPaperSize paperSize) {
            if (!Enum.IsDefined(typeof(ExcelPaperSize), paperSize)) {
                throw new ArgumentOutOfRangeException(nameof(paperSize), "Worksheet paper size must be a known Excel paper size.");
            }
        }

        /// <summary>Sets physical margins. Omitted header and footer margins are 0.3 inches.</summary>
        internal void SetMarginsCore(OfficeLength left, OfficeLength right, OfficeLength top, OfficeLength bottom, OfficeLength? header = null, OfficeLength? footer = null) {
            header ??= OfficeLength.Inches(0.3);
            footer ??= OfficeLength.Inches(0.3);
            if (left < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(left));
            if (right < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(right));
            if (top < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(top));
            if (bottom < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(bottom));
            if (header.Value < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(header));
            if (footer.Value < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(footer));
            EditWorksheetObject(worksheet => {
                OfficeXmlNode margins = AddWorksheetObject(worksheet, "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
                    "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
                    "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts");
                margins.SetDoubleAttribute("left", left.ToInches());
                margins.SetDoubleAttribute("right", right.ToInches());
                margins.SetDoubleAttribute("top", top.ToInches());
                margins.SetDoubleAttribute("bottom", bottom.ToInches());
                margins.SetDoubleAttribute("header", header.Value.ToInches());
                margins.SetDoubleAttribute("footer", footer.Value.ToInches());
                return true;
            });
        }

        /// <summary>
        /// Applies a preset set of margins.
        /// </summary>
        internal void SetMarginsPresetCore(ExcelMarginPreset preset) {
            // Values approximate Excel's presets in inches
            switch (preset) {
                case ExcelMarginPreset.Narrow:
                    SetMarginsCore(left: OfficeLength.Inches(0.25), right: OfficeLength.Inches(0.25), top: OfficeLength.Inches(0.5), bottom: OfficeLength.Inches(0.5), header: OfficeLength.Inches(0.3), footer: OfficeLength.Inches(0.3));
                    break;
                case ExcelMarginPreset.Moderate:
                    SetMarginsCore(left: OfficeLength.Inches(0.75), right: OfficeLength.Inches(0.75), top: OfficeLength.Inches(1.0), bottom: OfficeLength.Inches(1.0), header: OfficeLength.Inches(0.5), footer: OfficeLength.Inches(0.5));
                    break;
                case ExcelMarginPreset.Wide:
                    SetMarginsCore(left: OfficeLength.Inches(1.0), right: OfficeLength.Inches(1.0), top: OfficeLength.Inches(1.0), bottom: OfficeLength.Inches(1.0), header: OfficeLength.Inches(0.5), footer: OfficeLength.Inches(0.5));
                    break;
                default:
                    SetMarginsCore(left: OfficeLength.Inches(0.7), right: OfficeLength.Inches(0.7), top: OfficeLength.Inches(0.75), bottom: OfficeLength.Inches(0.75), header: OfficeLength.Inches(0.3), footer: OfficeLength.Inches(0.3));
                    break;
            }
        }

        internal void SetFitToPageDimension(bool width, uint? value) {
            Persistence.OwnsTablePrintDefaults = false;
            EditWorksheetObject(worksheet => {
                OfficeXmlNode pageSetup = GetOrCreatePageSetup(worksheet);
                pageSetup.SetUInt32Attribute(width ? "fitToWidth" : "fitToHeight", value);
                bool fit = pageSetup.HasAttribute("fitToWidth") || pageSetup.HasAttribute("fitToHeight");
                SetFitToPageMarkup(worksheet, fit ? true : null);
                return true;
            });
        }

        internal void ClearPageSetupAttribute(string localName) {
            Persistence.OwnsTablePrintDefaults = false;
            EditWorksheetObject(worksheet => worksheet.Child(WorksheetMarkup, "pageSetup").RemoveAttribute(localName));
        }

        internal void SetFitToPage(bool fitToPage, bool save = false) =>
            EditWorksheetObject(worksheet => {
                SetFitToPageMarkup(worksheet, fitToPage);
                return true;
            }, save);

        private static void SetFitToPageMarkup(OfficeXmlNode worksheet, bool? fitToPage) {
            OfficeXmlNode properties = worksheet.Child(WorksheetMarkup, "sheetPr");
            OfficeXmlNode setup = properties.Child(WorksheetMarkup, "pageSetUpPr");
            if (fitToPage.HasValue) {
                if (properties.IsNone) properties = worksheet.GetOrPrependChild(WorksheetMarkup, "sheetPr");
                if (setup.IsNone) setup = properties.Append(properties.CreateElement(WorksheetMarkup, "pageSetUpPr"));
            }
            if (!setup.IsNone) setup.SetBooleanAttribute("fitToPage", fitToPage);
        }

        internal void SetPageSetupCore(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null, ExcelPageOrder? pageOrder = null, ExcelPaperSize? paperSize = null) {
            Persistence.OwnsTablePrintDefaults = false;
            ApplyPageSetup(fitToWidth, fitToHeight, scale, pageOrder, paperSize, clearStaleFit: false);
        }

        internal void SetPageSetupAndClearStaleFit(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null, ExcelPageOrder? pageOrder = null, ExcelPaperSize? paperSize = null) =>
            ApplyPageSetup(fitToWidth, fitToHeight, scale, pageOrder, paperSize, clearStaleFit: true);

        private void ApplyPageSetup(uint? fitToWidth, uint? fitToHeight, uint? scale, ExcelPageOrder? pageOrder, ExcelPaperSize? paperSize, bool clearStaleFit) {
            if (scale is < 10U or > 400U) {
                throw new ArgumentOutOfRangeException(nameof(scale), "Manual print scale must be between 10 and 400 percent.");
            }
            if (paperSize.HasValue) ValidatePaperSize(paperSize.Value);

            EditWorksheetObject(worksheet => {
                OfficeXmlNode pageSetup = GetOrCreatePageSetup(worksheet);
                if (fitToWidth.HasValue || clearStaleFit) pageSetup.SetUInt32Attribute("fitToWidth", fitToWidth);
                if (fitToHeight.HasValue || clearStaleFit) pageSetup.SetUInt32Attribute("fitToHeight", fitToHeight);
                if (scale.HasValue || clearStaleFit) pageSetup.SetUInt32Attribute("scale", scale);
                if (paperSize.HasValue) pageSetup.SetUInt32Attribute("paperSize", (uint)paperSize.Value);
                if (pageOrder.HasValue) pageSetup.SetAttribute("pageOrder",
                    pageOrder == ExcelPageOrder.OverThenDown ? "overThenDown" : "downThenOver");
                if (fitToWidth.HasValue || fitToHeight.HasValue) SetFitToPageMarkup(worksheet, true);
                else if (clearStaleFit) SetFitToPageMarkup(worksheet, null);
                return true;
            });
        }
    }

    /// <summary>Common presets for page margins.</summary>
    public enum ExcelMarginPreset {
        /// <summary>Standard margins: left/right 0.7 inch, top/bottom 0.75 inch, header/footer 0.3 inch.</summary>
        Normal,
        /// <summary>Narrow margins: left/right 0.25 inch, top/bottom 0.5 inch, header/footer 0.3 inch.</summary>
        Narrow,
        /// <summary>Moderate margins: left/right 0.75 inch, top/bottom 1.0 inch, header/footer 0.5 inch.</summary>
        Moderate,
        /// <summary>Wide margins: left/right 1.0 inch, top/bottom 1.0 inch, header/footer 0.5 inch.</summary>
        Wide
    }

    /// <summary>Worksheet print page order.</summary>
    public enum ExcelPageOrder {
        /// <summary>Print pages down first, then over to the next column group.</summary>
        DownThenOver,
        /// <summary>Print pages over first, then down to the next row group.</summary>
        OverThenDown
    }

    /// <summary>Known OpenXML worksheet paper-size codes.</summary>
    public enum ExcelPaperSize : uint {
        /// <summary>US Letter, 8.5 x 11 inches.</summary>
        Letter = 1,
        /// <summary>US Letter Small, treated as Letter for image page geometry.</summary>
        LetterSmall = 2,
        /// <summary>US Tabloid, 11 x 17 inches.</summary>
        Tabloid = 3,
        /// <summary>US Ledger, 17 x 11 inches.</summary>
        Ledger = 4,
        /// <summary>US Legal, 8.5 x 14 inches.</summary>
        Legal = 5,
        /// <summary>US Statement, 5.5 x 8.5 inches.</summary>
        Statement = 6,
        /// <summary>US Executive, 7.25 x 10.5 inches.</summary>
        Executive = 7,
        /// <summary>ISO A3, 297 x 420 millimeters.</summary>
        A3 = 8,
        /// <summary>ISO A4, 210 x 297 millimeters.</summary>
        A4 = 9,
        /// <summary>ISO A4 Small, treated as A4 for image page geometry.</summary>
        A4Small = 10,
        /// <summary>ISO A5, 148 x 210 millimeters.</summary>
        A5 = 11,
        /// <summary>JIS B4, 257 x 364 millimeters.</summary>
        B4Jis = 12,
        /// <summary>JIS B5, 182 x 257 millimeters.</summary>
        B5Jis = 13
    }
}
