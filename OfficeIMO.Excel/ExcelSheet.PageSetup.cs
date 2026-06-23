using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Page setup helpers for orientation and margins.
    /// </summary>
    public partial class ExcelSheet {
        /// <summary>
        /// Sets page orientation (Portrait or Landscape) on the sheet's PageSetup.
        /// </summary>
        public void SetOrientation(ExcelPageOrientation orientation) {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = GetOrCreatePageSetup(ws);
                string val = orientation == ExcelPageOrientation.Landscape ? "landscape" : "portrait";
                pageSetup.SetAttribute(new OpenXmlAttribute("", "orientation", "", val));
                ws.Save();
            });
        }

        /// <summary>
        /// Sets a known worksheet paper size on the sheet's PageSetup.
        /// </summary>
        public void SetPaperSize(ExcelPaperSize paperSize) {
            ValidatePaperSize(paperSize);
            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = GetOrCreatePageSetup(ws);
                pageSetup.PaperSize = (uint)paperSize;
                ws.Save();
            });
        }

        private static PageSetup GetOrCreatePageSetup(Worksheet ws) {
            var pageSetup = ws.GetFirstChild<PageSetup>();
            if (pageSetup != null) {
                return pageSetup;
            }

            pageSetup = new PageSetup();
            var margins = ws.GetFirstChild<PageMargins>();
            if (margins != null) ws.InsertAfter(pageSetup, margins); else ws.Append(pageSetup);
            return pageSetup;
        }

        private static void ValidatePaperSize(ExcelPaperSize paperSize) {
            if (!Enum.IsDefined(typeof(ExcelPaperSize), paperSize)) {
                throw new ArgumentOutOfRangeException(nameof(paperSize), "Worksheet paper size must be a known Excel paper size.");
            }
        }

        /// <summary>
        /// Sets page margins in inches.
        /// </summary>
        public void SetMargins(double left, double right, double top, double bottom, double header = 0.3, double footer = 0.3) {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var margins = ws.GetFirstChild<PageMargins>();
                if (margins == null) {
                    margins = new PageMargins();
                    ws.InsertAt(margins, 0);
                }
                margins.Left = left; margins.Right = right; margins.Top = top; margins.Bottom = bottom; margins.Header = header; margins.Footer = footer;
                ws.Save();
            });
        }

        /// <summary>
        /// Applies a preset set of margins.
        /// </summary>
        public void SetMarginsPreset(ExcelMarginPreset preset) {
            // Values approximate Excel's presets in inches
            switch (preset) {
                case ExcelMarginPreset.Narrow:
                    SetMargins(left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3);
                    break;
                case ExcelMarginPreset.Moderate:
                    SetMargins(left: 0.75, right: 0.75, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5);
                    break;
                case ExcelMarginPreset.Wide:
                    SetMargins(left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5);
                    break;
                default:
                    SetMargins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3);
                    break;
            }
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

    /// <summary>Sheet page orientation.</summary>
    public enum ExcelPageOrientation {
        /// <summary>Portrait orientation (vertical).</summary>
        Portrait,
        /// <summary>Landscape orientation (horizontal).</summary>
        Landscape
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
