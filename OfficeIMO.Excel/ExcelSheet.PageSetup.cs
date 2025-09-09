using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Page setup helpers for orientation and margins.
    /// </summary>
    public partial class ExcelSheet
    {
        /// <summary>
        /// Sets page orientation (Portrait or Landscape) on the sheet's PageSetup.
        /// </summary>
        public void SetOrientation(ExcelPageOrientation orientation)
        {
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var pageSetup = ws.GetFirstChild<PageSetup>();
                if (pageSetup == null)
                {
                    pageSetup = new PageSetup();
                    var margins = ws.GetFirstChild<PageMargins>();
                    if (margins != null) ws.InsertAfter(pageSetup, margins); else ws.Append(pageSetup);
                }
                string val = orientation == ExcelPageOrientation.Landscape ? "landscape" : "portrait";
                pageSetup.SetAttribute(new OpenXmlAttribute("", "orientation", "", val));
                ws.Save();
            });
        }

        /// <summary>
        /// Sets page margins in inches.
        /// </summary>
        public void SetMargins(double left, double right, double top, double bottom, double header = 0.3, double footer = 0.3)
        {
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var margins = ws.GetFirstChild<PageMargins>();
                if (margins == null)
                {
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
        public void SetMarginsPreset(ExcelMarginPreset preset)
        {
            // Values approximate Excel's presets in inches
            switch (preset)
            {
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
    public enum ExcelMarginPreset
    {
        Normal,
        Narrow,
        Moderate,
        Wide
    }

    /// <summary>Simple orientation enum for Excel worksheets.</summary>
    public enum ExcelPageOrientation
    {
        Portrait,
        Landscape
    }
}
