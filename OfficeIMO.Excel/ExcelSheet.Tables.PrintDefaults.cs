using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Printable width of a Letter page with Excel's default 0.7 inch side margins, in points.</summary>
        private const double PortraitPrintableWidthPoints = (8.5D - 1.4D) * 72D;

        /// <summary>Printable width of a landscape Letter page with the default side margins, in points.</summary>
        private const double LandscapePrintableWidthPoints = (11D - 1.4D) * 72D;

        /// <summary>
        /// Gives a sheet that holds a table print settings a reader expects: the header row
        /// repeats on every page, a table wider than a portrait page prints in landscape, and
        /// one wider than a landscape page is fitted to the page width. Settings the caller
        /// chose are never changed; defaults applied here are refreshed when columns are fitted.
        /// </summary>
        private void ApplyTablePrintDefaults(int headerRow) {
            ExcelPrintTitles titles = GetPrintTitles();
            bool ownsDefaults = Persistence.OwnsTablePrintDefaults;
            if (!ownsDefaults && HasAuthoredPageSetup()) return;
            if (!titles.HasRows && headerRow > 0) {
                _excelDocument.SetPrintTitles(this, headerRow, headerRow, null, null, save: false);
            }

            ApplyTableOrientationDefaults();
            Persistence.OwnsTablePrintDefaults = true;
        }

        /// <summary>Re-evaluates orientation and fit after column widths change.</summary>
        private void RefreshTablePrintDefaults() {
            if (!Persistence.OwnsTablePrintDefaults) return;
            ApplyTableOrientationDefaults();
            Persistence.OwnsTablePrintDefaults = true;
        }

        private void ApplyTableOrientationDefaults() {
            double width = MeasureTableColumnsWidthPoints();
            if (width <= 0D) return;
            bool landscape = width > PortraitPrintableWidthPoints;
            bool fit = width > LandscapePrintableWidthPoints;
            // A table that fits a portrait page needs no page setup unless an earlier
            // default has to be undone.
            if (!landscape && !Persistence.OwnsTablePrintDefaults) return;
            SetOrientation(landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait);
            if (fit) {
                SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
            } else {
                SetPageSetupAndClearStaleFit();
            }
        }

        private bool HasAuthoredPageSetup() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                PageSetup? pageSetup = WorksheetRoot.GetFirstChild<PageSetup>();
                if (pageSetup == null) return false;
                return pageSetup.Orientation != null || pageSetup.FitToWidth != null || pageSetup.FitToHeight != null ||
                    pageSetup.Scale != null || WorksheetRoot.GetFirstChild<SheetProperties>()?.GetFirstChild<PageSetupProperties>()?.FitToPage?.Value == true;
            });
        }

        private double MeasureTableColumnsWidthPoints() {
            int lastColumn = 0;
            foreach (var part in _worksheetPart.TableDefinitionParts) {
                if (part.Table?.Reference?.Value is string reference &&
                    ExcelA1.TryParseRange(reference, out _, out _, out _, out int tableLastColumn)) {
                    lastColumn = Math.Max(lastColumn, tableLastColumn);
                }
            }

            if (lastColumn == 0) return 0D;
            IReadOnlyList<ExcelColumnInfo> columns = GetColumnDefinitions();
            double defaultWidth = DefaultColumnWidth is ExcelColumnWidth configured && configured.ToCharacters(this) > 0D ? configured.ToCharacters(this) : 8.43D;
            double width = 0D;
            for (int column = 1; column <= lastColumn; column++) {
                double characters = defaultWidth;
                bool hidden = false;
                for (int index = columns.Count - 1; index >= 0; index--) {
                    ExcelColumnInfo definition = columns[index];
                    if (column < definition.StartIndex || column > definition.EndIndex) continue;
                    hidden = definition.Hidden;
                    if (definition.Width is ExcelColumnWidth custom && custom.ToCharacters(this) > 0D) characters = custom.ToCharacters(this);
                    break;
                }

                // Calibri 11 has a seven pixel digit; a pixel is 0.75 pt.
                if (!hidden) width += Math.Truncate(characters * 7D + 0.5D) * 0.75D;
            }

            return width;
        }
    }
}
