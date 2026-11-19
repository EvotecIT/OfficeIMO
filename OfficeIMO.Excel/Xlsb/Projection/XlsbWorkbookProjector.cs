using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects the supported XLSB workbook model into OfficeIMO's normal editable workbook surface.</summary>
    internal static class XlsbWorkbookProjector {
        internal static ExcelDocument ToExcelDocument(XlsbWorkbook workbook, System.Threading.CancellationToken cancellationToken = default,
            int maxCells = 4_000_000) {
            ArgumentNullException.ThrowIfNull(workbook);
            cancellationToken.ThrowIfCancellationRequested();
            EnsureArrayProjectionBudget(workbook, maxCells, cancellationToken);

            ExcelDocument document = ExcelDocument.Create();
            try {
                if (workbook.Uses1904DateSystem) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                }
                if (workbook.Stylesheet != null) {
                    XlsbStylesheetProjector.Install(document, workbook.Stylesheet);
                }
                if (workbook.WorkbookProtection != null) {
                    XlsbWorkbookProtectionProjector.Apply(document, workbook.WorkbookProtection);
                }

                var relatedParts = new XlsbRelatedPartProjector(document, workbook, cancellationToken);
                foreach (XlsbSheet sourceSheet in workbook.Sheets) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (sourceSheet.IsChartSheet) {
                        relatedParts.ProjectChartSheet(sourceSheet);
                        continue;
                    }
                    ExcelSheet targetSheet = document.AddSheet(sourceSheet.Name);
                    targetSheet.Batch(sheet => {
                        foreach (XlsbCell cell in sourceSheet.Cells) {
                            cancellationToken.ThrowIfCancellationRequested();
                            switch (cell.Kind) {
                                case XlsbCellValueKind.Blank:
                                    break;
                                case XlsbCellValueKind.Error:
                                    sheet.SetLegacyErrorCellValue(cell.Row, cell.Column, cell.Value as string ?? "#VALUE!");
                                    break;
                                default:
                                    sheet.CellValue(cell.Row, cell.Column, cell.Value);
                                    break;
                            }

                            if (!string.IsNullOrWhiteSpace(cell.FormulaText)) {
                                sheet.CellFormula(cell.Row, cell.Column, cell.FormulaText!);
                            }

                            sheet.SetCellStyleIndex(cell.Row, cell.Column, cell.StyleIndex, save: false);
                        }
                    });

                    ApplyRichText(targetSheet, sourceSheet, workbook.Stylesheet, cancellationToken);

                    var arrays = sourceSheet.FormulaGroups.Where(group => group.Type == 426).ToArray();
                    var sourceCells = arrays.Length == 0 ? new Dictionary<(int Row, int Column), XlsbCell>()
                        : sourceSheet.Cells.Where(cell => !string.IsNullOrWhiteSpace(cell.FormulaText)).ToDictionary(cell => (cell.Row, cell.Column));
                    foreach (var group in arrays) {
                        sourceCells.TryGetValue((group.AnchorRow, group.AnchorColumn), out var anchor);
                        if (string.IsNullOrWhiteSpace(anchor?.FormulaText)) continue;
                        targetSheet.SetLegacyArrayFormula(group.Range.ToA1Reference(), anchor!.FormulaText!, cancellationToken);
                    }
                    XlsbWorksheetPropertiesProjector.Apply(targetSheet, sourceSheet.Properties);
                    XlsbWorksheetGeometryProjector.Apply(targetSheet, sourceSheet);
                    XlsbWorksheetProtectionProjector.Apply(targetSheet, sourceSheet.Protection);
                    XlsbWorksheetAutoFilterProjector.Apply(targetSheet, sourceSheet.AutoFilter);
                    XlsbWorksheetPrintSettingsProjector.Apply(
                        targetSheet,
                        sourceSheet.PrintOptions,
                        sourceSheet.PageMargins,
                        sourceSheet.PageSetup,
                        sourceSheet.HeaderFooter);
                    XlsbWorksheetHyperlinkProjector.Apply(targetSheet, sourceSheet);
                    XlsbWorksheetFeatureProjector.Apply(targetSheet, sourceSheet);
                    XlsbWorksheetCommentProjector.Apply(targetSheet, sourceSheet, workbook.Stylesheet, relatedParts);
                    targetSheet.EnsureWorksheetElementOrder();
                    targetSheet.WorksheetPart.Worksheet?.Save();

                    if (sourceSheet.State == 1 || sourceSheet.State == 2) {
                        targetSheet.ApplyImportedVisibility(hidden: true, veryHidden: sourceSheet.State == 2);
                    }
                }

                // Carry the source workbook view's active and first tab, then move it off a hidden sheet if needed.
                document.ApplyImportedActiveTab(workbook.ActiveTab, workbook.FirstSheet, workbook.Sheets.Count);
                document.EnsureActiveTabVisible();
                relatedParts.ProjectTheme();
                relatedParts.ProjectVba();
                XlsbDefinedNameProjector.Apply(document, workbook.DefinedNames);

                if (workbook.CalculationProperties != null) {
                    XlsbCalculationPropertiesProjector.Apply(document, workbook.CalculationProperties);
                }

                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        // A StrRun covers its text up to the next run; text before the first run uses the cell's own font.
        private static void ApplyRichText(ExcelSheet targetSheet, XlsbSheet sourceSheet, XlsbStylesheet? stylesheet,
            System.Threading.CancellationToken cancellationToken) {
            foreach (XlsbCell cell in sourceSheet.Cells) {
                cancellationToken.ThrowIfCancellationRequested();
                DocumentFormat.OpenXml.Spreadsheet.InlineString? inline = CreateRichText(cell, stylesheet);
                if (inline != null) targetSheet.SetImportedRichText(cell.Row, cell.Column, inline);
            }
        }

        /// <summary>Builds the inline rich text projected for a cell with formatting runs, or null for other cells.</summary>
        internal static DocumentFormat.OpenXml.Spreadsheet.InlineString? CreateRichText(XlsbCell cell, XlsbStylesheet? stylesheet) {
            if (cell.TextRuns == null || cell.FormulaBytes != null || cell.Value is not string text) return null;
            var inline = new DocumentFormat.OpenXml.Spreadsheet.InlineString();
            if (cell.TextRuns[0].Start > 0) inline.Append(CreateRun(text.Substring(0, cell.TextRuns[0].Start), null));
            for (int index = 0; index < cell.TextRuns.Length; index++) {
                int start = cell.TextRuns[index].Start;
                int end = index + 1 < cell.TextRuns.Length ? cell.TextRuns[index + 1].Start : text.Length;
                ushort font = cell.TextRuns[index].Font;
                XlsbFont? source = stylesheet != null && font < stylesheet.Fonts.Count ? stylesheet.Fonts[font] : null;
                inline.Append(CreateRun(text.Substring(start, end - start),
                    source == null ? null : XlsbStylesheetProjector.CreateRunProperties(source)));
            }
            return inline;
        }

        private static DocumentFormat.OpenXml.Spreadsheet.Run CreateRun(string text, DocumentFormat.OpenXml.Spreadsheet.RunProperties? properties) {
            var run = new DocumentFormat.OpenXml.Spreadsheet.Run();
            if (properties != null) run.Append(properties);
            run.Append(new DocumentFormat.OpenXml.Spreadsheet.Text(Utilities.ExcelSanitizer.EncodeXString(text)) {
                Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
            });
            return run;
        }

        private static void EnsureArrayProjectionBudget(XlsbWorkbook workbook, int maxCells,
            System.Threading.CancellationToken cancellationToken) {
            long projectedCells = workbook.Sheets.Sum(sheet => (long)sheet.Cells.Count);
            if (projectedCells > maxCells) throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {maxCells} cells.");
            // Each referencing cell gets its own copy of a shared string's runs, so runs count against the same
            // projection budget as cells.
            foreach (XlsbSheet sheet in workbook.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (XlsbCell cell in sheet.Cells) {
                    if (cell.TextRuns == null || cell.FormulaBytes != null) continue;
                    projectedCells += cell.TextRuns.Length;
                    if (projectedCells > maxCells) {
                        throw new InvalidDataException($"The XLSB rich text exceeds the configured limit of {maxCells} projected cells and text runs.");
                    }
                }
            }
            foreach (XlsbSheet sheet in workbook.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (sheet.IsChartSheet || !sheet.FormulaGroups.Any(group => group.Type == 426)) continue;
                var formulaAnchors = new HashSet<(int Row, int Column)>();
                foreach (XlsbCell cell in sheet.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrWhiteSpace(cell.FormulaText)) formulaAnchors.Add((cell.Row, cell.Column));
                }
                foreach (var group in sheet.FormulaGroups) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (group.Type != 426 || !formulaAnchors.Contains((group.AnchorRow, group.AnchorColumn))) continue;
                    long area = (long)(group.Range.LastRow - group.Range.FirstRow + 1) *
                        (group.Range.LastColumn - group.Range.FirstColumn + 1);
                    // The anchor is already counted as an input cell. Counting other
                    // explicit members again is conservative and keeps the preflight cheap.
                    if (area - 1 > maxCells - projectedCells)
                        throw new InvalidDataException($"The XLSB array formulas exceed the configured limit of {maxCells} projected cells.");
                    projectedCells += area - 1;
                }
            }
        }
    }
}
