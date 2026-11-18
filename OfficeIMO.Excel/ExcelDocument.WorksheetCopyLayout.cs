using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        // Worksheet children that describe the sheet's layout and printing rather than its cell data. A values copy writes
        // the cells itself and carries these over, so the copy keeps its widths, frozen panes, merged ranges, rules and
        // page setup. Relationship-backed features (drawings, comments, hyperlinks) and filters stay package-mode only.
        private static readonly HashSet<Type> ValuesCopyLayoutElementTypes = new HashSet<Type> {
            typeof(SheetProperties),
            typeof(SheetViews),
            typeof(SheetFormatProperties),
            typeof(Columns),
            typeof(MergeCells),
            typeof(ConditionalFormatting),
            typeof(DataValidations),
            typeof(PrintOptions),
            typeof(PageMargins),
            typeof(PageSetup),
            typeof(HeaderFooter),
            typeof(RowBreaks),
            typeof(ColumnBreaks)
        };

        // Row attributes that carry layout. Cell spans and the row index belong to the cells the values copy wrote.
        private static readonly string[] ValuesCopyRowLayoutAttributes = {
            "s", "customFormat", "ht", "hidden", "customHeight", "outlineLevel", "collapsed", "thickTop", "thickBot"
        };

        // The sheet-local built-in names that belong to a worksheet's print settings.
        private static readonly string[] CopiedPrintDefinedNames = { "_xlnm.Print_Area", "_xlnm.Print_Titles" };

        /// <summary>
        /// Copies the source worksheet's layout onto a sheet filled by a values copy: sheet properties and views, default
        /// sizes, column definitions, row heights and visibility, merged cells, conditional formats, data validations and
        /// print settings. Column, row and conditional-format styles are mapped into this workbook's stylesheet.
        /// </summary>
        /// <returns>The style maps, which the copied tables' differential formats also use.</returns>
        private WorksheetStyleCopyMap CopyWorksheetLayoutForValuesCopy(ExcelDocument sourceDocument, ExcelSheet sourceSheet, ExcelSheet targetSheet) {
            WorksheetPart sourcePart = sourceSheet.WorksheetPart;
            WorksheetLayoutSnapshot layout = ReadWorksheetLayout(sourcePart);
            bool needsStyles = !ReferenceEquals(sourceDocument, this)
                && (layout.UsesStyles || sourcePart.TableDefinitionParts.Any());
            if (layout.Elements.Count == 0 && layout.Rows.Count == 0 && !needsStyles) {
                return WorksheetStyleCopyMap.Empty;
            }

            return Locking.ExecuteWrite(EnsureLock(), () => {
                Worksheet worksheet = targetSheet.WorksheetPart.Worksheet!;
                WorksheetStyleCopyMap styleMap = needsStyles
                    ? MergeSourceWorkbookStyles(sourceDocument.WorkbookPartRoot, WorkbookPartRoot)
                    : WorksheetStyleCopyMap.Empty;

                bool? targetTabSelected = worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.TabSelected?.Value;
                foreach (OpenXmlElement element in layout.Elements) {
                    foreach (OpenXmlElement existing in worksheet.ChildElements.Where(child => child.GetType() == element.GetType()).ToList()) {
                        existing.Remove();
                    }
                }

                foreach (OpenXmlElement element in layout.Elements) {
                    switch (element) {
                        case SheetProperties properties:
                            // A code name binds the sheet to a VBA module that a values copy does not bring along.
                            properties.CodeName = null;
                            properties.FilterMode = null;
                            break;
                        case Columns columns:
                            foreach (Column column in columns.Elements<Column>()) {
                                column.Style = MapCopiedStyleIndex(column.Style?.Value, styleMap.CellFormats);
                            }
                            break;
                        case PageSetup pageSetup:
                            // Printer settings are a related part that is not copied.
                            pageSetup.Id = null;
                            break;
                    }

                    worksheet.Append(element);
                }

                if (!ReferenceEquals(sourceDocument, this)) {
                    BindCopiedSheetViewsToWorkbookViews(worksheet);
                    BindCopiedWorksheetColors(sourceDocument.WorkbookPartRoot, WorkbookPartRoot, worksheet);
                }

                // The new sheet keeps the tab selection it was created with, whatever the source had.
                bool firstView = true;
                foreach (SheetView view in worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>() ?? Enumerable.Empty<SheetView>()) {
                    view.TabSelected = firstView ? targetTabSelected : null;
                    firstView = false;
                }

                RemapCopiedWorksheetDifferentialFormats(worksheet, styleMap.DifferentialFormats);
                ApplyCopiedRowLayout(worksheet, layout.Rows, styleMap.CellFormats);
                targetSheet.EnsureWorksheetElementOrder(save: false);
                worksheet.Save();
                return styleMap;
            });
        }

        /// <summary>
        /// Keeps a copied worksheet's views pointing at workbook views this workbook has. A sheet view belongs to the
        /// workbook view with its <c>workbookViewId</c>, so views for workbook windows the target lacks are dropped.
        /// When none of the copied views belongs to the first workbook view, the first copied view is bound to it so the
        /// sheet keeps its panes and zoom in the main window.
        /// </summary>
        private void BindCopiedSheetViewsToWorkbookViews(Worksheet worksheet) {
            SheetViews? views = worksheet.GetFirstChild<SheetViews>();
            if (views == null) {
                return;
            }

            uint workbookViewCount = (uint)Math.Max(1, WorkbookRoot.GetFirstChild<BookViews>()?.Elements<WorkbookView>().Count() ?? 0);
            var bound = new HashSet<uint>();
            SheetView? firstUnbound = null;
            foreach (SheetView view in views.Elements<SheetView>().ToList()) {
                uint id = view.WorkbookViewId?.Value ?? 0U;
                if (id < workbookViewCount && bound.Add(id)) {
                    continue;
                }

                view.Remove();
                firstUnbound ??= view;
            }

            if (!bound.Contains(0U) && firstUnbound != null) {
                firstUnbound.WorkbookViewId = 0U;
                views.PrependChild(firstUnbound);
            }

            if (!views.Elements<SheetView>().Any()) {
                views.Remove();
            }
        }

        /// <summary>
        /// Gives a worksheet copied from another workbook the source's tab, color-scale and data-bar colors where the
        /// target's theme or palette would show its theme and indexed colors differently. Cell data carries no colors.
        /// </summary>
        private static void BindCopiedWorksheetColors(WorkbookPart sourceWorkbookPart, WorkbookPart targetWorkbookPart, Worksheet worksheet) {
            var sourceColors = WorkbookStyleColorResolver.Create(sourceWorkbookPart, sourceWorkbookPart.WorkbookStylesPart?.Stylesheet ?? new Stylesheet());
            var targetColors = WorkbookStyleColorResolver.Create(targetWorkbookPart, targetWorkbookPart.WorkbookStylesPart?.Stylesheet ?? new Stylesheet());
            foreach (OpenXmlElement child in worksheet.ChildElements) {
                if (child is not SheetData) {
                    sourceColors.BindColorsTo(child, targetColors);
                }
            }
        }

        private static UInt32Value? MapCopiedStyleIndex(uint? styleIndex, IReadOnlyDictionary<uint, uint> cellFormatMap) {
            if (!styleIndex.HasValue) {
                return null;
            }

            return cellFormatMap.TryGetValue(styleIndex.Value, out uint mapped) ? mapped : styleIndex.Value;
        }

        private static void ApplyCopiedRowLayout(Worksheet worksheet, IReadOnlyList<Row> sourceRows, IReadOnlyDictionary<uint, uint> cellFormatMap) {
            if (sourceRows.Count == 0) {
                return;
            }

            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
            Row? current = sheetData.GetFirstChild<Row>();
            foreach (Row sourceRow in sourceRows.OrderBy(row => row.RowIndex!.Value)) {
                uint rowIndex = sourceRow.RowIndex!.Value;
                while (current != null && (current.RowIndex?.Value ?? 0U) < rowIndex) {
                    current = current.NextSibling<Row>();
                }

                Row target;
                if (current != null && current.RowIndex?.Value == rowIndex) {
                    target = current;
                } else {
                    target = new Row { RowIndex = rowIndex };
                    if (current != null) {
                        sheetData.InsertBefore(target, current);
                    } else {
                        sheetData.Append(target);
                    }
                }

                foreach (OpenXmlAttribute attribute in sourceRow.GetAttributes()) {
                    if (attribute.LocalName != "r") {
                        target.SetAttribute(attribute);
                    }
                }

                target.StyleIndex = MapCopiedStyleIndex(sourceRow.StyleIndex?.Value, cellFormatMap);
            }
        }

        private static WorksheetLayoutSnapshot ReadWorksheetLayout(WorksheetPart part) {
            var layout = new WorksheetLayoutSnapshot();
            if (part.IsRootElementLoaded) {
                foreach (OpenXmlElement child in part.Worksheet!.ChildElements) {
                    if (ValuesCopyLayoutElementTypes.Contains(child.GetType())) {
                        layout.Add(child.CloneNode(true));
                    } else if (child is SheetData sheetData) {
                        int previousRowIndex = 0;
                        foreach (Row row in sheetData.Elements<Row>()) {
                            int rowIndex = ExcelWorksheetCoordinates.GetRowIndex(row, ref previousRowIndex);
                            layout.AddRow(rowIndex > 0 ? (uint)rowIndex : null, row.GetAttributes());
                        }
                    }
                }

                return layout;
            }

            using OpenXmlReader reader = OpenXmlReader.Create(part);
            int previousStreamedRowIndex = 0;
            bool hasCurrent = reader.Read();
            while (hasCurrent) {
                if (reader.IsStartElement && reader.Depth == 1 && ValuesCopyLayoutElementTypes.Contains(reader.ElementType)) {
                    layout.Add(reader.LoadCurrentElement()!);
                } else if (reader.IsStartElement && reader.Depth == 2 && reader.ElementType == typeof(Row)) {
                    IReadOnlyList<OpenXmlAttribute> attributes = reader.Attributes;
                    string? rowIndexText = attributes.FirstOrDefault(attribute => attribute.LocalName == "r" && string.IsNullOrEmpty(attribute.NamespaceUri)).Value;
                    if (uint.TryParse(rowIndexText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out uint rowIndex) && rowIndex > 0U) {
                        previousStreamedRowIndex = checked((int)rowIndex);
                        layout.AddRow(rowIndex, attributes);
                    } else {
                        // A row without its optional index takes it from its first cell or follows the previous row,
                        // as the values reader does, so its layout lands on the row its values were written to.
                        var row = (Row)reader.LoadCurrentElement()!;
                        int resolved = ExcelWorksheetCoordinates.GetRowIndex(row, ref previousStreamedRowIndex);
                        layout.AddRow(resolved > 0 ? (uint)resolved : null, row.GetAttributes());
                    }
                }

                hasCurrent = reader.Read();
            }

            return layout;
        }

        /// <summary>
        /// Copies the source worksheet's print area and print titles to the copied worksheet, pointing references to copied
        /// sheets at their new names.
        /// </summary>
        private void CopyPrintDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            string sourceSheetName,
            ExcelSheet targetSheet,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            IReadOnlyDictionary<int, int>? externalReferenceMap = null) {
            IReadOnlyList<DefinedName> printNames = GetSourcePrintDefinedNames(sourceDocument, sourceSheetName);
            if (printNames.Count == 0) {
                return;
            }

            Locking.ExecuteWrite(EnsureLock(), () => {
                ushort targetPosition = GetSheetPositionIndex(targetSheet);
                DefinedNames targetNames = WorkbookRoot.DefinedNames ??= new DefinedNames();
                foreach (DefinedName sourceName in printNames) {
                    foreach (DefinedName existing in targetNames.Elements<DefinedName>()
                        .Where(name => name.LocalSheetId?.Value == targetPosition
                            && string.Equals(name.Name?.Value, sourceName.Name!.Value, StringComparison.OrdinalIgnoreCase))
                        .ToList()) {
                        existing.Remove();
                    }

                    var copy = (DefinedName)sourceName.CloneNode(true);
                    copy.LocalSheetId = targetPosition;
                    copy.Text = RewriteCopiedFormula(copy.Text, sheetNameMap, tableNameMap, externalReferenceMap);

                    targetNames.Append(copy);
                }

                WorkbookRoot.Save();
            });
        }

        /// <summary>Returns the source worksheet's print area and print titles.</summary>
        private static IReadOnlyList<DefinedName> GetSourcePrintDefinedNames(ExcelDocument sourceDocument, string sourceSheetName) {
            DefinedNames? sourceNames = sourceDocument.WorkbookRoot.DefinedNames;
            if (sourceNames == null || !sourceDocument.TryGetSheetPositionIndexByName(sourceSheetName, out ushort sourcePosition)) {
                return Array.Empty<DefinedName>();
            }

            return sourceNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId?.Value == sourcePosition
                    && CopiedPrintDefinedNames.Contains(name.Name?.Value, StringComparer.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(name.Text))
                .ToList();
        }

        private sealed class WorksheetLayoutSnapshot {
            internal List<OpenXmlElement> Elements { get; } = new List<OpenXmlElement>();

            internal List<Row> Rows { get; } = new List<Row>();

            internal bool UsesStyles { get; private set; }

            internal void Add(OpenXmlElement element) {
                Elements.Add(element);
                if (element is Columns columns && columns.Elements<Column>().Any(column => column.Style != null)) {
                    UsesStyles = true;
                } else if (element is ConditionalFormatting formatting
                    && formatting.Elements<ConditionalFormattingRule>().Any(rule => rule.FormatId != null)) {
                    UsesStyles = true;
                }
            }

            internal void AddRow(uint? rowIndex, IEnumerable<OpenXmlAttribute> attributes) {
                if (!rowIndex.HasValue) {
                    return;
                }

                Row? layoutRow = null;
                foreach (OpenXmlAttribute attribute in attributes) {
                    if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                        || Array.IndexOf(ValuesCopyRowLayoutAttributes, attribute.LocalName) < 0) {
                        continue;
                    }

                    layoutRow ??= new Row { RowIndex = rowIndex.Value };
                    layoutRow.SetAttribute(new OpenXmlAttribute(attribute.Prefix ?? string.Empty, attribute.LocalName, string.Empty, attribute.Value));
                    if (attribute.LocalName == "s") {
                        UsesStyles = true;
                    }
                }

                if (layoutRow != null) {
                    Rows.Add(layoutRow);
                }
            }
        }
    }
}
