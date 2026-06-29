using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static IReadOnlyList<byte[]> BuildWindow1Payloads(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            IReadOnlyList<WorkbookView> workbookViews = document.WorkbookRoot
                .GetFirstChild<BookViews>()?
                .Elements<WorkbookView>()
                .ToArray() ?? Array.Empty<WorkbookView>();
            if (workbookViews.Count == 0) {
                return new[] { BuildWindow1Payload(null, sheets) };
            }

            var payloads = new List<byte[]>(workbookViews.Count);
            foreach (WorkbookView workbookView in workbookViews) {
                payloads.Add(BuildWindow1Payload(workbookView, sheets));
            }

            return payloads;
        }

        private static byte[] BuildWindow1Payload(WorkbookView? workbookView, IReadOnlyList<ExcelSheet> sheets) {
            ushort activeSheetIndex = CoerceSheetIndex(workbookView?.ActiveTab?.Value, sheets.Count, 0);
            ushort firstVisibleSheetTabIndex = CoerceSheetIndex(workbookView?.FirstSheet?.Value, sheets.Count, activeSheetIndex);
            ushort selectedSheetTabCount = CountSelectedSheetTabs(sheets);
            ushort sheetTabRatio = CoerceUInt16(workbookView?.TabRatio?.Value, 600);
            short horizontalPosition = CoerceInt16(workbookView?.XWindow?.Value, 0);
            short verticalPosition = CoerceInt16(workbookView?.YWindow?.Value, 0);
            short width = CoercePositiveInt16(workbookView?.WindowWidth?.Value, 0x4000);
            short height = CoercePositiveInt16(workbookView?.WindowHeight?.Value, 0x2000);

            ushort flags = 0;
            VisibilityValues? visibility = workbookView?.Visibility?.Value;
            if (visibility == VisibilityValues.Hidden || visibility == VisibilityValues.VeryHidden) {
                flags |= 0x0001;
            }

            if (workbookView?.Minimized?.Value == true) {
                flags |= 0x0002;
            }

            if (visibility == VisibilityValues.VeryHidden) {
                flags |= 0x0004;
            }

            if (workbookView?.ShowHorizontalScroll?.Value ?? true) {
                flags |= 0x0008;
            }

            if (workbookView?.ShowVerticalScroll?.Value ?? true) {
                flags |= 0x0010;
            }

            if (workbookView?.ShowSheetTabs?.Value ?? true) {
                flags |= 0x0020;
            }

            using var stream = new MemoryStream();
            WriteInt16(stream, horizontalPosition);
            WriteInt16(stream, verticalPosition);
            WriteInt16(stream, width);
            WriteInt16(stream, height);
            WriteUInt16(stream, flags);
            WriteUInt16(stream, activeSheetIndex);
            WriteUInt16(stream, firstVisibleSheetTabIndex);
            WriteUInt16(stream, selectedSheetTabCount);
            WriteUInt16(stream, sheetTabRatio);
            return stream.ToArray();
        }

        private static ushort CoerceSheetIndex(uint? index, int sheetCount, ushort fallback) {
            if (sheetCount <= 0) {
                return 0;
            }

            if (!index.HasValue || index.Value >= sheetCount) {
                return fallback < sheetCount ? fallback : (ushort)0;
            }

            return checked((ushort)index.Value);
        }

        private static ushort CoerceUInt16(uint? value, ushort fallback) {
            if (!value.HasValue) {
                return fallback;
            }

            return value.Value > ushort.MaxValue
                ? fallback
                : checked((ushort)value.Value);
        }

        private static short CoerceInt16(int? value, short fallback) {
            if (!value.HasValue || value.Value < short.MinValue || value.Value > short.MaxValue) {
                return fallback;
            }

            return checked((short)value.Value);
        }

        private static short CoercePositiveInt16(uint? value, short fallback) {
            if (!value.HasValue || value.Value == 0 || value.Value > short.MaxValue) {
                return fallback;
            }

            return checked((short)value.Value);
        }

        private static void WriteInt16(Stream stream, short value) {
            WriteUInt16(stream, unchecked((ushort)value));
        }

        private static ushort CountSelectedSheetTabs(IReadOnlyList<ExcelSheet> sheets) {
            int selectedCount = 0;
            foreach (ExcelSheet sheet in sheets) {
                SheetView? sheetView = sheet.WorksheetPart.Worksheet!
                    .GetFirstChild<SheetViews>()?
                    .Elements<SheetView>()
                    .FirstOrDefault();
                if (sheetView?.TabSelected?.Value == true) {
                    selectedCount++;
                }
            }

            if (selectedCount <= 0) {
                return 1;
            }

            return checked((ushort)Math.Min(selectedCount, sheets.Count));
        }

        private static byte[] BuildWindow2Payload(LegacyXlsWorksheetView view) {
            using var stream = new MemoryStream();
            ushort options = 0;
            if (view.ShowFormulas) {
                options |= 0x0001;
            }

            if (view.ShowGridlines) {
                options |= 0x0002;
            }

            if (view.ShowRowColumnHeadings) {
                options |= 0x0004;
            }

            if (view.FrozenRowCount > 0 || view.FrozenColumnCount > 0) {
                options |= 0x0008;
            }

            if (view.ShowZeroValues) {
                options |= 0x0010;
            }

            if (view.DefaultGridColor) {
                options |= 0x0020;
            }

            if (view.RightToLeft) {
                options |= 0x0040;
            }

            if (view.ShowOutlineSymbols) {
                options |= 0x0080;
            }

            if ((view.FrozenRowCount > 0 || view.FrozenColumnCount > 0) && view.FrozenWithoutSplit) {
                options |= 0x0100;
            }

            if (view.TabSelected) {
                options |= 0x0200;
            }

            if (view.PageBreakPreview) {
                options |= 0x0800;
            }

            WriteUInt16(stream, options);
            WriteUInt16(stream, view.TopLeftCell.Row);
            WriteUInt16(stream, view.TopLeftCell.Column);
            WriteUInt16(stream, view.GridLineColorIndex);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, view.PageBreakPreview ? GetWindow2ZoomOrDefault(view.ZoomScale, nameof(view.ZoomScale)) : (ushort)0);
            WriteUInt16(stream, GetWindow2ZoomOrDefault(view.ZoomScaleNormal, nameof(view.ZoomScaleNormal)));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static ushort GetWindow2ZoomOrDefault(uint? zoomScale, string name) {
            if (!zoomScale.HasValue) {
                return 0;
            }

            if (zoomScale.Value < 10U || zoomScale.Value > 400U) {
                throw new NotSupportedException($"Native XLS saving supports {name} from 10 through 400 percent; this worksheet uses {zoomScale.Value} percent.");
            }

            return checked((ushort)zoomScale.Value);
        }

        private static byte[] BuildZoomScalePayload(uint zoomScale) {
            if (zoomScale < 10U || zoomScale > 400U) {
                throw new NotSupportedException($"Native XLS saving supports worksheet zoom from 10 through 400 percent; this worksheet uses {zoomScale} percent.");
            }

            ushort numerator = checked((ushort)zoomScale);
            ushort denominator = 100;
            ushort divisor = GreatestCommonDivisor(numerator, denominator);
            numerator = checked((ushort)(numerator / divisor));
            denominator = checked((ushort)(denominator / divisor));

            using var stream = new MemoryStream();
            WriteUInt16(stream, numerator);
            WriteUInt16(stream, denominator);
            return stream.ToArray();
        }

        private static byte[] BuildPageLayoutViewPayload(uint? zoomScale) {
            ushort scale = GetPageLayoutZoomOrDefault(zoomScale);
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x088b);
            WriteUInt16(stream, 0x0005);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, scale);
            WriteUInt16(stream, 0x0001);
            return stream.ToArray();
        }

        private static void ReserveWorksheetTabColors(IReadOnlyList<ExcelSheet> sheets, LegacyXlsFontTable fontTable) {
            foreach (ExcelSheet sheet in sheets) {
                if (!TryGetSheetTabColorIndex(sheet, fontTable, out _, out string? reason)) {
                    throw new NotSupportedException($"Native XLS saving does not yet support {reason}. Save as .xlsx or remove this feature before saving as .xls.");
                }
            }
        }

        private static bool TryCreateSheetExtensionPayload(ExcelSheet sheet, LegacyXlsFontTable fontTable, out byte[]? payload) {
            payload = null;
            if (!TryGetSheetTabColorIndex(sheet, fontTable, out ushort colorIndex, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason}. Save as .xlsx or remove this feature before saving as .xls.");
            }

            if (colorIndex == 0x7fff) {
                return false;
            }

            if (colorIndex > 0x7f) {
                throw new NotSupportedException($"Native XLS saving does not yet support worksheet tab color index {colorIndex}. Save as .xlsx or remove this feature before saving as .xls.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0862);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 20);
            WriteUInt32(stream, colorIndex);
            payload = stream.ToArray();
            return true;
        }

        private static bool TryGetSheetTabColorIndex(ExcelSheet sheet, LegacyXlsFontTable fontTable, out ushort colorIndex, out string? reason) {
            colorIndex = 0x7fff;
            reason = null;
            TabColor? tabColor = sheet.WorksheetPart.Worksheet!
                .GetFirstChild<SheetProperties>()
                ?.TabColor;
            if (tabColor == null) {
                return true;
            }

            return fontTable.TryGetColorIndex(tabColor, "worksheet tab", "tab colors", out colorIndex, out reason);
        }

        private static ushort GetPageLayoutZoomOrDefault(uint? zoomScale) {
            if (!zoomScale.HasValue) {
                return 100;
            }

            if (zoomScale.Value < 10U || zoomScale.Value > 400U) {
                throw new NotSupportedException($"Native XLS saving supports Page Layout worksheet zoom from 10 through 400 percent; this worksheet uses {zoomScale.Value} percent.");
            }

            return checked((ushort)zoomScale.Value);
        }

        private static ushort GreatestCommonDivisor(ushort left, ushort right) {
            while (right != 0) {
                ushort remainder = (ushort)(left % right);
                left = right;
                right = remainder;
            }

            return left == 0 ? (ushort)1 : left;
        }

        private static byte[] BuildPanePayload(int frozenColumnCount, int frozenRowCount) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)frozenColumnCount));
            WriteUInt16(stream, checked((ushort)frozenRowCount));
            WriteUInt16(stream, checked((ushort)frozenRowCount));
            WriteUInt16(stream, checked((ushort)frozenColumnCount));
            stream.WriteByte(ResolveActivePane(frozenRowCount, frozenColumnCount));
            stream.WriteByte(0);
            return stream.ToArray();
        }

        private static byte[] BuildPanePayload(LegacyXlsSplitPaneView splitPane) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, splitPane.HorizontalSplit);
            WriteUInt16(stream, splitPane.VerticalSplit);
            WriteUInt16(stream, splitPane.TopRow);
            WriteUInt16(stream, splitPane.LeftColumn);
            stream.WriteByte(splitPane.ActivePane);
            stream.WriteByte(0);
            return stream.ToArray();
        }

        private static byte ResolveActivePane(int frozenRowCount, int frozenColumnCount) {
            if (frozenRowCount > 0 && frozenColumnCount > 0) {
                return 0;
            }

            if (frozenRowCount > 0) {
                return 2;
            }

            return 1;
        }
    }
}
