using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPageSettingsTests {
        [Fact]
        public void PageSettingsSaveLoadAndRoundTripTypedCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Settings", 11, 8.5);
            page.SetMargins(0.4, 0.5, 0.6, 0.7);
            page.PrintOrientation = VisioPagePrintOrientation.Landscape;
            page.PageLockReplace = true;
            page.PageLockDuplicate = true;
            page.AddRectangle(5.5, 4.25, 2, 1, "Configured");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPageSettingsXml(filePath, "Settings", "0.4", "0.5", "0.6", "0.7", "2", "1", "1");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Settings");
            Assert.Equal(0.4, loadedPage.LeftMargin, 6);
            Assert.Equal(0.5, loadedPage.RightMargin, 6);
            Assert.Equal(0.6, loadedPage.TopMargin, 6);
            Assert.Equal(0.7, loadedPage.BottomMargin, 6);
            Assert.Equal(VisioPagePrintOrientation.Landscape, loadedPage.PrintOrientation);
            Assert.True(loadedPage.PageLockReplace);
            Assert.True(loadedPage.PageLockDuplicate);

            loadedPage.SetMargins(1, VisioMeasurementUnit.Centimeters);
            loadedPage.PrintOrientation = VisioPagePrintOrientation.Portrait;
            loadedPage.PageLockDuplicate = false;
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPageSettingsXml(roundTripPath, "Settings", "1", "1", "1", "1", "1", "1", "0", unit: "CM");

            VisioDocument reloadedRoundTrip = VisioDocument.Load(roundTripPath);
            VisioPage reloadedSettings = reloadedRoundTrip.Pages.Single(current => current.Name == "Settings");
            Assert.Equal(1D.ToInches(VisioMeasurementUnit.Centimeters), reloadedSettings.LeftMargin, 6);
            Assert.Equal(1D.ToInches(VisioMeasurementUnit.Centimeters), reloadedSettings.RightMargin, 6);
            Assert.Equal(1D.ToInches(VisioMeasurementUnit.Centimeters), reloadedSettings.TopMargin, 6);
            Assert.Equal(1D.ToInches(VisioMeasurementUnit.Centimeters), reloadedSettings.BottomMargin, 6);
        }

        [Fact]
        public void NonDefaultUnitPageDoesNotForcePrintOrientationWhenUnset() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Metric", 21, 29.7, VisioMeasurementUnit.Centimeters);
            page.AddRectangle(10.5, 14.85, 4, 2, "Metric", VisioMeasurementUnit.Centimeters);
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPageSettingsXml(filePath, "Metric", "0.635", "0.635", "0.635", "0.635", null, "0", "0", unit: "CM");
        }

        [Fact]
        public void PageSetupBehaviorCellsSaveLoadAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Setup", 8.5, 11);
            page.DrawingSizeType = VisioDrawingSizeType.Custom;
            page.AutoResizeDrawing = false;
            page.AllowShapeSplitting = false;
            page.UiVisibility = VisioPageUiVisibility.Hidden;
            page.AddRectangle(4.25, 5.5, 2, 1, "Setup");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPageSetupXml(filePath, "Setup", "3", "0", "0", "1");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Setup");
            Assert.Equal(VisioDrawingSizeType.Custom, loadedPage.DrawingSizeType);
            Assert.False(loadedPage.AutoResizeDrawing);
            Assert.False(loadedPage.AllowShapeSplitting);
            Assert.Equal(VisioPageUiVisibility.Hidden, loadedPage.UiVisibility);

            loadedPage.DrawingSizeType = VisioDrawingSizeType.FitToDrawingContents;
            loadedPage.AutoResizeDrawing = true;
            loadedPage.AllowShapeSplitting = true;
            loadedPage.UiVisibility = VisioPageUiVisibility.Normal;
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPageSetupXml(roundTripPath, "Setup", "1", "1", "1", "0");
        }

        [Fact]
        public void PageConnectorRoutingPolicyCellsSaveLoadAndCanBeCleared() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string clearedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Routing", 11, 8.5);
            page.ConnectorRouteStyle = VisioPageRouteStyle.FlowchartTopToBottom;
            page.ConnectorRouteAppearance = VisioLineRouteExtension.Curved;
            page.LineJumpStyle = VisioLineJumpStyle.Gap;
            page.LineJumpCode = VisioLineJumpCode.DisplayOrder;
            page.HorizontalLineJumpDirection = VisioHorizontalLineJumpDirection.Up;
            page.VerticalLineJumpDirection = VisioVerticalLineJumpDirection.Right;
            page.SetConnectorSpacing(0.4, 0.5, 0.6, 0.7);
            page.AddRectangle(3, 4, 1.5, 0.8, "Start");
            page.AddRectangle(8, 4, 1.5, 0.8, "End");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPageRoutingXml(filePath, "Routing", "5", "2", "2", "4", "1", "2");
            AssertPageRoutingSpacingXml(filePath, "Routing", "0.4", "0.5", "0.6", "0.7");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Routing");
            Assert.Equal(VisioPageRouteStyle.FlowchartTopToBottom, loadedPage.ConnectorRouteStyle);
            Assert.Equal(VisioLineRouteExtension.Curved, loadedPage.ConnectorRouteAppearance);
            Assert.Equal(VisioLineJumpStyle.Gap, loadedPage.LineJumpStyle);
            Assert.Equal(VisioLineJumpCode.DisplayOrder, loadedPage.LineJumpCode);
            Assert.Equal(VisioHorizontalLineJumpDirection.Up, loadedPage.HorizontalLineJumpDirection);
            Assert.Equal(VisioVerticalLineJumpDirection.Right, loadedPage.VerticalLineJumpDirection);
            Assert.Equal(0.4, loadedPage.LineToLineX.GetValueOrDefault(), 6);
            Assert.Equal(0.5, loadedPage.LineToLineY.GetValueOrDefault(), 6);
            Assert.Equal(0.6, loadedPage.LineToNodeX.GetValueOrDefault(), 6);
            Assert.Equal(0.7, loadedPage.LineToNodeY.GetValueOrDefault(), 6);

            loadedPage.ConnectorRouteStyle = VisioPageRouteStyle.Network;
            loadedPage.ConnectorRouteAppearance = VisioLineRouteExtension.Straight;
            loadedPage.LineJumpStyle = VisioLineJumpStyle.Square;
            loadedPage.LineJumpCode = VisioLineJumpCode.Horizontal;
            loadedPage.HorizontalLineJumpDirection = VisioHorizontalLineJumpDirection.Down;
            loadedPage.VerticalLineJumpDirection = VisioVerticalLineJumpDirection.Left;
            loadedPage.SetConnectorSpacing(8, 9, 10, 11, VisioMeasurementUnit.Millimeters);
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPageRoutingXml(roundTripPath, "Routing", "9", "1", "3", "1", "2", "1");
            AssertPageRoutingSpacingXml(roundTripPath, "Routing", "8", "9", "10", "11", "MM");

            VisioDocument cleared = VisioDocument.Load(roundTripPath);
            VisioPage clearedPage = cleared.Pages.Single(current => current.Name == "Routing");
            clearedPage.ClearConnectorRoutingPolicy();
            clearedPage.ClearConnectorSpacing();
            cleared.Save(clearedPath);

            Assert.Empty(VisioValidator.Validate(clearedPath));
            AssertPageRoutingXml(clearedPath, "Routing", null, null, null, null, null, null);
            AssertPageRoutingSpacingXml(clearedPath, "Routing", null, null, null, null);
        }

        [Fact]
        public void PagePlacementPolicyCellsSaveLoadAndCanBeCleared() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string clearedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Placement", 11, 8.5);
            page.PlacementStyle = VisioPlacementStyle.HierarchyTopToBottomCenter;
            page.PlacementDepth = VisioPlacementDepth.Deep;
            page.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
            page.MoveShapesAwayOnDrop = true;
            page.ResizePageToFitLayout = false;
            page.AddRectangle(3, 5, 1.5, 0.8, "Parent");
            page.AddRectangle(7, 5, 1.5, 0.8, "Child");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPagePlacementXml(filePath, "Placement", "17", "2", "5", "1", "0");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Placement");
            Assert.Equal(VisioPlacementStyle.HierarchyTopToBottomCenter, loadedPage.PlacementStyle);
            Assert.Equal(VisioPlacementDepth.Deep, loadedPage.PlacementDepth);
            Assert.Equal(VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90, loadedPage.PlacementFlip);
            Assert.True(loadedPage.MoveShapesAwayOnDrop);
            Assert.False(loadedPage.ResizePageToFitLayout);

            loadedPage.PlacementStyle = VisioPlacementStyle.LeftToRight;
            loadedPage.PlacementDepth = VisioPlacementDepth.Shallow;
            loadedPage.PlacementFlip = VisioPlacementFlip.None;
            loadedPage.MoveShapesAwayOnDrop = false;
            loadedPage.ResizePageToFitLayout = true;
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPagePlacementXml(roundTripPath, "Placement", "2", "3", "8", "0", "1");

            VisioDocument cleared = VisioDocument.Load(roundTripPath);
            cleared.Pages.Single(current => current.Name == "Placement").ClearPlacementPolicy();
            cleared.Save(clearedPath);

            Assert.Empty(VisioValidator.Validate(clearedPath));
            AssertPagePlacementXml(clearedPath, "Placement", null, null, null, null, null);
        }

        [Fact]
        public void ConnectorSpacingLoadsEachCellWithItsDeclaredUnit() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Mixed Routing", 11, 8.5);
            page.SetConnectorSpacing(1, 1, 1, 1);
            document.Save();

            RewritePageSheetCells(filePath, "Mixed Routing", (pageSheet, ns) => {
                SetPageSheetCell(pageSheet, ns, "LineToLineX", "2", "IN");
                SetPageSheetCell(pageSheet, ns, "LineToLineY", "5.08", "CM");
                SetPageSheetCell(pageSheet, ns, "LineToNodeX", "50.8", "MM");
                SetPageSheetCell(pageSheet, ns, "LineToNodeY", "2", "IN");
            });

            VisioPage loadedPage = VisioDocument.Load(filePath).Pages.Single(current => current.Name == "Mixed Routing");
            Assert.Equal(2, loadedPage.LineToLineX.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LineToLineY.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LineToNodeX.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LineToNodeY.GetValueOrDefault(), 6);
        }

        [Fact]
        public void PageLayoutGridCellsSaveLoadAndCanBeCleared() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string clearedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Grid", 11, 8.5);
            page.EnableLayoutGrid = true;
            page.SetLayoutGridSizing(1.2, 1.1, 0.35, 0.4);
            page.AddRectangle(3, 5, 1.5, 0.8, "One");
            page.AddRectangle(7, 5, 1.5, 0.8, "Two");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPageLayoutGridXml(filePath, "Grid", "1", "1.2", "1.1", "0.35", "0.4");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Grid");
            Assert.True(loadedPage.EnableLayoutGrid);
            Assert.Equal(1.2, loadedPage.LayoutBlockSizeX.GetValueOrDefault(), 6);
            Assert.Equal(1.1, loadedPage.LayoutBlockSizeY.GetValueOrDefault(), 6);
            Assert.Equal(0.35, loadedPage.LayoutAvenueSizeX.GetValueOrDefault(), 6);
            Assert.Equal(0.4, loadedPage.LayoutAvenueSizeY.GetValueOrDefault(), 6);

            loadedPage.EnableLayoutGrid = false;
            loadedPage.SetLayoutGridSizing(20, 21, 8, 9, VisioMeasurementUnit.Millimeters);
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPageLayoutGridXml(roundTripPath, "Grid", "0", "20", "21", "8", "9", "MM");

            VisioDocument cleared = VisioDocument.Load(roundTripPath);
            cleared.Pages.Single(current => current.Name == "Grid").ClearLayoutGridPolicy();
            cleared.Save(clearedPath);

            Assert.Empty(VisioValidator.Validate(clearedPath));
            AssertPageLayoutGridXml(clearedPath, "Grid", null, null, null, null, null);
        }

        [Fact]
        public void LayoutGridSizingLoadsEachCellWithItsDeclaredUnit() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Mixed Grid", 11, 8.5);
            page.EnableLayoutGrid = true;
            page.SetLayoutGridSizing(1, 1, 1, 1);
            document.Save();

            RewritePageSheetCells(filePath, "Mixed Grid", (pageSheet, ns) => {
                SetPageSheetCell(pageSheet, ns, "BlockSizeX", "2", "IN");
                SetPageSheetCell(pageSheet, ns, "BlockSizeY", "5.08", "CM");
                SetPageSheetCell(pageSheet, ns, "AvenueSizeX", "50.8", "MM");
                SetPageSheetCell(pageSheet, ns, "AvenueSizeY", "2", "IN");
            });

            VisioPage loadedPage = VisioDocument.Load(filePath).Pages.Single(current => current.Name == "Mixed Grid");
            Assert.Equal(2, loadedPage.LayoutBlockSizeX.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LayoutBlockSizeY.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LayoutAvenueSizeX.GetValueOrDefault(), 6);
            Assert.Equal(2, loadedPage.LayoutAvenueSizeY.GetValueOrDefault(), 6);
        }

        private static void AssertPageSettingsXml(
            string filePath,
            string pageName,
            string expectedLeftMargin,
            string expectedRightMargin,
            string expectedTopMargin,
            string expectedBottomMargin,
            string? expectedOrientation,
            string expectedLockReplace,
            string expectedLockDuplicate,
            string? unit = null) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertCellValue(pageSheet, ns, "PageLeftMargin", expectedLeftMargin, unit);
            AssertCellValue(pageSheet, ns, "PageRightMargin", expectedRightMargin, unit);
            AssertCellValue(pageSheet, ns, "PageTopMargin", expectedTopMargin, unit);
            AssertCellValue(pageSheet, ns, "PageBottomMargin", expectedBottomMargin, unit);
            AssertOptionalCellValue(pageSheet, ns, "PrintPageOrientation", expectedOrientation);
            AssertCellValue(pageSheet, ns, "PageLockReplace", expectedLockReplace, "BOOL");
            AssertCellValue(pageSheet, ns, "PageLockDuplicate", expectedLockDuplicate, "BOOL");
        }

        private static void AssertPageSetupXml(
            string filePath,
            string pageName,
            string expectedDrawingSizeType,
            string expectedDrawingResizeType,
            string expectedPageShapeSplit,
            string expectedUiVisibility) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertCellValue(pageSheet, ns, "DrawingSizeType", expectedDrawingSizeType);
            AssertCellValue(pageSheet, ns, "DrawingResizeType", expectedDrawingResizeType);
            AssertCellValue(pageSheet, ns, "PageShapeSplit", expectedPageShapeSplit);
            AssertCellValue(pageSheet, ns, "UIVisibility", expectedUiVisibility);
        }

        private static void AssertPageRoutingXml(
            string filePath,
            string pageName,
            string? expectedRouteStyle,
            string? expectedRouteAppearance,
            string? expectedLineJumpStyle,
            string? expectedLineJumpCode,
            string? expectedLineJumpDirX,
            string? expectedLineJumpDirY) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertOptionalCellValue(pageSheet, ns, "RouteStyle", expectedRouteStyle);
            AssertOptionalCellValue(pageSheet, ns, "LineRouteExt", expectedRouteAppearance);
            AssertOptionalCellValue(pageSheet, ns, "LineJumpStyle", expectedLineJumpStyle);
            AssertOptionalCellValue(pageSheet, ns, "LineJumpCode", expectedLineJumpCode);
            AssertOptionalCellValue(pageSheet, ns, "PageLineJumpDirX", expectedLineJumpDirX);
            AssertOptionalCellValue(pageSheet, ns, "PageLineJumpDirY", expectedLineJumpDirY);
        }

        private static void AssertPageRoutingSpacingXml(
            string filePath,
            string pageName,
            string? expectedLineToLineX,
            string? expectedLineToLineY,
            string? expectedLineToNodeX,
            string? expectedLineToNodeY,
            string? expectedUnit = null) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertOptionalCellValue(pageSheet, ns, "LineToLineX", expectedLineToLineX, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "LineToLineY", expectedLineToLineY, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "LineToNodeX", expectedLineToNodeX, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "LineToNodeY", expectedLineToNodeY, expectedUnit);
        }

        private static void AssertPagePlacementXml(
            string filePath,
            string pageName,
            string? expectedPlaceStyle,
            string? expectedPlaceDepth,
            string? expectedPlaceFlip,
            string? expectedPlowCode,
            string? expectedResizePage) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertOptionalCellValue(pageSheet, ns, "PlaceStyle", expectedPlaceStyle);
            AssertOptionalCellValue(pageSheet, ns, "PlaceDepth", expectedPlaceDepth);
            AssertOptionalCellValue(pageSheet, ns, "PlaceFlip", expectedPlaceFlip);
            AssertOptionalCellValue(pageSheet, ns, "PlowCode", expectedPlowCode);
            AssertOptionalCellValue(pageSheet, ns, "ResizePage", expectedResizePage);
        }

        private static void AssertPageLayoutGridXml(
            string filePath,
            string pageName,
            string? expectedEnableGrid,
            string? expectedBlockSizeX,
            string? expectedBlockSizeY,
            string? expectedAvenueSizeX,
            string? expectedAvenueSizeY,
            string? expectedUnit = null) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            XElement pageSheet = page.Element(ns + "PageSheet")!;

            AssertOptionalCellValue(pageSheet, ns, "EnableGrid", expectedEnableGrid, expectedEnableGrid == null ? null : "BOOL");
            AssertOptionalCellValue(pageSheet, ns, "BlockSizeX", expectedBlockSizeX, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "BlockSizeY", expectedBlockSizeY, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "AvenueSizeX", expectedAvenueSizeX, expectedUnit);
            AssertOptionalCellValue(pageSheet, ns, "AvenueSizeY", expectedAvenueSizeY, expectedUnit);
        }

        private static void AssertCellValue(XElement pageSheet, XNamespace ns, string name, string expectedValue, string? expectedUnit = null) {
            XElement cell = pageSheet.Elements(ns + "Cell").Single(current => (string?)current.Attribute("N") == name);
            Assert.Equal(expectedValue, cell.Attribute("V")!.Value);
            if (expectedUnit != null) {
                Assert.Equal(expectedUnit, cell.Attribute("U")?.Value);
            }
        }

        private static void AssertOptionalCellValue(XElement pageSheet, XNamespace ns, string name, string? expectedValue) {
            AssertOptionalCellValue(pageSheet, ns, name, expectedValue, null);
        }

        private static void AssertOptionalCellValue(XElement pageSheet, XNamespace ns, string name, string? expectedValue, string? expectedUnit) {
            XElement[] cells = pageSheet.Elements(ns + "Cell")
                .Where(current => (string?)current.Attribute("N") == name)
                .ToArray();
            if (expectedValue == null) {
                Assert.Empty(cells);
                return;
            }

            XElement cell = Assert.Single(cells);
            Assert.Equal(expectedValue, cell.Attribute("V")!.Value);
            if (expectedUnit != null) {
                Assert.Equal(expectedUnit, cell.Attribute("U")?.Value);
            }
        }

        private static void RewritePageSheetCells(string filePath, string pageName, Action<XElement, XNamespace> mutatePageSheet) {
            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry pagesEntry = archive.GetEntry("visio/pages/pages.xml") ?? throw new InvalidOperationException("Missing pages.xml");
            XDocument pages;
            using (Stream stream = pagesEntry.Open()) {
                pages = XDocument.Load(stream);
            }

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            mutatePageSheet(page.Element(ns + "PageSheet")!, ns);

            pagesEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/pages.xml");
            using Stream replacementStream = replacement.Open();
            pages.Save(replacementStream);
        }

        private static void SetPageSheetCell(XElement pageSheet, XNamespace ns, string name, string value, string unit) {
            XElement cell = pageSheet.Elements(ns + "Cell").Single(current => (string?)current.Attribute("N") == name);
            cell.SetAttributeValue("V", value);
            cell.SetAttributeValue("U", unit);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
