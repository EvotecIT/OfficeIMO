using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load core and parse helpers for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {
        private const int MaximumLoadedFontFamilyCharacters = 256;

        /// <summary>
        /// Loads an existing <c>.vsdx</c> file into a <see cref="VisioDocument"/>.
        /// </summary>
        /// <param name="filePath">Path to the <c>.vsdx</c> file.</param>
        private static VisioDocument LoadCore(string filePath) {
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return LoadCore(package, filePath);
        }

        private static VisioDocument LoadCore(Package package, string? filePath) {
            VisioDocument document = new() { _filePath = filePath };
            Dictionary<int, string> faceNamesById = new();

            document.Title = package.PackageProperties.Title;
            document.Author = package.PackageProperties.Creator;

            PackageRelationship documentRel = GetRequiredSingleRelationship(
                package.GetRelationshipsByType(DocumentRelationshipType),
                "package document");
            Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), documentRel.TargetUri);
            if (!package.PartExists(documentUri)) {
                throw new InvalidDataException($"Document relationship points to missing part '{documentUri}'.");
            }
            PackagePart documentPart = package.GetPart(documentUri);
            if (documentPart.ContentType != DocumentContentType) {
                throw new InvalidDataException($"Unexpected Visio document content type: {documentPart.ContentType}");
            }
            XDocument documentXml = LoadPackageXml(documentPart, "Visio document XML part");
            if (documentXml.Root != null) {
                foreach (XAttribute attribute in documentXml.Root.Attributes().Where(ShouldPreserveDocumentAttribute)) {
                    document.PreservedDocumentAttributes.Add(new XAttribute(attribute));
                }
                foreach (XElement element in documentXml.Root.Elements().Where(ShouldPreserveDocumentElement)) {
                    document.PreservedDocumentElements.Add(new XElement(element));
                }

                XElement? documentSettings = documentXml.Root.Element(XName.Get("DocumentSettings", VisioNamespace));
                if (documentSettings != null) {
                    foreach (XAttribute attribute in documentSettings.Attributes().Where(ShouldPreserveDocumentSettingsAttribute)) {
                        document.PreservedDocumentSettingsAttributes.Add(new XAttribute(attribute));
                    }
                    foreach (XElement element in documentSettings.Elements().Where(ShouldPreserveDocumentSettingsElement)) {
                        document.PreservedDocumentSettingsElements.Add(new XElement(element));
                    }

                    XElement? relayout = documentSettings.Element(XName.Get("RelayoutAndRerouteUponOpen", VisioNamespace));
                    if (relayout != null && !string.Equals(relayout.Value, "0", StringComparison.OrdinalIgnoreCase)) {
                        document._requestRecalcOnOpen = true;
                    }
                }

                XElement? colors = documentXml.Root.Element(XName.Get("Colors", VisioNamespace));
                if (colors != null) {
                    foreach (XAttribute attribute in colors.Attributes().Where(ShouldPreserveColorsAttribute)) {
                        document.PreservedColorsAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in colors.Elements().Where(ShouldPreserveColorsElement)) {
                        document.PreservedColorsElements.Add(new XElement(element));
                    }
                }

                XElement? faceNames = documentXml.Root.Element(XName.Get("FaceNames", VisioNamespace));
                if (faceNames != null) {
                    foreach (XAttribute attribute in faceNames.Attributes().Where(ShouldPreserveFaceNamesAttribute)) {
                        document.PreservedFaceNamesAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in faceNames.Elements().Where(ShouldPreserveFaceNamesElement)) {
                        document.PreservedFaceNamesElements.Add(new XElement(element));
                        if (string.Equals(element.Name.LocalName, "FaceName", StringComparison.OrdinalIgnoreCase) &&
                            int.TryParse(element.Attribute("ID")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int faceId)) {
                            string? name = element.Attribute("Name")?.Value;
                            if (!string.IsNullOrWhiteSpace(name) && !faceNamesById.ContainsKey(faceId)) {
                                string normalizedName = name!.Trim();
                                faceNamesById[faceId] = normalizedName.Length <= MaximumLoadedFontFamilyCharacters
                                    ? normalizedName
                                    : normalizedName.Substring(0, MaximumLoadedFontFamilyCharacters);
                            }
                        }
                    }
                }

                XElement? styleSheets = documentXml.Root.Element(XName.Get("StyleSheets", VisioNamespace));
                if (styleSheets != null) {
                    foreach (XAttribute attribute in styleSheets.Attributes().Where(ShouldPreserveStyleSheetsAttribute)) {
                        document.PreservedStyleSheetsAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in styleSheets.Elements().Where(ShouldPreserveStyleSheetsElement)) {
                        document.PreservedStyleSheetsElements.Add(new XElement(element));
                    }

                    foreach (XElement styleSheet in styleSheets.Elements(XName.Get("StyleSheet", VisioNamespace))) {
                        string id = styleSheet.Attribute("ID")?.Value ?? string.Empty;
                        if (!IsGeneratedStyleSheet(id)) {
                            document.PreservedAdditionalStyleSheets.Add(new XElement(styleSheet));
                            continue;
                        }

                        PreservedStyleSheetData preserved = GetOrCreatePreservedStyleSheet(document, id);
                        foreach (XAttribute attribute in styleSheet.Attributes().Where(attribute => ShouldPreserveStyleSheetAttribute(attribute, id))) {
                            preserved.Attributes.Add(new XAttribute(attribute));
                        }

                        foreach (XElement element in styleSheet.Elements().Where(element => ShouldPreserveStyleSheetElement(element, id))) {
                            preserved.ChildElements.Add(new XElement(element));
                        }
                    }
                }
            }

            PackageRelationship? themeRel = documentPart.GetRelationshipsByType(ThemeRelationshipType).FirstOrDefault();
            if (themeRel != null) {
                Uri themeUri = PackUriHelper.ResolvePartUri(documentPart.Uri, themeRel.TargetUri);
                PackagePart themePart = package.GetPart(themeUri);
                XDocument themeDoc = LoadPackageXml(themePart, "Visio theme XML part");
                document.PackageTheme = new VisioPackageTheme {
                    Name = themeDoc.Root?.Attribute("name")?.Value,
                    TemplateXml = new XDocument(themeDoc)
                };
            }

            PackageRelationship pagesRel = GetRequiredSingleRelationship(
                documentPart.GetRelationshipsByType(PagesRelationshipType),
                "document pages");
            Uri pagesUri = PackUriHelper.ResolvePartUri(documentPart.Uri, pagesRel.TargetUri);
            if (!package.PartExists(pagesUri)) {
                throw new InvalidDataException($"Pages relationship points to missing part '{pagesUri}'.");
            }
            PackagePart pagesPart = package.GetPart(pagesUri);

            // Load masters (if exist) to populate references on shapes
            Dictionary<string, VisioMaster> masters = new();
            if (documentPart.GetRelationshipsByType(MastersRelationshipType).FirstOrDefault() is PackageRelationship mastersRel) {
                Uri mastersUri = PackUriHelper.ResolvePartUri(documentPart.Uri, mastersRel.TargetUri);
                PackagePart mastersPart = package.GetPart(mastersUri);
                XDocument mastersDoc = LoadPackageXml(mastersPart, "Visio masters XML part");
                XNamespace ns = VisioNamespace;
                XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                List<XAttribute> preservedMastersRootAttributes = mastersDoc.Root?
                    .Attributes()
                    .Where(ShouldPreserveMastersRootAttribute)
                    .Select(attribute => new XAttribute(attribute))
                    .ToList() ?? new List<XAttribute>();
                List<XElement> preservedMastersRootElements = mastersDoc.Root?
                    .Elements()
                    .Where(ShouldPreserveMastersRootElement)
                    .Select(element => new XElement(element))
                    .ToList() ?? new List<XElement>();

                foreach (XElement masterElement in mastersDoc.Root?.Elements(ns + "Master") ?? Enumerable.Empty<XElement>()) {
                    string masterId = masterElement.Attribute("ID")?.Value ?? string.Empty;
                    string masterNameU = masterElement.Attribute("NameU")?.Value ?? string.Empty;
                    string? mRelIdValue = masterElement.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value;
                    if (string.IsNullOrEmpty(mRelIdValue)) {
                        continue;
                    }
                    string mRelId = mRelIdValue!;

                    PackageRelationship rel = mastersPart.GetRelationship(mRelId);
                    Uri masterUri = PackUriHelper.ResolvePartUri(mastersPart.Uri, rel.TargetUri);
                    PackagePart masterPart = package.GetPart(masterUri);
                    XDocument masterDoc = LoadPackageXml(masterPart, "Visio master XML part");
                    XElement? masterShapesElement = masterDoc.Root?.Element(ns + "Shapes");
                    XElement? masterShapeElement = masterShapesElement?.Elements(ns + "Shape").FirstOrDefault();
                    VisioShape masterShape = masterShapeElement != null ? ParseShapeCore(masterShapeElement, ns, faceNamesById) : new VisioShape("1");
                    VisioMaster master = new(masterId, masterNameU, masterShape);
                    foreach (XAttribute attribute in masterElement.Attributes().Where(ShouldPreserveMasterAttribute)) {
                        master.PreservedMasterAttributes.Add(new XAttribute(attribute));
                    }
                    XElement? masterPageSheet = masterElement.Element(ns + "PageSheet");
                    if (masterPageSheet != null) {
                        foreach (XAttribute attribute in masterPageSheet.Attributes().Where(ShouldPreserveMasterPageSheetAttribute)) {
                            master.PreservedPageSheetAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement cell in masterPageSheet.Elements(ns + "Cell").Where(ShouldPreserveMasterPageSheetCell)) {
                            master.PreservedPageSheetCells.Add(new XElement(cell));
                        }
                        foreach (XElement section in masterPageSheet.Elements(ns + "Section")) {
                            if (string.Equals(section.Attribute("N")?.Value, "User", StringComparison.OrdinalIgnoreCase)) {
                                List<VisioUserCell> userCells = new();
                                ParseUserCells(section, ns, userCells);
                                master.IsPackageBacked = userCells.Any(IsPackageBackedMasterUserCell);
                                VisioStencilMetadata.Apply(master, userCells);
                            }
                            master.PreservedPageSheetSections.Add(new XElement(section));
                        }
                    }
                    foreach (XElement element in masterElement.Elements().Where(ShouldPreserveMasterElement)) {
                        master.PreservedMasterElements.Add(new XElement(element));
                    }
                    if (masterShapesElement != null) {
                        foreach (XAttribute attribute in masterShapesElement.Attributes().Where(ShouldPreserveMasterShapesAttribute)) {
                            master.PreservedShapesAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement additionalShape in masterShapesElement.Elements(ns + "Shape").Skip(1)) {
                            master.PreservedAdditionalShapeElements.Add(new XElement(additionalShape));
                        }
                    }
                    if (masterDoc.Root != null) {
                        foreach (XAttribute attribute in masterDoc.Root.Attributes().Where(ShouldPreserveMasterContentAttribute)) {
                            master.PreservedMasterContentAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement element in masterDoc.Root.Elements().Where(element => element.Name != ns + "Shapes")) {
                            master.PreservedMasterContentElements.Add(new XElement(element));
                        }
                    }
                    foreach (XAttribute attribute in preservedMastersRootAttributes) {
                        master.PreservedMastersRootAttributes.Add(new XAttribute(attribute));
                    }
                    foreach (XElement element in preservedMastersRootElements) {
                        master.PreservedMastersRootElements.Add(new XElement(element));
                    }
                    masters[masterId] = master;
                    document.RegisterMaster(master);
                }
            }

            XDocument pagesDoc = LoadPackageXml(pagesPart, "Visio pages XML part");
            XNamespace vNs = VisioNamespace;
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            foreach (XElement pageRef in pagesDoc.Root?.Elements(vNs + "Page") ?? Enumerable.Empty<XElement>()) {
                string name = pageRef.Attribute("Name")?.Value ?? "Page";
                int pageId = int.TryParse(pageRef.Attribute("ID")?.Value, out int tmp) ? tmp : document.Pages.Count;
                VisioPage page = document.AddPage(name, id: pageId);
                page.NameU = pageRef.Attribute("NameU")?.Value ?? name;
                page.IsBackground = TryParseTruthyCellValue(pageRef.Attribute("Background")?.Value);
                if (int.TryParse(pageRef.Attribute("BackPage")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int backPageId)) {
                    page.SetLoadedBackgroundPageId(backPageId);
                }

                foreach (XAttribute attribute in pageRef.Attributes().Where(ShouldPreservePageAttribute)) {
                    page.PreservedPageAttributes.Add(new XAttribute(attribute));
                }
                string? viewScaleValue = pageRef.Attribute("ViewScale")?.Value;
                double parsedViewScale = ParseDouble(viewScaleValue);
                if (double.IsNaN(parsedViewScale) || double.IsInfinity(parsedViewScale) || parsedViewScale <= 0) {
                    page.ViewScale = 1;
                } else {
                    page.ViewScale = parsedViewScale;
                }
                double viewCenterX = ParseDouble(pageRef.Attribute("ViewCenterX")?.Value);
                double viewCenterY = ParseDouble(pageRef.Attribute("ViewCenterY")?.Value);

                XElement? pageSheet = pageRef.Element(vNs + "PageSheet");
                VisioMeasurementUnit? detectedScaleUnit = null;
                VisioScaleSetting? pendingDrawingScale = null;
                bool pageScaleApplied = false;
                double? pageWidthInches = null;
                double? pageHeightInches = null;
                double? leftMargin = null;
                double? rightMargin = null;
                double? topMargin = null;
                double? bottomMargin = null;
                VisioMeasurementUnit marginUnit = VisioMeasurementUnit.Inches;
                double? lineToLineX = null;
                double? lineToLineY = null;
                double? lineToNodeX = null;
                double? lineToNodeY = null;
                VisioMeasurementUnit? connectorSpacingUnit = null;
                double? blockSizeX = null;
                double? blockSizeY = null;
                double? avenueSizeX = null;
                double? avenueSizeY = null;
                VisioMeasurementUnit? layoutGridUnit = null;
                if (pageSheet != null) {
                    foreach (XElement cell in pageSheet.Elements(vNs + "Cell")) {
                        string? cellName = cell.Attribute("N")?.Value;
                        string? valueAttr = cell.Attribute("V")?.Value;
                        string? unitAttr = cell.Attribute("U")?.Value;
                        switch (cellName) {
                            case "PageWidth":
                                if (!detectedScaleUnit.HasValue) {
                                    detectedScaleUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, page.ScaleMeasurementUnit);
                                }
                                double parsedPageWidth = ParseDouble(valueAttr);
                                if (!double.IsNaN(parsedPageWidth) && !double.IsInfinity(parsedPageWidth) && parsedPageWidth > 0) {
                                    VisioMeasurementUnit pageWidthUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, detectedScaleUnit ?? page.ScaleMeasurementUnit);
                                    pageWidthInches = parsedPageWidth.ToInches(pageWidthUnit);
                                }
                                break;
                            case "PageHeight":
                                if (!detectedScaleUnit.HasValue) {
                                    detectedScaleUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, page.ScaleMeasurementUnit);
                                }
                                double parsedPageHeight = ParseDouble(valueAttr);
                                if (!double.IsNaN(parsedPageHeight) && !double.IsInfinity(parsedPageHeight) && parsedPageHeight > 0) {
                                    VisioMeasurementUnit pageHeightUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, detectedScaleUnit ?? page.ScaleMeasurementUnit);
                                    pageHeightInches = parsedPageHeight.ToInches(pageHeightUnit);
                                }
                                break;
                            case "PageScale":
                                double pageScaleInches = ParseDouble(valueAttr);
                                if (!double.IsNaN(pageScaleInches) && !double.IsInfinity(pageScaleInches) && pageScaleInches > 0) {
                                    VisioMeasurementUnit fallbackUnit = detectedScaleUnit ?? page.ScaleMeasurementUnit;
                                    VisioScaleSetting loadedPageScale = VisioScaleSetting.FromInches(pageScaleInches, unitAttr, fallbackUnit);
                                    if (detectedScaleUnit.HasValue) {
                                        page.DefaultUnit = detectedScaleUnit.Value;
                                    }
                                    page.ApplyLoadedPageScale(loadedPageScale);
                                    pageScaleApplied = true;
                                }
                                break;
                            case "DrawingScale":
                                double drawingScaleInches = ParseDouble(valueAttr);
                                if (!double.IsNaN(drawingScaleInches) && !double.IsInfinity(drawingScaleInches) && drawingScaleInches > 0) {
                                    VisioMeasurementUnit drawingFallback = detectedScaleUnit ?? page.ScaleMeasurementUnit;
                                    VisioScaleSetting loadedDrawingScale = VisioScaleSetting.FromInches(drawingScaleInches, unitAttr, drawingFallback);
                                    if (pageScaleApplied) {
                                        page.ApplyLoadedDrawingScale(loadedDrawingScale);
                                    } else {
                                        pendingDrawingScale = loadedDrawingScale;
                                    }
                                }
                                break;
                            case "DrawingSizeType":
                                if (TryParseCellIntValue(valueAttr, out int drawingSizeType) &&
                                    Enum.IsDefined(typeof(VisioDrawingSizeType), drawingSizeType)) {
                                    page.DrawingSizeType = (VisioDrawingSizeType)drawingSizeType;
                                }
                                break;
                            case "InhibitSnap":
                                page.Snap = !TryParseTruthyCellValue(valueAttr);
                                break;
                            case "PageLockReplace":
                                page.PageLockReplace = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "PageLockDuplicate":
                                page.PageLockDuplicate = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "UIVisibility":
                                if (TryParseCellIntValue(valueAttr, out int uiVisibility) &&
                                    Enum.IsDefined(typeof(VisioPageUiVisibility), uiVisibility)) {
                                    page.UiVisibility = (VisioPageUiVisibility)uiVisibility;
                                }
                                break;
                            case "DrawingResizeType":
                                page.AutoResizeDrawing = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "PageShapeSplit":
                                page.AllowShapeSplitting = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "PlaceStyle":
                                if (TryParseCellIntValue(valueAttr, out int placementStyle) &&
                                    Enum.IsDefined(typeof(VisioPlacementStyle), placementStyle)) {
                                    page.PlacementStyle = (VisioPlacementStyle)placementStyle;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "PlaceDepth":
                                if (TryParseCellIntValue(valueAttr, out int placementDepth) &&
                                    Enum.IsDefined(typeof(VisioPlacementDepth), placementDepth)) {
                                    page.PlacementDepth = (VisioPlacementDepth)placementDepth;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "PlaceFlip":
                                if (TryParseCellIntValue(valueAttr, out int placementFlip) &&
                                    IsValidPlacementFlip(placementFlip)) {
                                    page.PlacementFlip = (VisioPlacementFlip)placementFlip;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "PlowCode":
                                page.MoveShapesAwayOnDrop = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "ResizePage":
                                page.ResizePageToFitLayout = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "EnableGrid":
                                page.EnableLayoutGrid = TryParseTruthyCellValue(valueAttr);
                                break;
                            case "BlockSizeX":
                                VisioMeasurementUnit blockSizeXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= blockSizeXUnit;
                                blockSizeX = ParseNonNegativeDouble(valueAttr)?.ToInches(blockSizeXUnit);
                                break;
                            case "BlockSizeY":
                                VisioMeasurementUnit blockSizeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= blockSizeYUnit;
                                blockSizeY = ParseNonNegativeDouble(valueAttr)?.ToInches(blockSizeYUnit);
                                break;
                            case "AvenueSizeX":
                                VisioMeasurementUnit avenueSizeXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= avenueSizeXUnit;
                                avenueSizeX = ParseNonNegativeDouble(valueAttr)?.ToInches(avenueSizeXUnit);
                                break;
                            case "AvenueSizeY":
                                VisioMeasurementUnit avenueSizeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= avenueSizeYUnit;
                                avenueSizeY = ParseNonNegativeDouble(valueAttr)?.ToInches(avenueSizeYUnit);
                                break;
                            case "RouteStyle":
                                if (TryParseCellIntValue(valueAttr, out int routeStyle) &&
                                    Enum.IsDefined(typeof(VisioPageRouteStyle), routeStyle)) {
                                    page.ConnectorRouteStyle = (VisioPageRouteStyle)routeStyle;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "LineRouteExt":
                                if (TryParseCellIntValue(valueAttr, out int routeAppearance) &&
                                    Enum.IsDefined(typeof(VisioLineRouteExtension), routeAppearance)) {
                                    page.ConnectorRouteAppearance = (VisioLineRouteExtension)routeAppearance;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "LineJumpStyle":
                                if (TryParseCellIntValue(valueAttr, out int jumpStyle) &&
                                    Enum.IsDefined(typeof(VisioLineJumpStyle), jumpStyle)) {
                                    page.LineJumpStyle = (VisioLineJumpStyle)jumpStyle;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "LineJumpCode":
                                if (TryParseCellIntValue(valueAttr, out int jumpCode) &&
                                    Enum.IsDefined(typeof(VisioLineJumpCode), jumpCode)) {
                                    page.LineJumpCode = (VisioLineJumpCode)jumpCode;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "PageLineJumpDirX":
                                if (TryParseCellIntValue(valueAttr, out int jumpDirX) &&
                                    Enum.IsDefined(typeof(VisioHorizontalLineJumpDirection), jumpDirX)) {
                                    page.HorizontalLineJumpDirection = (VisioHorizontalLineJumpDirection)jumpDirX;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "PageLineJumpDirY":
                                if (TryParseCellIntValue(valueAttr, out int jumpDirY) &&
                                    Enum.IsDefined(typeof(VisioVerticalLineJumpDirection), jumpDirY)) {
                                    page.VerticalLineJumpDirection = (VisioVerticalLineJumpDirection)jumpDirY;
                                } else {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                            case "LineToLineX":
                                VisioMeasurementUnit lineToLineXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToLineXUnit;
                                lineToLineX = ParseNonNegativeDouble(valueAttr)?.ToInches(lineToLineXUnit);
                                break;
                            case "LineToLineY":
                                VisioMeasurementUnit lineToLineYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToLineYUnit;
                                lineToLineY = ParseNonNegativeDouble(valueAttr)?.ToInches(lineToLineYUnit);
                                break;
                            case "LineToNodeX":
                                VisioMeasurementUnit lineToNodeXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToNodeXUnit;
                                lineToNodeX = ParseNonNegativeDouble(valueAttr)?.ToInches(lineToNodeXUnit);
                                break;
                            case "LineToNodeY":
                                VisioMeasurementUnit lineToNodeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToNodeYUnit;
                                lineToNodeY = ParseNonNegativeDouble(valueAttr)?.ToInches(lineToNodeYUnit);
                                break;
                            case "PageLeftMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                leftMargin = ParseNonNegativeDouble(valueAttr)?.ToInches(marginUnit);
                                break;
                            case "PageRightMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                rightMargin = ParseNonNegativeDouble(valueAttr)?.ToInches(marginUnit);
                                break;
                            case "PageTopMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                topMargin = ParseNonNegativeDouble(valueAttr)?.ToInches(marginUnit);
                                break;
                            case "PageBottomMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                bottomMargin = ParseNonNegativeDouble(valueAttr)?.ToInches(marginUnit);
                                break;
                            case "PrintPageOrientation":
                                if (TryParseCellIntValue(valueAttr, out int orientation) &&
                                    Enum.IsDefined(typeof(VisioPagePrintOrientation), orientation)) {
                                    page.PrintOrientation = (VisioPagePrintOrientation)orientation;
                                }
                                break;
                            default:
                                if (ShouldPreservePageCell(cellName)) {
                                    page.PreservedPageSheetCells.Add(new XElement(cell));
                                }
                                break;
                        }
                    }

                    foreach (XElement section in pageSheet.Elements(vNs + "Section").Where(ShouldPreservePageSection)) {
                        if (string.Equals(section.Attribute("N")?.Value, "Layer", StringComparison.OrdinalIgnoreCase)) {
                            ParseLayerSection(page, section, vNs);
                        } else {
                            page.PreservedPageSheetSections.Add(new XElement(section));
                        }
                    }
                }

                if (!pageScaleApplied && detectedScaleUnit.HasValue) {
                    page.DefaultUnit = detectedScaleUnit.Value;
                    page.ScaleMeasurementUnit = detectedScaleUnit.Value;
                } else if (detectedScaleUnit.HasValue) {
                    page.DefaultUnit = detectedScaleUnit.Value;
                }

                if (pendingDrawingScale.HasValue) {
                    page.ApplyLoadedDrawingScale(pendingDrawingScale.Value);
                }

                if (pageWidthInches.HasValue && pageWidthInches.Value > 0) {
                    page.Width = pageWidthInches.Value;
                }
                if (pageHeightInches.HasValue && pageHeightInches.Value > 0) {
                    page.Height = pageHeightInches.Value;
                }
                page.SetLoadedMargins(leftMargin, rightMargin, topMargin, bottomMargin, marginUnit);
                page.SetLoadedConnectorSpacingInches(
                    lineToLineX,
                    lineToLineY,
                    lineToNodeX,
                    lineToNodeY,
                    connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                page.SetLoadedLayoutGridSizingInches(
                    blockSizeX,
                    blockSizeY,
                    avenueSizeX,
                    avenueSizeY,
                    layoutGridUnit ?? VisioMeasurementUnit.Inches);
                page.ViewCenterX = viewCenterX;
                page.ViewCenterY = viewCenterY;

                string? relIdValue = pageRef.Element(vNs + "Rel")?.Attribute(relNs + "id")?.Value;
                if (string.IsNullOrEmpty(relIdValue)) {
                    continue;
                }
                string relId = relIdValue!;

                PackageRelationship pageRel = pagesPart.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRel.TargetUri);
                PackagePart pagePart = package.GetPart(pageUri);
                XDocument pageDoc = LoadPackageXml(pagePart, "Visio page XML part");

                if (pageDoc.Root != null) {
                    foreach (XAttribute attribute in pageDoc.Root.Attributes().Where(ShouldPreservePageContentsAttribute)) {
                        page.PreservedPageContentAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in pageDoc.Root.Elements().Where(ShouldPreservePageContentsElement)) {
                        page.PreservedPageContentElements.Add(new XElement(element));
                    }
                }

                XElement? shapesRoot = pageDoc.Root?.Element(vNs + "Shapes");
                if (shapesRoot != null) {
                    foreach (XAttribute attribute in shapesRoot.Attributes().Where(ShouldPreserveShapesContainerAttribute)) {
                        page.PreservedShapesContainerAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in shapesRoot.Elements().Where(ShouldPreserveShapesContainerElement)) {
                        page.PreservedShapesContainerElements.Add(new XElement(element));
                    }
                }

                Dictionary<string, VisioShape> shapeMap = new();
                List<XElement> connectorElements = new();
                Dictionary<XElement, VisioShape> loadedShapesByElement = new();
                foreach (XElement shapeElement in shapesRoot?.Elements(vNs + "Shape") ?? Enumerable.Empty<XElement>()) {
                    if (IsConnectorShape(shapeElement, masters)) {
                        connectorElements.Add(shapeElement);
                        continue;
                    }

                    VisioShape shape = ParseShapeCore(shapeElement, vNs, faceNamesById);
                    ApplyMasterReferences(shape, shapeElement, vNs, masters);
                    ApplyLayerNamesFromIndexes(page, shape);

                    page.Shapes.Add(shape);
                    RegisterShapeHierarchy(shape, shapeMap);
                    loadedShapesByElement[shapeElement] = shape;
                }

                HydrateContainerRelationships(page, shapeMap);

                XElement? connectsRoot = pageDoc.Root?.Element(vNs + "Connects");
                if (connectsRoot != null) {
                    foreach (XAttribute attribute in connectsRoot.Attributes().Where(ShouldPreserveConnectsAttribute)) {
                        page.PreservedConnectsAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in connectsRoot.Elements().Where(ShouldPreserveConnectsElement)) {
                        page.PreservedConnectsElements.Add(new XElement(element));
                    }
                }

                List<XElement> orderedConnectElements = connectsRoot?.Elements(vNs + "Connect").ToList() ?? new List<XElement>();
                Dictionary<string, (string? fromId, string? fromCell, string? toId, string? toCell, List<XAttribute> beginAttributes, List<XName> beginOrder, List<XAttribute> endAttributes, List<XName> endOrder, XElement? beginElement, XElement? endElement)> connectionMap = new();
                foreach (XElement connectElement in orderedConnectElements) {
                    string? connectorId = connectElement.Attribute("FromSheet")?.Value;
                    string? fromCell = connectElement.Attribute("FromCell")?.Value;
                    string? toSheet = connectElement.Attribute("ToSheet")?.Value;
                    string? toCell = connectElement.Attribute("ToCell")?.Value;
                    if (connectorId == null || fromCell == null || toSheet == null) {
                        continue;
                    }
                    var info = connectionMap.TryGetValue(connectorId, out var existing)
                        ? existing
                        : (null, null, null, null, new List<XAttribute>(), new List<XName>(), new List<XAttribute>(), new List<XName>(), null, null);
                    if (fromCell == "BeginX") {
                        if (info.beginElement != null) {
                            continue;
                        }
                        info.fromId = toSheet;
                        info.fromCell = toCell;
                        CaptureConnectAttributes(connectElement, info.beginAttributes, info.beginOrder);
                        info.beginElement = connectElement;
                    } else if (fromCell == "EndX") {
                        if (info.endElement != null) {
                            continue;
                        }
                        info.toId = toSheet;
                        info.toCell = toCell;
                        CaptureConnectAttributes(connectElement, info.endAttributes, info.endOrder);
                        info.endElement = connectElement;
                    } else {
                        continue;
                    }
                    connectionMap[connectorId] = info;
                }

                Dictionary<string, VisioConnector> loadedConnectorsByPersistedId = new(StringComparer.Ordinal);
                Dictionary<XElement, VisioConnector> loadedConnectorsByElement = new();
                foreach (XElement connectorElement in connectorElements) {
                    string persistedId = connectorElement.Attribute("ID")?.Value ?? string.Empty;
                    if (!connectionMap.TryGetValue(persistedId, out var ids)) {
                        continue;
                    }
                    if (ids.fromId == null || ids.toId == null) {
                        continue;
                    }

                    string id = GetOriginalId(connectorElement, vNs) ?? persistedId;
                    string fromId = ids.fromId!;
                    string toId = ids.toId!;
                    if (!shapeMap.TryGetValue(fromId, out VisioShape? fromShape) || !shapeMap.TryGetValue(toId, out VisioShape? toShape)) {
                        continue;
                    }
                    VisioConnector connector = new VisioConnector(id, fromShape!, toShape!) {
                        PersistedId = persistedId
                    };

                    foreach (XElement cell in connectorElement.Elements(vNs + "Cell")) {
                        string? n = cell.Attribute("N")?.Value;
                        string? v = cell.Attribute("V")?.Value;
                        switch (n) {
                            case "BeginArrow":
                                if (TryParseCellIntValue(v, out int beginArrow)) {
                                    connector.BeginArrow = (EndArrow)beginArrow;
                                }
                                break;
                            case "EndArrow":
                                if (TryParseCellIntValue(v, out int endArrow)) {
                                    connector.EndArrow = (EndArrow)endArrow;
                                }
                                break;
                            case "LineWeight":
                                connector.LineWeight = ParseDouble(v);
                                break;
                            case "LinePattern":
                                if (TryParseCellIntValue(v, out int connectorLinePattern)) {
                                    connector.LinePattern = connectorLinePattern;
                                }
                                break;
                            case "LineColor":
                                connector.LineColor = ParseColor(v, connector.LineColor);
                                break;
                            case "LeftMargin":
                                EnsureConnectorTextStyle(connector).LeftMargin = ParseDouble(v);
                                break;
                            case "RightMargin":
                                EnsureConnectorTextStyle(connector).RightMargin = ParseDouble(v);
                                break;
                            case "TopMargin":
                                EnsureConnectorTextStyle(connector).TopMargin = ParseDouble(v);
                                break;
                            case "BottomMargin":
                                EnsureConnectorTextStyle(connector).BottomMargin = ParseDouble(v);
                                break;
                            case "VerticalAlign":
                                if (TryParseCellIntValue(v, out int connectorVerticalAlign) &&
                                    Enum.IsDefined(typeof(VisioTextVerticalAlignment), connectorVerticalAlign)) {
                                    EnsureConnectorTextStyle(connector).VerticalAlignment = (VisioTextVerticalAlignment)connectorVerticalAlign;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "TextBkgnd":
                                EnsureConnectorTextStyle(connector).BackgroundColor = ParseColor(v, default);
                                break;
                            case "TextBkgndTrans":
                                EnsureConnectorTextStyle(connector).BackgroundTransparency = ParseDouble(v);
                                break;
                            case "TxtPinX":
                                EnsureConnectorLabelPlacement(connector).AbsolutePinX = ParseDouble(v);
                                break;
                            case "TxtPinY":
                                EnsureConnectorLabelPlacement(connector).AbsolutePinY = ParseDouble(v);
                                break;
                            case "TxtWidth":
                                EnsureConnectorLabelPlacement(connector).Width = ParseDouble(v);
                                break;
                            case "TxtHeight":
                                EnsureConnectorLabelPlacement(connector).Height = ParseDouble(v);
                                break;
                            case "TxtLocPinX":
                                EnsureConnectorLabelPlacement(connector).LocPinX = ParseDouble(v);
                                break;
                            case "TxtLocPinY":
                                EnsureConnectorLabelPlacement(connector).LocPinY = ParseDouble(v);
                                break;
                            case "LayerMember":
                                ParseLayerIndexes(v, connector.LayerIndexes);
                                break;
                            case "ShapeRouteStyle":
                                if (TryParseCellIntValue(v, out int connectorRouteStyle) &&
                                    Enum.IsDefined(typeof(VisioPageRouteStyle), connectorRouteStyle)) {
                                    connector.RouteStyle = (VisioPageRouteStyle)connectorRouteStyle;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConLineRouteExt":
                                if (TryParseCellIntValue(v, out int connectorRouteAppearance) &&
                                    Enum.IsDefined(typeof(VisioLineRouteExtension), connectorRouteAppearance)) {
                                    connector.RouteAppearance = (VisioLineRouteExtension)connectorRouteAppearance;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConLineJumpStyle":
                                if (TryParseCellIntValue(v, out int connectorJumpStyle) &&
                                    Enum.IsDefined(typeof(VisioLineJumpStyle), connectorJumpStyle)) {
                                    connector.LineJumpStyle = (VisioLineJumpStyle)connectorJumpStyle;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConLineJumpCode":
                                if (TryParseCellIntValue(v, out int connectorJumpCode) &&
                                    Enum.IsDefined(typeof(VisioConnectorLineJumpCode), connectorJumpCode)) {
                                    connector.LineJumpCode = (VisioConnectorLineJumpCode)connectorJumpCode;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConLineJumpDirX":
                                if (TryParseCellIntValue(v, out int connectorJumpDirX) &&
                                    Enum.IsDefined(typeof(VisioHorizontalLineJumpDirection), connectorJumpDirX)) {
                                    connector.HorizontalJumpDirection = (VisioHorizontalLineJumpDirection)connectorJumpDirX;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConLineJumpDirY":
                                if (TryParseCellIntValue(v, out int connectorJumpDirY) &&
                                    Enum.IsDefined(typeof(VisioVerticalLineJumpDirection), connectorJumpDirY)) {
                                    connector.VerticalJumpDirection = (VisioVerticalLineJumpDirection)connectorJumpDirY;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            case "ConFixedCode":
                                if (TryParseCellIntValue(v, out int connectorRerouteBehavior) &&
                                    Enum.IsDefined(typeof(VisioConnectorRerouteBehavior), connectorRerouteBehavior)) {
                                    connector.RerouteBehavior = (VisioConnectorRerouteBehavior)connectorRerouteBehavior;
                                } else {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                            default:
                                if (VisioProtection.IsCellName(n) &&
                                    connector.Protection.TrySetCellValue(n, ParseNullableBoolCell(v))) {
                                    break;
                                }

                                if (ShouldPreserveConnectorCell(n)) {
                                    connector.PreservedCellElements.Add(new XElement(cell));
                                }
                                break;
                        }
                    }

                    connector.Kind = DetermineConnectorKind(connectorElement, vNs, masters);
                    ApplyLayerNamesFromIndexes(page, connector);
                    XElement? connectorCharSection = connectorElement.Elements(vNs + "Section")
                        .FirstOrDefault(section => IsCharacterSection(section.Attribute("N")?.Value));
                    if (connectorCharSection != null && TryParseSimpleConnectorCharSection(connector, connectorCharSection, vNs, faceNamesById)) {
                        connector.HasModeledCharSection = true;
                    }

                    XElement? connectorParaSection = connectorElement.Elements(vNs + "Section")
                        .FirstOrDefault(section => IsParagraphSection(section.Attribute("N")?.Value));
                    if (connectorParaSection != null && TryParseSimpleConnectorParaSection(connector, connectorParaSection, vNs)) {
                        connector.HasModeledParaSection = true;
                    }

                    foreach (XElement geometrySection in connectorElement.Elements(vNs + "Section")
                                 .Where(section => string.Equals(section.Attribute("N")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))) {
                        connector.PreservedGeometrySections.Add(new XElement(geometrySection));
                    }
                    foreach (XElement section in connectorElement.Elements(vNs + "Section")
                                 .Where(section => ShouldPreserveConnectorSection(connector, section))) {
                        connector.PreservedNonGeometrySections.Add(new XElement(section));
                    }
                    XElement? connectorHyperlinkSection = connectorElement.Elements(vNs + "Section")
                        .FirstOrDefault(section => string.Equals(section.Attribute("N")?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase));
                    if (connectorHyperlinkSection != null) {
                        ParseHyperlinks(connectorHyperlinkSection, vNs, connector.Hyperlinks);
                    }

                    XElement? connectorPropSection = connectorElement.Elements(vNs + "Section")
                        .FirstOrDefault(section => string.Equals(section.Attribute("N")?.Value, "Prop", StringComparison.OrdinalIgnoreCase));
                    if (connectorPropSection != null) {
                        ParseShapeDataRows(connectorPropSection, vNs, connector.ShapeData, connector.PreservedDataRows, connector.Data);
                    }

                    connector.FromConnectionPoint = ResolveConnectionPoint(fromShape, ids.fromCell);
                    connector.ToConnectionPoint = ResolveConnectionPoint(toShape, ids.toCell);
                    TryHydrateConnectorWaypoints(connector, connectorElement, vNs);
                    connector.PreservedFromConnectionCell = ids.fromCell;
                    connector.PreservedToConnectionCell = ids.toCell;
                    CopyPreservedAttributes(ids.beginAttributes, connector.PreservedBeginConnectAttributes);
                    CopyPreservedAttributeOrder(ids.beginOrder, connector.PreservedBeginConnectAttributeOrder);
                    CopyPreservedAttributes(ids.endAttributes, connector.PreservedEndConnectAttributes);
                    CopyPreservedAttributeOrder(ids.endOrder, connector.PreservedEndConnectAttributeOrder);
                    XElement? connectorTextElement = connectorElement.Element(vNs + "Text");
                    connector.Label = connectorTextElement?.Value;
                    connector.PreservedTextElement = connectorTextElement != null ? new XElement(connectorTextElement) : null;
                    connector.PreservedTextValue = connectorTextElement?.Value;
                    CaptureConnectorShapeChildOrder(connector, connectorElement);
                    page.Connectors.Add(connector);
                    loadedConnectorsByPersistedId[persistedId] = connector;
                    loadedConnectorsByElement[connectorElement] = connector;
                }

                foreach (XElement shapeChild in shapesRoot?.Elements() ?? Enumerable.Empty<XElement>()) {
                    if (!string.Equals(shapeChild.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase)) {
                        page.PreservedShapesChildren.Add(new VisioPage.PreservedShapeChildEntry(shapeChild));
                        continue;
                    }

                    if (loadedShapesByElement.TryGetValue(shapeChild, out VisioShape? loadedShape)) {
                        page.PreservedShapesChildren.Add(new VisioPage.PreservedShapeChildEntry(loadedShape));
                        continue;
                    }

                    if (loadedConnectorsByElement.TryGetValue(shapeChild, out VisioConnector? loadedConnector)) {
                        page.PreservedShapesChildren.Add(new VisioPage.PreservedShapeChildEntry(loadedConnector));
                        continue;
                    }

                    page.PreservedShapesChildren.Add(new VisioPage.PreservedShapeChildEntry(shapeChild));
                }

                foreach (XElement connectChild in connectsRoot?.Elements() ?? Enumerable.Empty<XElement>()) {
                    if (!string.Equals(connectChild.Name.LocalName, "Connect", StringComparison.OrdinalIgnoreCase)) {
                        page.PreservedConnectChildren.Add(new VisioPage.PreservedConnectChildEntry(connectChild));
                        continue;
                    }

                    string? connectorId = connectChild.Attribute("FromSheet")?.Value;
                    string? fromCell = connectChild.Attribute("FromCell")?.Value;
                    if (connectorId != null &&
                        fromCell != null &&
                        loadedConnectorsByPersistedId.TryGetValue(connectorId, out VisioConnector? connector) &&
                        connectionMap.TryGetValue(connectorId, out var ids)) {
                        if (string.Equals(fromCell, "BeginX", StringComparison.OrdinalIgnoreCase) &&
                            ReferenceEquals(ids.beginElement, connectChild)) {
                            page.PreservedConnectRows.Add(new VisioPage.PreservedConnectRowEntry(connector, VisioConnectorEndpointScope.Start));
                            page.PreservedConnectChildren.Add(new VisioPage.PreservedConnectChildEntry(connector, VisioConnectorEndpointScope.Start));
                            continue;
                        }

                        if (string.Equals(fromCell, "EndX", StringComparison.OrdinalIgnoreCase) &&
                            ReferenceEquals(ids.endElement, connectChild)) {
                            page.PreservedConnectRows.Add(new VisioPage.PreservedConnectRowEntry(connector, VisioConnectorEndpointScope.End));
                            page.PreservedConnectChildren.Add(new VisioPage.PreservedConnectChildEntry(connector, VisioConnectorEndpointScope.End));
                            continue;
                        }
                    }

                    page.PreservedConnectRows.Add(new VisioPage.PreservedConnectRowEntry(connectChild));
                    page.PreservedConnectChildren.Add(new VisioPage.PreservedConnectChildEntry(connectChild));
                }
            }

            LoadComments(package, documentPart, document);
            ResolvePageBackgrounds(document);
            return document;
        }

    }
}
