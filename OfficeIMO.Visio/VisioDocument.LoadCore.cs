using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Threading;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load core and parse helpers for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {
        private static VisioDocument LoadCore(Package package, string? filePath, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
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
            if (!VisioPackageFormat.TryFromContentType(documentPart.ContentType,
                    out VisioPackageType packageType)) {
                throw new InvalidDataException($"Unexpected Visio document content type: {documentPart.ContentType}");
            }
            document._packageType = packageType;
            LoadVbaProject(documentPart, document);
            XDocument documentXml = LoadPackageXml(documentPart, "Visio document XML part", cancellationToken);
            VisioFontWireCodec? fontWire = documentXml.Root == null ? null : VisioFontWireCodec.ReadDocument(documentXml.Root);
            if (documentXml.Root != null) fontWire!.DecodeCells(documentXml.Root, VisioNativeCellMetadata.Read(documentXml.Root.Elements()));
            var textBackgroundColors = ReadTextBackgroundPalette(documentXml.Root?.Element(XName.Get("Colors", VisioNamespace))?.Elements() ?? Enumerable.Empty<XElement>());
            var paintStyles = new VisioNativePaintStyleResolver(documentXml.Root);
            if (documentXml.Root is XElement documentRoot) {
                PreserveDocumentMetadata(document, documentRoot, faceNamesById, cancellationToken);
            }

            PackageRelationship? themeRel = documentPart.GetRelationshipsByType(ThemeRelationshipType).FirstOrDefault();
            if (themeRel != null) {
                Uri themeUri = PackUriHelper.ResolvePartUri(documentPart.Uri, themeRel.TargetUri);
                PackagePart themePart = package.GetPart(themeUri);
                XDocument themeDoc = LoadPackageXml(themePart, "Visio theme XML part", cancellationToken);
                document.PackageTheme = new VisioPackageTheme {
                    Name = themeDoc.Root?.Attribute("name")?.Value,
                    TemplateXml = new XDocument(themeDoc)
                };
            }

            PackageRelationship[] pageRelationships = documentPart
                .GetRelationshipsByType(PagesRelationshipType).ToArray();
            PackagePart? pagesPart = null;
            if (pageRelationships.Length == 0) {
                if (!document.IsStencil) {
                    throw new InvalidDataException("Required relationship 'document pages' is missing.");
                }
            } else {
                if (pageRelationships.Length != 1) {
                    throw new InvalidDataException("Relationship 'document pages' must occur exactly once.");
                }
                PackageRelationship pagesRel = pageRelationships[0];
                Uri pagesUri = PackUriHelper.ResolvePartUri(documentPart.Uri, pagesRel.TargetUri);
                if (!package.PartExists(pagesUri)) {
                    throw new InvalidDataException($"Pages relationship points to missing part '{pagesUri}'.");
                }
                pagesPart = package.GetPart(pagesUri);
            }

            var nativeCells = VisioNativeCellMetadata.Read(document.PreservedDocumentElements);
            fontWire?.BindModelFaces(document.PreservedFaceNamesElements);

            // Load masters (if exist) to populate references on shapes
            Dictionary<string, VisioMaster> masters = new();
            if (documentPart.GetRelationshipsByType(MastersRelationshipType).FirstOrDefault() is PackageRelationship mastersRel) {
                Uri mastersUri = PackUriHelper.ResolvePartUri(documentPart.Uri, mastersRel.TargetUri);
                PackagePart mastersPart = package.GetPart(mastersUri);
                XDocument mastersDoc = LoadPackageXml(mastersPart, "Visio masters XML part", cancellationToken);
                if (mastersDoc.Root != null) fontWire?.DecodeCells(mastersDoc.Root, nativeCells, "Masters");
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
                    cancellationToken.ThrowIfCancellationRequested();
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
                    XDocument masterDoc = LoadPackageXml(masterPart, "Visio master XML part", cancellationToken);
                    if (masterDoc.Root != null) fontWire?.DecodeCells(masterDoc.Root, nativeCells, VisioNativeCellMetadata.MasterScope(masterNameU));
                    HashSet<string> foreignRelationshipIds = ForeignRelationshipIds(masterDoc.Elements());
                    var foreignResources = new List<VisioForeignResource>();
                    document.CaptureForeignResources(masterPart, masterDoc, foreignResources);
                    XElement? masterShapesElement = masterDoc.Root?.Element(ns + "Shapes");
                    XElement? masterShapeElement = masterShapesElement?.Elements(ns + "Shape").FirstOrDefault();
                    VisioShape masterShape = masterShapeElement != null ? ParseShapeCore(masterShapeElement, ns, faceNamesById, textBackgroundColors: textBackgroundColors, paintStyles: paintStyles, cancellationToken: cancellationToken) : new VisioShape("1");
                    if (masterShapeElement != null) VisioNativeCellMetadata.BindTree(masterShape, masterShapeElement, VisioNativeCellMetadata.MasterScope(masterNameU), nativeCells);
                    BindForeignResources(masterShape, foreignResources);
                    VisioMaster master = new(masterId, masterNameU, masterShape);
                    master.ForeignResources.AddRange(foreignResources);
                    // Captured foreign relationships have bounded, owned bytes and remapped IDs.
                    // They do not prevent merging edits into the retained native ShapeSheet.
                    if (masterPart.GetRelationships().All(relationship => foreignRelationshipIds.Contains(relationship.Id))) {
                        master.RawMasterContentXml = new XDocument(masterDoc);
                        master.LoadedModelShapeXml = document.CreateMasterModelShapeXml(masterShape, masterDoc);
                    }
                    foreach (XAttribute attribute in masterElement.Attributes().Where(ShouldPreserveMasterAttribute)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        master.PreservedMasterAttributes.Add(new XAttribute(attribute));
                    }
                    XElement? masterPageSheet = masterElement.Element(ns + "PageSheet");
                    if (masterPageSheet != null) {
                        master.NativePageSheetMetadata = VisioNativeCellMetadata.Bind(masterPageSheet, "Masters", nativeCells);
                        master.LoadedPageSheetXml = new XElement(masterPageSheet);
                        foreach (XAttribute attribute in masterPageSheet.Attributes().Where(ShouldPreserveMasterPageSheetAttribute)) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedPageSheetAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement cell in masterPageSheet.Elements(ns + "Cell").Where(ShouldPreserveMasterPageSheetCell)) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedPageSheetCells.Add(new XElement(cell));
                        }
                        foreach (XElement section in masterPageSheet.Elements(ns + "Section")) {
                            cancellationToken.ThrowIfCancellationRequested();
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
                        cancellationToken.ThrowIfCancellationRequested();
                        master.PreservedMasterElements.Add(new XElement(element));
                    }
                    if (masterShapesElement != null) {
                        foreach (XAttribute attribute in masterShapesElement.Attributes().Where(ShouldPreserveMasterShapesAttribute)) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedShapesAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement additionalShape in masterShapesElement.Elements(ns + "Shape").Skip(1)) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedAdditionalShapeElements.Add(new XElement(additionalShape));
                            VisioNativeCellMetadata.BindRawShapeTree(master.NativeAdditionalShapeMetadata, additionalShape,
                                VisioNativeCellMetadata.MasterScope(masterNameU), nativeCells);
                        }
                    }
                    if (masterDoc.Root != null) {
                        foreach (XAttribute attribute in masterDoc.Root.Attributes().Where(ShouldPreserveMasterContentAttribute)) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedMasterContentAttributes.Add(new XAttribute(attribute));
                        }
                        foreach (XElement element in masterDoc.Root.Elements().Where(element => element.Name != ns + "Shapes")) {
                            cancellationToken.ThrowIfCancellationRequested();
                            master.PreservedMasterContentElements.Add(new XElement(element));
                        }
                    }
                    foreach (XAttribute attribute in preservedMastersRootAttributes) {
                        cancellationToken.ThrowIfCancellationRequested();
                        master.PreservedMastersRootAttributes.Add(new XAttribute(attribute));
                    }
                    foreach (XElement element in preservedMastersRootElements) {
                        cancellationToken.ThrowIfCancellationRequested();
                        master.PreservedMastersRootElements.Add(new XElement(element));
                    }
                    masters[masterId] = master;
                    document.CaptureMasterFontScope(master);
                    document.RegisterMaster(master);
                }
            }

            XDocument? pagesDoc = pagesPart == null ? null :
                LoadPackageXml(pagesPart, "Visio pages XML part", cancellationToken);
            if (pagesDoc?.Root != null) fontWire?.DecodeCells(pagesDoc.Root, nativeCells, "Pages");
            XNamespace vNs = VisioNamespace;
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            foreach (XElement pageRef in pagesDoc?.Root?.Elements(vNs + "Page") ?? Enumerable.Empty<XElement>()) {
                cancellationToken.ThrowIfCancellationRequested();
                string name = pageRef.Attribute("Name")?.Value ?? "Page";
                int pageId = int.TryParse(pageRef.Attribute("ID")?.Value, out int tmp) ? tmp : document.Pages.Count;
                VisioPage page = document.AddPage(name, id: pageId);
                page.NameU = pageRef.Attribute("NameU")?.Value ?? name;
                page.IsBackground = TryParseTruthyCellValue(pageRef.Attribute("Background")?.Value);
                if (int.TryParse(pageRef.Attribute("BackPage")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int backPageId)) {
                    page.SetLoadedBackgroundPageId(backPageId);
                }

                foreach (XAttribute attribute in pageRef.Attributes().Where(ShouldPreservePageAttribute)) {
                    cancellationToken.ThrowIfCancellationRequested();
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
                if (pageSheet != null) page.NativePageSheetMetadata = VisioNativeCellMetadata.Bind(pageSheet, "Pages", nativeCells);
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
                        cancellationToken.ThrowIfCancellationRequested();
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
                                    // ShapeSheet distance caches are internal inches; U is a display-unit hint.
                                    pageWidthInches = parsedPageWidth;
                                }
                                break;
                            case "PageHeight":
                                if (!detectedScaleUnit.HasValue) {
                                    detectedScaleUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, page.ScaleMeasurementUnit);
                                }
                                double parsedPageHeight = ParseDouble(valueAttr);
                                if (!double.IsNaN(parsedPageHeight) && !double.IsInfinity(parsedPageHeight) && parsedPageHeight > 0) {
                                    pageHeightInches = parsedPageHeight;
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
                                blockSizeX = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "BlockSizeY":
                                VisioMeasurementUnit blockSizeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= blockSizeYUnit;
                                blockSizeY = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "AvenueSizeX":
                                VisioMeasurementUnit avenueSizeXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= avenueSizeXUnit;
                                avenueSizeX = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "AvenueSizeY":
                                VisioMeasurementUnit avenueSizeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, layoutGridUnit ?? VisioMeasurementUnit.Inches);
                                layoutGridUnit ??= avenueSizeYUnit;
                                avenueSizeY = ParseNonNegativeDouble(valueAttr);
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
                                lineToLineX = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "LineToLineY":
                                VisioMeasurementUnit lineToLineYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToLineYUnit;
                                lineToLineY = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "LineToNodeX":
                                VisioMeasurementUnit lineToNodeXUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToNodeXUnit;
                                lineToNodeX = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "LineToNodeY":
                                VisioMeasurementUnit lineToNodeYUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, connectorSpacingUnit ?? VisioMeasurementUnit.Inches);
                                connectorSpacingUnit ??= lineToNodeYUnit;
                                lineToNodeY = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "PageLeftMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                leftMargin = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "PageRightMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                rightMargin = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "PageTopMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                topMargin = ParseNonNegativeDouble(valueAttr);
                                break;
                            case "PageBottomMargin":
                                marginUnit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitAttr, marginUnit);
                                bottomMargin = ParseNonNegativeDouble(valueAttr);
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
                        cancellationToken.ThrowIfCancellationRequested();
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
                if (pageSheet != null) {
                    page.PageSheetLengthCells = new VisioPageSheetLengthCells(pageSheet, document.CreatePageSheetModel(VisioNamespace, page));
                }
                page.ViewCenterX = viewCenterX;
                page.ViewCenterY = viewCenterY;

                string? relIdValue = pageRef.Element(vNs + "Rel")?.Attribute(relNs + "id")?.Value;
                if (string.IsNullOrEmpty(relIdValue)) {
                    continue;
                }
                string relId = relIdValue!;

                PackageRelationship pageRel = pagesPart!.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRel.TargetUri);
                PackagePart pagePart = package.GetPart(pageUri);
                XDocument pageDoc = LoadPackageXml(pagePart, "Visio page XML part", cancellationToken);
                if (pageDoc.Root != null) fontWire?.DecodeCells(pageDoc.Root, nativeCells, VisioNativeCellMetadata.PageScope(pageId));
                document.CaptureForeignResources(pagePart, pageDoc, page.ForeignResources);

                if (pageDoc.Root != null) {
                    foreach (XAttribute attribute in pageDoc.Root.Attributes().Where(ShouldPreservePageContentsAttribute)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedPageContentAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in pageDoc.Root.Elements().Where(ShouldPreservePageContentsElement)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedPageContentElements.Add(new XElement(element));
                    }
                }

                XElement? shapesRoot = pageDoc.Root?.Element(vNs + "Shapes");
                if (shapesRoot != null) {
                    foreach (XAttribute attribute in shapesRoot.Attributes().Where(ShouldPreserveShapesContainerAttribute)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedShapesContainerAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in shapesRoot.Elements().Where(ShouldPreserveShapesContainerElement)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedShapesContainerElements.Add(new XElement(element));
                    }
                }

                Dictionary<string, VisioShape> shapeMap = new();
                List<XElement> connectorElements = new();
                Dictionary<XElement, VisioShape> loadedShapesByElement = new();
                foreach (XElement shapeElement in shapesRoot?.Elements(vNs + "Shape") ?? Enumerable.Empty<XElement>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (IsConnectorShape(shapeElement, masters)) {
                        connectorElements.Add(shapeElement);
                        continue;
                    }

                    VisioShape shape = ParseShapeCore(shapeElement, vNs, faceNamesById, textBackgroundColors: textBackgroundColors, paintStyles: paintStyles, cancellationToken: cancellationToken);
                    VisioNativeCellMetadata.BindTree(shape, shapeElement, VisioNativeCellMetadata.PageScope(pageId), nativeCells);
                    BindForeignResources(shape, page.ForeignResources);
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
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedConnectsAttributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in connectsRoot.Elements().Where(ShouldPreserveConnectsElement)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        page.PreservedConnectsElements.Add(new XElement(element));
                    }
                }

                List<XElement> orderedConnectElements = connectsRoot?.Elements(vNs + "Connect").ToList() ?? new List<XElement>();
                Dictionary<string, (string? fromId, string? fromCell, string? toId, string? toCell, List<XAttribute> beginAttributes, List<XName> beginOrder, List<XAttribute> endAttributes, List<XName> endOrder, XElement? beginElement, XElement? endElement)> connectionMap = new();
                foreach (XElement connectElement in orderedConnectElements) {
                    cancellationToken.ThrowIfCancellationRequested();
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
                    cancellationToken.ThrowIfCancellationRequested();
                    string persistedId = connectorElement.Attribute("ID")?.Value ?? string.Empty;
                    bool hasConnections = connectionMap.TryGetValue(persistedId, out var ids);
                    VisioShape? fromShape = null, toShape = null;
                    // A missing attachment is a free endpoint; a broken attachment stays
                    // preserved native XML instead of silently discarding its target.
                    if (ids.fromId != null && !shapeMap.TryGetValue(ids.fromId, out fromShape) ||
                        ids.toId != null && !shapeMap.TryGetValue(ids.toId, out toShape)) continue;
                    VisioConnector? connector = LoadConnector(connectorElement, vNs, masters, faceNamesById, page, fromShape, toShape, ids.fromCell, ids.toCell, textBackgroundColors, paintStyles);
                    if (connector == null) continue;
                    connector.NativeCellMetadata = VisioNativeCellMetadata.Bind(connectorElement, VisioNativeCellMetadata.PageScope(pageId), nativeCells);
                    if (hasConnections) {
                        CopyPreservedAttributes(ids.beginAttributes, connector.PreservedBeginConnectAttributes);
                        CopyPreservedAttributeOrder(ids.beginOrder, connector.PreservedBeginConnectAttributeOrder);
                        CopyPreservedAttributes(ids.endAttributes, connector.PreservedEndConnectAttributes);
                        CopyPreservedAttributeOrder(ids.endOrder, connector.PreservedEndConnectAttributeOrder);
                    }
                    CaptureConnectorShapeChildOrder(connector, connectorElement);
                    var foreignIds = ForeignRelationshipIds(new[] { connectorElement });
                    connector.ForeignResources.AddRange(page.ForeignResources.Where(resource => foreignIds.Contains(resource.RelationshipId)));
                    page.Connectors.Add(connector);
                    loadedConnectorsByPersistedId[persistedId] = connector;
                    loadedConnectorsByElement[connectorElement] = connector;
                }

                foreach (XElement shapeChild in shapesRoot?.Elements() ?? Enumerable.Empty<XElement>()) {
                    cancellationToken.ThrowIfCancellationRequested();
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
                    cancellationToken.ThrowIfCancellationRequested();
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

            LoadComments(package, documentPart, document, cancellationToken);
            ResolvePageBackgrounds(document);
            cancellationToken.ThrowIfCancellationRequested();
            return document;
        }

    }
}
