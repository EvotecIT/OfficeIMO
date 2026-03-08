using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load core and parse helpers for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {
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

            PackageRelationship? themeRel = documentPart.GetRelationshipsByType(ThemeRelationshipType).FirstOrDefault();
            if (themeRel != null) {
                Uri themeUri = PackUriHelper.ResolvePartUri(documentPart.Uri, themeRel.TargetUri);
                PackagePart themePart = package.GetPart(themeUri);
                XDocument themeDoc = XDocument.Load(themePart.GetStream());
                document.Theme = new VisioTheme {
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
                XDocument mastersDoc = XDocument.Load(mastersPart.GetStream());
                XNamespace ns = VisioNamespace;
                XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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
                    XDocument masterDoc = XDocument.Load(masterPart.GetStream());
                    XElement? masterShapeElement = masterDoc.Root?.Element(ns + "Shapes")?.Element(ns + "Shape");
                    VisioShape masterShape = masterShapeElement != null ? ParseShape(masterShapeElement, ns) : new VisioShape("1");
                    VisioMaster master = new(masterId, masterNameU, masterShape);
                    masters[masterId] = master;
                }
            }

            XDocument pagesDoc = XDocument.Load(pagesPart.GetStream());
            XNamespace vNs = VisioNamespace;
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            foreach (XElement pageRef in pagesDoc.Root?.Elements(vNs + "Page") ?? Enumerable.Empty<XElement>()) {
                string name = pageRef.Attribute("Name")?.Value ?? "Page";
                int pageId = int.TryParse(pageRef.Attribute("ID")?.Value, out int tmp) ? tmp : document.Pages.Count;
                VisioPage page = document.AddPage(name, id: pageId);
                page.NameU = pageRef.Attribute("NameU")?.Value ?? name;
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
                            case "InhibitSnap":
                                page.Snap = !TryParseTruthyCellValue(valueAttr);
                                break;
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
                XDocument pageDoc = XDocument.Load(pagePart.GetStream());

                XElement? shapesRoot = pageDoc.Root?.Element(vNs + "Shapes");
                Dictionary<string, VisioShape> shapeMap = new();
                List<XElement> connectorElements = new();
                foreach (XElement shapeElement in shapesRoot?.Elements(vNs + "Shape") ?? Enumerable.Empty<XElement>()) {
                    if (IsConnectorShape(shapeElement, masters)) {
                        connectorElements.Add(shapeElement);
                        continue;
                    }

                    VisioShape shape = ParseShape(shapeElement, vNs);
                    ApplyMasterReferences(shape, shapeElement, vNs, masters);

                    page.Shapes.Add(shape);
                    RegisterShapeHierarchy(shape, shapeMap);
                }

                Dictionary<string, (string? fromId, string? fromCell, string? toId, string? toCell)> connectionMap = new();
                foreach (XElement connectElement in pageDoc.Root?.Element(vNs + "Connects")?.Elements(vNs + "Connect") ?? Enumerable.Empty<XElement>()) {
                    string? connectorId = connectElement.Attribute("FromSheet")?.Value;
                    string? fromCell = connectElement.Attribute("FromCell")?.Value;
                    string? toSheet = connectElement.Attribute("ToSheet")?.Value;
                    string? toCell = connectElement.Attribute("ToCell")?.Value;
                    if (connectorId == null || fromCell == null || toSheet == null) {
                        continue;
                    }
                    var info = connectionMap.TryGetValue(connectorId, out var existing) ? existing : (null, null, null, null);
                    if (fromCell == "BeginX") {
                        info.fromId = toSheet;
                        info.fromCell = toCell;
                    } else if (fromCell == "EndX") {
                        info.toId = toSheet;
                        info.toCell = toCell;
                    }
                    connectionMap[connectorId] = info;
                }

                foreach (XElement connectorElement in connectorElements) {
                    string persistedId = connectorElement.Attribute("ID")?.Value ?? string.Empty;
                    if (!connectionMap.TryGetValue(persistedId, out var ids) || ids.fromId == null || ids.toId == null) {
                        continue;
                    }
                    string id = GetOriginalId(connectorElement, vNs) ?? persistedId;
                    string fromId = ids.fromId!;
                    string toId = ids.toId!;
                    if (!shapeMap.TryGetValue(fromId, out VisioShape? fromShape) || !shapeMap.TryGetValue(toId, out VisioShape? toShape)) {
                        continue;
                    }
                    VisioConnector connector = new VisioConnector(id, fromShape!, toShape!);

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
                        }
                    }

                    connector.Kind = DetermineConnectorKind(connectorElement, vNs, masters);
                    foreach (XElement geometrySection in connectorElement.Elements(vNs + "Section")
                                 .Where(section => string.Equals(section.Attribute("N")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))) {
                        connector.PreservedGeometrySections.Add(new XElement(geometrySection));
                    }

                    connector.FromConnectionPoint = ResolveConnectionPoint(fromShape, ids.fromCell);
                    connector.ToConnectionPoint = ResolveConnectionPoint(toShape, ids.toCell);
                    connector.Label = connectorElement.Element(vNs + "Text")?.Value;
                    page.Connectors.Add(connector);
                }
            }

            return document;
        }

        private const int MaxShapeNestingDepth = 100;
        private static readonly double DefaultLineWeight = VisioShape.DefaultLineWeight;

        private static double ParseDouble(string? value) {
            string? normalized = NormalizeCellLiteral(value);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0;
        }

        private static SixLabors.ImageSharp.Color ParseColor(string? value, SixLabors.ImageSharp.Color fallback) {
            string? normalized = NormalizeCellLiteral(value);
            return string.IsNullOrWhiteSpace(normalized) ? fallback : VisioHelpers.FromVisioColor(normalized!);
        }

        private static VisioShape ParseShape(XElement shapeElement, XNamespace ns, VisioShape? parent = null, int depth = 0) {
            if (depth > MaxShapeNestingDepth) {
                throw new InvalidOperationException("Maximum nesting depth exceeded");
            }

            VisioShape shape = ParseShapeBasics(shapeElement, ns);
            shape.Parent = parent;

            ParseShapeTransform(shape, shapeElement, ns);
            ParseShapeProperties(shape, shapeElement, ns);
            ParseChildShapes(shape, shapeElement, ns, depth);

            return shape;
        }

        private static VisioShape ParseShapeBasics(XElement shapeElement, XNamespace ns) {
            string persistedId = shapeElement.Attribute("ID")?.Value ?? string.Empty;
            string id = GetOriginalId(shapeElement, ns) ?? persistedId;
            VisioShape shape = new(id) {
                Name = shapeElement.Attribute("Name")?.Value,
                NameU = shapeElement.Attribute("NameU")?.Value,
                Text = shapeElement.Element(ns + "Text")?.Value
            };

            shape.PersistedId = persistedId;
            shape.Type = shapeElement.Attribute("Type")?.Value;
            shape.MasterShapeId = shapeElement.Attribute("MasterShape")?.Value;

            return shape;
        }

        private static void ParseShapeTransform(VisioShape shape, XElement shapeElement, XNamespace ns) {
            List<XElement> cellElements = shapeElement.Elements(ns + "Cell").ToList();
            bool pinXFound = false;
            bool pinYFound = false;
            bool widthFound = false;
            bool heightFound = false;
            bool locPinXFound = false;
            bool locPinYFound = false;
            bool angleFound = false;
            bool lineWeightFound = false;

            foreach (XElement cell in cellElements) {
                string? n = cell.Attribute("N")?.Value;
                string? v = cell.Attribute("V")?.Value;
                switch (n) {
                    case "PinX":
                        shape.PinX = ParseDouble(v);
                        pinXFound = true;
                        break;
                    case "PinY":
                        shape.PinY = ParseDouble(v);
                        pinYFound = true;
                        break;
                    case "Width":
                        shape.Width = ParseDouble(v);
                        widthFound = true;
                        shape.HasExplicitWidth = true;
                        break;
                    case "Height":
                        shape.Height = ParseDouble(v);
                        heightFound = true;
                        shape.HasExplicitHeight = true;
                        break;
                    case "LocPinX":
                        shape.LocPinX = ParseDouble(v);
                        locPinXFound = true;
                        shape.HasExplicitLocPinX = true;
                        break;
                    case "LocPinY":
                        shape.LocPinY = ParseDouble(v);
                        locPinYFound = true;
                        shape.HasExplicitLocPinY = true;
                        break;
                    case "Angle":
                        shape.Angle = ParseDouble(v);
                        angleFound = true;
                        break;
                    case "LineWeight":
                        shape.LineWeight = ParseDouble(v);
                        lineWeightFound = true;
                        break;
                    case "LinePattern":
                        if (TryParseCellIntValue(v, out int linePattern)) {
                            shape.LinePattern = linePattern;
                        }
                        break;
                    case "FillPattern":
                        if (TryParseCellIntValue(v, out int fillPattern)) {
                            shape.FillPattern = fillPattern;
                        }
                        break;
                    case "LineColor":
                        shape.LineColor = ParseColor(v, shape.LineColor);
                        break;
                    case "FillForegnd":
                        shape.FillColor = ParseColor(v, shape.FillColor);
                        break;
                }
            }

            XElement? xform = shapeElement.Element(ns + "XForm") ?? shapeElement.Element(ns + "XForm1D");
            if (xform != null) {
                if (!pinXFound) {
                    XElement? pinX = xform.Element(ns + "PinX");
                    if (pinX != null) {
                        shape.PinX = ParseDouble(pinX.Value);
                        pinXFound = true;
                    }
                }
                if (!pinYFound) {
                    XElement? pinY = xform.Element(ns + "PinY");
                    if (pinY != null) {
                        shape.PinY = ParseDouble(pinY.Value);
                        pinYFound = true;
                    }
                }
                if (!widthFound) {
                    XElement? width = xform.Element(ns + "Width");
                    if (width != null) {
                        shape.Width = ParseDouble(width.Value);
                        widthFound = true;
                        shape.HasExplicitWidth = true;
                    }
                }
                if (!heightFound) {
                    XElement? height = xform.Element(ns + "Height");
                    if (height != null) {
                        shape.Height = ParseDouble(height.Value);
                        heightFound = true;
                        shape.HasExplicitHeight = true;
                    }
                }
                if (!locPinXFound) {
                    XElement? locPinX = xform.Element(ns + "LocPinX");
                    if (locPinX != null) {
                        shape.LocPinX = ParseDouble(locPinX.Value);
                        locPinXFound = true;
                        shape.HasExplicitLocPinX = true;
                    }
                }
                if (!locPinYFound) {
                    XElement? locPinY = xform.Element(ns + "LocPinY");
                    if (locPinY != null) {
                        shape.LocPinY = ParseDouble(locPinY.Value);
                        locPinYFound = true;
                        shape.HasExplicitLocPinY = true;
                    }
                }
                if (!angleFound) {
                    XElement? angle = xform.Element(ns + "Angle");
                    if (angle != null) {
                        shape.Angle = ParseDouble(angle.Value);
                        angleFound = true;
                    }
                }
            }

            if (!locPinXFound) {
                shape.LocPinX = shape.Width / 2;
            }
            if (!locPinYFound) {
                shape.LocPinY = shape.Height / 2;
            }
            if (!angleFound) {
                shape.Angle = 0;
            }
            if (!lineWeightFound) {
                shape.LineWeight = DefaultLineWeight;
            }
        }

        private static void ParseShapeProperties(VisioShape shape, XElement shapeElement, XNamespace ns) {
            List<XElement> sectionElements = shapeElement.Elements(ns + "Section").ToList();

            XElement? connectionSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "Connection");
            if (connectionSection != null) {
                foreach (XElement row in connectionSection.Elements(ns + "Row")) {
                    double x = 0;
                    double y = 0;
                    double dirX = 0;
                    double dirY = 0;
                    int? sectionIndex = null;
                    if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedSectionIndex) &&
                        parsedSectionIndex >= 0) {
                        sectionIndex = parsedSectionIndex;
                    }

                    foreach (XElement cell in row.Elements(ns + "Cell")) {
                        string? n = cell.Attribute("N")?.Value;
                        string? v = cell.Attribute("V")?.Value;
                        switch (n) {
                            case "X":
                                x = ParseDouble(v);
                                break;
                            case "Y":
                                y = ParseDouble(v);
                                break;
                            case "DirX":
                                dirX = ParseDouble(v);
                                break;
                            case "DirY":
                                dirY = ParseDouble(v);
                                break;
                        }
                    }
                    shape.ConnectionPoints.Add(new VisioConnectionPoint(x, y, dirX, dirY) {
                        SectionIndex = sectionIndex
                    });
                }
            }

            XElement? propSection = sectionElements.FirstOrDefault(e => e.Attribute("N")?.Value == "Prop");
            if (propSection != null) {
                foreach (XElement row in propSection.Elements(ns + "Row")) {
                    string? key = row.Attribute("N")?.Value;
                    XElement? valueCell = row.Elements(ns + "Cell").FirstOrDefault(c => c.Attribute("N")?.Value == "Value");
                    string? value = valueCell?.Attribute("V")?.Value;
                    if (!string.IsNullOrEmpty(key) && value != null && !string.Equals(key, OriginalIdPropName, StringComparison.Ordinal)) {
                        string keyNonNull = key!;
                        shape.Data[keyNonNull] = value;
                    }
                }
            }
        }

        private static string? GetOriginalId(XElement shapeElement, XNamespace ns) {
            XElement? propSection = shapeElement.Elements(ns + "Section")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "Prop");
            if (propSection == null) {
                return null;
            }

            XElement? originalIdRow = propSection.Elements(ns + "Row")
                .FirstOrDefault(row => string.Equals(row.Attribute("N")?.Value, OriginalIdPropName, StringComparison.Ordinal));
            return originalIdRow?.Elements(ns + "Cell")
                .FirstOrDefault(c => c.Attribute("N")?.Value == "Value")
                ?.Attribute("V")?.Value;
        }

        private static void ParseChildShapes(VisioShape shape, XElement shapeElement, XNamespace ns, int depth) {
            XElement? childShapes = shapeElement.Element(ns + "Shapes");
            if (childShapes == null) {
                return;
            }

            List<XElement> childElements = childShapes.Elements(ns + "Shape").ToList();
            foreach (XElement childElement in childElements) {
                VisioShape childShape = ParseShape(childElement, ns, shape, depth + 1);
                shape.Children.Add(childShape);
            }
        }

        private static void ApplyMasterReferences(VisioShape shape, XElement shapeElement, XNamespace ns, Dictionary<string, VisioMaster> masters, VisioMaster? inheritedMaster = null, VisioShape? inheritedMasterShape = null) {
            VisioMaster? effectiveMaster = inheritedMaster;
            VisioShape? effectiveMasterShape = inheritedMasterShape;

            string? masterIdAttr = shapeElement.Attribute("Master")?.Value;
            if (!string.IsNullOrEmpty(masterIdAttr) && masters.TryGetValue(masterIdAttr!, out VisioMaster? resolvedMaster)) {
                effectiveMaster = resolvedMaster;
                effectiveMasterShape = resolvedMaster.Shape;
            }

            if (effectiveMaster != null) {
                shape.Master = effectiveMaster;
            }

            if (!string.IsNullOrEmpty(shape.MasterShapeId) && effectiveMaster != null) {
                VisioShape? referencedMasterShape = effectiveMaster.Shape.FindDescendantById(shape.MasterShapeId!);
                if (referencedMasterShape != null) {
                    shape.MasterShape = referencedMasterShape;
                    effectiveMasterShape = referencedMasterShape;
                }
            }

            VisioShape? fallbackMasterShape = shape.MasterShape ?? effectiveMasterShape ?? effectiveMaster?.Shape;
            if (fallbackMasterShape != null) {
                if (!shape.HasExplicitWidth) {
                    shape.Width = fallbackMasterShape.Width;
                }
                if (!shape.HasExplicitHeight) {
                    shape.Height = fallbackMasterShape.Height;
                }
                if (!shape.HasExplicitLocPinX) {
                    shape.LocPinX = fallbackMasterShape.LocPinX;
                }
                if (!shape.HasExplicitLocPinY) {
                    shape.LocPinY = fallbackMasterShape.LocPinY;
                }
            }

            XElement? childShapes = shapeElement.Element(ns + "Shapes");
            if (childShapes != null && shape.Children.Count > 0) {
                List<XElement> childElements = childShapes.Elements(ns + "Shape").ToList();
                int count = Math.Min(childElements.Count, shape.Children.Count);
                for (int i = 0; i < count; i++) {
                    VisioShape? inheritedChildMasterShape = null;
                    if (fallbackMasterShape != null && i < fallbackMasterShape.Children.Count) {
                        inheritedChildMasterShape = fallbackMasterShape.Children[i];
                    }

                    ApplyMasterReferences(shape.Children[i], childElements[i], ns, masters, effectiveMaster, inheritedChildMasterShape ?? fallbackMasterShape);
                }
            }
        }

        private static void RegisterShapeHierarchy(VisioShape shape, Dictionary<string, VisioShape> shapeMap) {
            shapeMap[shape.Id] = shape;
            if (!string.IsNullOrEmpty(shape.PersistedId)) {
                shapeMap[shape.PersistedId!] = shape;
            }
            foreach (VisioShape child in shape.Children) {
                RegisterShapeHierarchy(child, shapeMap);
            }
        }

        private static bool IsConnectorShape(XElement shapeElement, IReadOnlyDictionary<string, VisioMaster> masters) {
            string? nameU = shapeElement.Attribute("NameU")?.Value;
            if (string.Equals(nameU, "Connector", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(nameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (TryGetTruthyCellValue(shapeElement, "OneD")) {
                return true;
            }

            string? masterId = shapeElement.Attribute("Master")?.Value;
            if (!string.IsNullOrEmpty(masterId) &&
                masters.TryGetValue(masterId!, out VisioMaster? master) &&
                string.Equals(master.NameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
            
            return false;
        }

        private static ConnectorKind DetermineConnectorKind(XElement connectorElement, XNamespace ns, IReadOnlyDictionary<string, VisioMaster> masters) {
            if (HasDynamicConnectorIdentity(connectorElement, masters)) {
                return ConnectorKind.Dynamic;
            }

            XElement? geometrySection = connectorElement.Elements(ns + "Section")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "Geometry");
            if (geometrySection == null) {
                return ConnectorKind.Dynamic;
            }

            List<XElement> rows = geometrySection.Elements(ns + "Row").ToList();
            List<XElement> drawableRows = rows
                .Where(row => !string.Equals(row.Attribute("T")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (drawableRows.Count == 0) {
                return ConnectorKind.Dynamic;
            }

            if (drawableRows.Any(IsCurvedGeometryRow)) {
                return ConnectorKind.Curved;
            }

            List<(double X, double Y)> points = new();
            foreach (XElement row in drawableRows) {
                string? type = row.Attribute("T")?.Value;
                if (!string.Equals(type, "MoveTo", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(type, "LineTo", StringComparison.OrdinalIgnoreCase)) {
                    return ConnectorKind.Curved;
                }

                points.Add((GetCellValue(row, ns, "X"), GetCellValue(row, ns, "Y")));
            }

            if (points.Count <= 2) {
                return ConnectorKind.Straight;
            }

            bool allOrthogonal = true;
            for (int i = 1; i < points.Count; i++) {
                (double previousX, double previousY) = points[i - 1];
                (double currentX, double currentY) = points[i];
                bool sameX = Math.Abs(previousX - currentX) <= 1e-9;
                bool sameY = Math.Abs(previousY - currentY) <= 1e-9;
                if (!sameX && !sameY) {
                    allOrthogonal = false;
                    break;
                }
            }

            return allOrthogonal ? ConnectorKind.RightAngle : ConnectorKind.Curved;
        }

        private static bool HasDynamicConnectorIdentity(XElement connectorElement, IReadOnlyDictionary<string, VisioMaster> masters) {
            string? nameU = connectorElement.Attribute("NameU")?.Value;
            if (string.Equals(nameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string? masterId = connectorElement.Attribute("Master")?.Value;
            return !string.IsNullOrEmpty(masterId) &&
                   masters.TryGetValue(masterId!, out VisioMaster? master) &&
                   string.Equals(master.NameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCurvedGeometryRow(XElement row) {
            string? type = row.Attribute("T")?.Value;
            if (string.IsNullOrEmpty(type)) {
                return false;
            }

            string rowType = type!;
            return rowType.IndexOf("Arc", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Spline", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Bezier", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("NURBS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rowType.IndexOf("Curve", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static double GetCellValue(XElement row, XNamespace ns, string cellName) {
            return ParseDouble(row.Elements(ns + "Cell")
                .FirstOrDefault(cell => string.Equals(cell.Attribute("N")?.Value, cellName, StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")?.Value);
        }

        private static bool TryGetTruthyCellValue(XElement element, string cellName) {
            string? value = element.Elements()
                .FirstOrDefault(child => string.Equals(child.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                                         string.Equals(child.Attribute("N")?.Value, cellName, StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")?.Value;
            return TryParseTruthyCellValue(value);
        }

        private static bool TryParseTruthyCellValue(string? value) {
            string? normalized = NormalizeCellLiteral(value);
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            if (bool.TryParse(normalized, out bool boolValue)) {
                return boolValue;
            }

            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue) &&
                   numericValue != 0;
        }

        private static bool TryParseCellIntValue(string? value, out int result) {
            string? normalized = NormalizeCellLiteral(value);
            if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out result)) {
                return true;
            }

            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue)) {
                int integerValue = Convert.ToInt32(numericValue);
                if (Math.Abs(numericValue - integerValue) <= 1e-9) {
                    result = integerValue;
                    return true;
                }
            }

            result = 0;
            return false;
        }

        private static string? NormalizeCellLiteral(string? value) {
            if (value is null) {
                return null;
            }

            string normalized = value.Trim();
            if (normalized.Length == 0) {
                return null;
            }
            while (normalized.StartsWith("GUARD(", StringComparison.OrdinalIgnoreCase) && normalized.EndsWith(")", StringComparison.Ordinal)) {
                normalized = normalized.Substring(6, normalized.Length - 7).Trim();
            }

            return normalized;
        }

        private static PackageRelationship GetRequiredSingleRelationship(IEnumerable<PackageRelationship> relationships, string description) {
            List<PackageRelationship> matches = relationships.ToList();
            if (matches.Count != 1) {
                throw new InvalidDataException($"Expected exactly one {description} relationship, but found {matches.Count}.");
            }

            return matches[0];
        }

        private static VisioConnectionPoint? ResolveConnectionPoint(VisioShape shape, string? connectionCell) {
            if (string.IsNullOrWhiteSpace(connectionCell) || string.Equals(connectionCell, "PinX", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            const string prefix = "Connections.X";
            string resolvedConnectionCell = connectionCell!;
            if (!resolvedConnectionCell.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string suffix = resolvedConnectionCell.Substring(prefix.Length);
            if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)) {
                return null;
            }

            int sectionIndex = index - 1;
            foreach (VisioConnectionPoint point in shape.ConnectionPoints) {
                if (point.SectionIndex == sectionIndex) {
                    return point;
                }
            }

            bool hasExplicitSectionIndices = shape.ConnectionPoints.Any(point => point.SectionIndex.HasValue);
            if (hasExplicitSectionIndices) {
                return null;
            }

            return sectionIndex >= 0 && sectionIndex < shape.ConnectionPoints.Count ? shape.ConnectionPoints[sectionIndex] : null;
        }
    }
}
