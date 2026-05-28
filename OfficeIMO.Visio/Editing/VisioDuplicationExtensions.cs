using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for duplicating Visio content while keeping copied shapes independent.
    /// </summary>
    public static class VisioDuplicationExtensions {
        private const double DefaultDuplicateOffsetX = 0.35D;
        private const double DefaultDuplicateOffsetY = -0.35D;

        /// <summary>
        /// Duplicates a page in the same document, preserving page settings, layers, shapes, and internal connectors.
        /// </summary>
        /// <param name="document">Document that owns the source page and receives the duplicate.</param>
        /// <param name="sourcePage">Page to duplicate.</param>
        /// <param name="name">Optional name for the duplicate. When omitted, a unique copy name is generated.</param>
        /// <returns>The duplicated page.</returns>
        public static VisioPage DuplicatePage(this VisioDocument document, VisioPage sourcePage, string? name = null) {
            return DuplicatePage(document, sourcePage, new VisioPageDuplicationOptions { Name = name });
        }

        /// <summary>
        /// Duplicates a page in the same document, preserving page settings, layers, shapes, and internal connectors.
        /// </summary>
        /// <param name="document">Document that owns the source page and receives the duplicate.</param>
        /// <param name="sourcePage">Page to duplicate.</param>
        /// <param name="options">Optional duplication settings.</param>
        /// <returns>The duplicated page.</returns>
        public static VisioPage DuplicatePage(this VisioDocument document, VisioPage sourcePage, VisioPageDuplicationOptions? options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (sourcePage == null) {
                throw new ArgumentNullException(nameof(sourcePage));
            }

            if (!document.Pages.Contains(sourcePage)) {
                throw new InvalidOperationException("The source page must belong to the target document.");
            }

            VisioPageDuplicationOptions effectiveOptions = options ?? new VisioPageDuplicationOptions();
            VisioPage? duplicatedBackgroundPage = null;
            if (effectiveOptions.DuplicateBackgroundPage &&
                !sourcePage.IsBackground &&
                sourcePage.BackgroundPage != null) {
                duplicatedBackgroundPage = DuplicatePageCore(
                    document,
                    sourcePage.BackgroundPage,
                    effectiveOptions.BackgroundPageName,
                    backgroundPageOverride: null);
            }

            return DuplicatePageCore(document, sourcePage, effectiveOptions.Name, duplicatedBackgroundPage);
        }

        private static VisioPage DuplicatePageCore(VisioDocument document, VisioPage sourcePage, string? name, VisioPage? backgroundPageOverride) {
            string duplicateName = ResolveDuplicatePageName(document, sourcePage, name);
            VisioPage clone = document.AddPage(duplicateName, sourcePage.Width, sourcePage.Height);
            clone.NameU = duplicateName;
            CopyPageSettings(sourcePage, clone);
            CopyLayers(sourcePage, clone);
            CopyPagePreservation(sourcePage, clone);

            IdAllocator ids = new(clone, sourcePage);
            Dictionary<VisioShape, VisioShape> shapeMap = new();
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap = new();

            foreach (VisioShape shape in sourcePage.Shapes) {
                VisioShape shapeClone = CloneShape(shape, ids, 0D, 0D, applyOffset: false, shapeMap, connectionPointMap);
                clone.Shapes.Add(shapeClone);
            }

            RemapContainerMembership(shapeMap);

            foreach (VisioConnector connector in sourcePage.Connectors) {
                if (!shapeMap.TryGetValue(connector.From, out VisioShape? clonedFrom) ||
                    !shapeMap.TryGetValue(connector.To, out VisioShape? clonedTo)) {
                    continue;
                }

                VisioConnector connectorClone = CloneConnector(connector, ids, clonedFrom, clonedTo, 0D, 0D, connectionPointMap);
                clone.Connectors.Add(connectorClone);
            }

            if (sourcePage.IsBackground) {
                clone.IsBackground = true;
                if (backgroundPageOverride != null) {
                    clone.SetBackgroundPage(backgroundPageOverride);
                } else if (sourcePage.BackgroundPage != null) {
                    clone.SetBackgroundPage(sourcePage.BackgroundPage);
                }
            } else if (backgroundPageOverride != null) {
                clone.SetBackgroundPage(backgroundPageOverride);
            } else if (sourcePage.BackgroundPage != null) {
                clone.SetBackgroundPage(sourcePage.BackgroundPage);
            }

            return clone;
        }

        /// <summary>
        /// Duplicates this page in its owner document.
        /// </summary>
        /// <param name="page">Page to duplicate.</param>
        /// <param name="name">Optional name for the duplicate. When omitted, a unique copy name is generated.</param>
        /// <returns>The duplicated page.</returns>
        public static VisioPage Duplicate(this VisioPage page, string? name = null) {
            return Duplicate(page, new VisioPageDuplicationOptions { Name = name });
        }

        /// <summary>
        /// Duplicates this page in its owner document.
        /// </summary>
        /// <param name="page">Page to duplicate.</param>
        /// <param name="options">Optional duplication settings.</param>
        /// <returns>The duplicated page.</returns>
        public static VisioPage Duplicate(this VisioPage page, VisioPageDuplicationOptions? options) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (page.OwnerDocument == null) {
                throw new InvalidOperationException("The page is not associated with a document.");
            }

            return page.OwnerDocument.DuplicatePage(page, options);
        }

        /// <summary>
        /// Duplicates shapes on the same page and optionally copies connectors whose endpoints are both duplicated.
        /// </summary>
        /// <param name="page">Page that owns the shapes.</param>
        /// <param name="shapes">Shapes to duplicate. Nested children are copied with their selected ancestor.</param>
        /// <param name="offsetX">Horizontal offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="offsetY">Vertical offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="includeInternalConnectors">Whether connectors between duplicated shapes should also be copied.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection DuplicateShapes(this VisioPage page, IEnumerable<VisioShape> shapes, double offsetX = DefaultDuplicateOffsetX, double offsetY = DefaultDuplicateOffsetY, bool includeInternalConnectors = true) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<VisioShape> selectedShapes = shapes.Distinct().ToList();
            if (selectedShapes.Count == 0) {
                return new VisioShapeSelection(Array.Empty<VisioShape>(), page);
            }

            HashSet<VisioShape> pageShapes = new(page.AllShapes());
            foreach (VisioShape shape in selectedShapes) {
                if (!pageShapes.Contains(shape)) {
                    throw new InvalidOperationException("All duplicated shapes must belong to the target page.");
                }
            }

            HashSet<VisioShape> selectedSet = new(selectedShapes);
            List<VisioShape> rootShapes = selectedShapes
                .Where(shape => !HasSelectedAncestor(shape, selectedSet))
                .ToList();

            IdAllocator ids = new(page);
            Dictionary<VisioShape, VisioShape> shapeMap = new();
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap = new();
            List<VisioShape> duplicatedRoots = new();

            foreach (VisioShape root in rootShapes) {
                VisioShape clone = CloneShape(root, ids, offsetX, offsetY, applyOffset: true, shapeMap, connectionPointMap);
                page.Shapes.Add(clone);
                duplicatedRoots.Add(clone);
            }

            RemapContainerMembership(shapeMap);

            if (includeInternalConnectors) {
                foreach (VisioConnector connector in page.Connectors.ToList()) {
                    if (!shapeMap.TryGetValue(connector.From, out VisioShape? clonedFrom) ||
                        !shapeMap.TryGetValue(connector.To, out VisioShape? clonedTo)) {
                        continue;
                    }

                    VisioConnector clonedConnector = CloneConnector(connector, ids, clonedFrom, clonedTo, offsetX, offsetY, connectionPointMap);
                    page.Connectors.Add(clonedConnector);
                }
            }

            return new VisioShapeSelection(duplicatedRoots, page);
        }

        /// <summary>
        /// Duplicates a page-backed selection on the same page.
        /// </summary>
        /// <param name="selection">Selection to duplicate.</param>
        /// <param name="offsetX">Horizontal offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="offsetY">Vertical offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="includeInternalConnectors">Whether connectors between duplicated shapes should also be copied.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection Duplicate(this VisioShapeSelection selection, double offsetX = DefaultDuplicateOffsetX, double offsetY = DefaultDuplicateOffsetY, bool includeInternalConnectors = true) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Use page.DuplicateShapes(selection, ...) instead.");
            }

            return selection.OwnerPage.DuplicateShapes(selection, offsetX, offsetY, includeInternalConnectors);
        }

        private static bool HasSelectedAncestor(VisioShape shape, HashSet<VisioShape> selected) {
            VisioShape? parent = shape.Parent;
            while (parent != null) {
                if (selected.Contains(parent)) {
                    return true;
                }

                parent = parent.Parent;
            }

            return false;
        }

        private static VisioShape CloneShape(
            VisioShape source,
            IdAllocator ids,
            double offsetX,
            double offsetY,
            bool applyOffset,
            Dictionary<VisioShape, VisioShape> shapeMap,
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap) {
            VisioShape clone = new(ids.Next(), source.PinX + (applyOffset ? offsetX : 0D), source.PinY + (applyOffset ? offsetY : 0D), source.Width, source.Height, source.Text ?? string.Empty) {
                Name = source.Name,
                NameU = source.NameU,
                Type = source.Type,
                Master = source.Master,
                MasterShapeId = source.MasterShapeId,
                MasterShape = source.MasterShape,
                LineWeight = source.LineWeight,
                LocPinX = source.LocPinX,
                LocPinY = source.LocPinY,
                Angle = source.Angle,
                LineColor = source.LineColor,
                FillColor = source.FillColor,
                LinePattern = source.LinePattern,
                FillPattern = source.FillPattern,
                PlacementStyle = source.PlacementStyle,
                PlacementFlip = source.PlacementFlip,
                PlowCode = source.PlowCode,
                AllowPlacementOnTop = source.AllowPlacementOnTop,
                AllowHorizontalConnectorRoutingThrough = source.AllowHorizontalConnectorRoutingThrough,
                AllowVerticalConnectorRoutingThrough = source.AllowVerticalConnectorRoutingThrough,
                CanSplitShapes = source.CanSplitShapes,
                CanBeSplit = source.CanBeSplit,
                RelationshipsValue = source.RelationshipsValue,
                RelationshipsFormula = source.RelationshipsFormula,
                TextStyle = source.TextStyle?.Clone(),
                PreservedTextElement = source.PreservedTextElement == null ? null : new XElement(source.PreservedTextElement),
                PreservedTextValue = source.PreservedTextValue,
                HasModeledCharSection = source.HasModeledCharSection,
                HasModeledParaSection = source.HasModeledParaSection
            };

            CopyStringSet(source.LayerNames, clone.LayerNames);
            CopyStringList(source.ContainerMemberIds, clone.ContainerMemberIds);
            CopyStringList(source.ContainerOwnerIds, clone.ContainerOwnerIds);
            CopyDictionary(source.Data, clone.Data);
            CopyConnectionPoints(source, clone, connectionPointMap);
            CopyHyperlinks(source.Hyperlinks, clone.Hyperlinks);
            CopyUserCells(source.UserCells, clone.UserCells);
            CopyShapeData(source.ShapeData, clone.ShapeData, clone.Data);
            CopyProtection(source.Protection, clone.Protection);
            CopyElements(source.PreservedGeometrySections, clone.PreservedGeometrySections);
            CopyElements(source.PreservedCellElements, clone.PreservedCellElements);
            CopyElements(source.PreservedNonGeometrySections, clone.PreservedNonGeometrySections);
            CopyElements(source.PreservedDataRows, clone.PreservedDataRows);

            shapeMap[source] = clone;

            foreach (VisioShape child in source.Children) {
                clone.Children.Add(CloneShape(child, ids, offsetX, offsetY, applyOffset: false, shapeMap, connectionPointMap));
            }

            return clone;
        }

        private static VisioConnector CloneConnector(
            VisioConnector source,
            IdAllocator ids,
            VisioShape clonedFrom,
            VisioShape clonedTo,
            double offsetX,
            double offsetY,
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap) {
            VisioConnector clone = new(ids.Next(), clonedFrom, clonedTo) {
                Kind = source.Kind,
                BeginArrow = source.BeginArrow,
                EndArrow = source.EndArrow,
                Label = source.Label,
                LabelPlacement = CloneLabelPlacement(source.LabelPlacement, offsetX, offsetY),
                TextStyle = source.TextStyle?.Clone(),
                LineColor = source.LineColor,
                LineWeight = source.LineWeight,
                LinePattern = source.LinePattern,
                RouteStyle = source.RouteStyle,
                RouteAppearance = source.RouteAppearance,
                LineJumpStyle = source.LineJumpStyle,
                LineJumpCode = source.LineJumpCode,
                HorizontalJumpDirection = source.HorizontalJumpDirection,
                VerticalJumpDirection = source.VerticalJumpDirection,
                RerouteBehavior = source.RerouteBehavior,
                PreservedTextElement = source.PreservedTextElement == null ? null : new XElement(source.PreservedTextElement),
                PreservedTextValue = source.PreservedTextValue,
                HasModeledCharSection = source.HasModeledCharSection,
                HasModeledParaSection = source.HasModeledParaSection
            };

            if (source.FromConnectionPoint != null &&
                connectionPointMap.TryGetValue(source.FromConnectionPoint, out VisioConnectionPoint? clonedFromConnectionPoint)) {
                clone.FromConnectionPoint = clonedFromConnectionPoint;
            }

            if (source.ToConnectionPoint != null &&
                connectionPointMap.TryGetValue(source.ToConnectionPoint, out VisioConnectionPoint? clonedToConnectionPoint)) {
                clone.ToConnectionPoint = clonedToConnectionPoint;
            }

            foreach (VisioConnectorWaypoint waypoint in source.Waypoints) {
                clone.Waypoints.Add(new VisioConnectorWaypoint(waypoint.X + offsetX, waypoint.Y + offsetY));
            }

            CopyStringSet(source.LayerNames, clone.LayerNames);
            CopyHyperlinks(source.Hyperlinks, clone.Hyperlinks);
            CopyDictionary(source.Data, clone.Data);
            CopyShapeData(source.ShapeData, clone.ShapeData, clone.Data);
            CopyProtection(source.Protection, clone.Protection);
            CopyElements(source.PreservedGeometrySections, clone.PreservedGeometrySections);
            CopyElements(source.PreservedCellElements, clone.PreservedCellElements);
            CopyElements(source.PreservedNonGeometrySections, clone.PreservedNonGeometrySections);
            CopyElements(source.PreservedDataRows, clone.PreservedDataRows);
            CopyConnectorPreservation(source, clone);
            return clone;
        }

        private static VisioConnectorLabelPlacement? CloneLabelPlacement(VisioConnectorLabelPlacement? source, double offsetX, double offsetY) {
            if (source == null) {
                return null;
            }

            VisioConnectorLabelPlacement clone = source.Clone();
            if (clone.AbsolutePinX.HasValue) {
                clone.AbsolutePinX += offsetX;
            }

            if (clone.AbsolutePinY.HasValue) {
                clone.AbsolutePinY += offsetY;
            }

            return clone;
        }

        private static void RemapContainerMembership(Dictionary<VisioShape, VisioShape> shapeMap) {
            Dictionary<string, string> idMap = shapeMap.ToDictionary(pair => pair.Key.Id, pair => pair.Value.Id, StringComparer.OrdinalIgnoreCase);
            foreach (VisioShape clone in shapeMap.Values) {
                RemapIds(clone.ContainerMemberIds, idMap);
                RemapIds(clone.ContainerOwnerIds, idMap);
            }
        }

        private static void RemapIds(IList<string> ids, IReadOnlyDictionary<string, string> idMap) {
            for (int i = 0; i < ids.Count; i++) {
                if (idMap.TryGetValue(ids[i], out string? newId)) {
                    ids[i] = newId;
                }
            }
        }

        private static void CopyConnectionPoints(VisioShape source, VisioShape clone, Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap) {
            foreach (VisioConnectionPoint point in source.ConnectionPoints) {
                VisioConnectionPoint clonedPoint = new(point.X, point.Y, point.DirX, point.DirY) {
                    SectionIndex = point.SectionIndex
                };
                clone.ConnectionPoints.Add(clonedPoint);
                connectionPointMap[point] = clonedPoint;
            }
        }

        private static void CopyHyperlinks(IEnumerable<VisioHyperlink> source, IList<VisioHyperlink> target) {
            foreach (VisioHyperlink hyperlink in source) {
                VisioHyperlink clone = new(hyperlink.Address, hyperlink.Description, hyperlink.SubAddress) {
                    RowName = hyperlink.RowName,
                    ExtraInfo = hyperlink.ExtraInfo,
                    Frame = hyperlink.Frame,
                    NewWindow = hyperlink.NewWindow,
                    Default = hyperlink.Default,
                    Invisible = hyperlink.Invisible,
                    SortKey = hyperlink.SortKey,
                    RowIndex = hyperlink.RowIndex
                };
                CopyAttributes(hyperlink.PreservedRowAttributes, clone.PreservedRowAttributes);
                CopyElements(hyperlink.PreservedCells, clone.PreservedCells);
                foreach (KeyValuePair<string, XElement> cell in hyperlink.PreservedKnownCells) {
                    clone.PreservedKnownCells[cell.Key] = new XElement(cell.Value);
                }

                target.Add(clone);
            }
        }

        private static void CopyUserCells(IEnumerable<VisioUserCell> source, IList<VisioUserCell> target) {
            foreach (VisioUserCell userCell in source) {
                VisioUserCell clone = new(userCell.Name, userCell.Value) {
                    Unit = userCell.Unit,
                    Formula = userCell.Formula,
                    Prompt = userCell.Prompt,
                    PromptFormula = userCell.PromptFormula,
                    RowIndex = userCell.RowIndex
                };
                CopyAttributes(userCell.PreservedRowAttributes, clone.PreservedRowAttributes);
                CopyAttributes(userCell.PreservedValueAttributes, clone.PreservedValueAttributes);
                CopyAttributes(userCell.PreservedPromptAttributes, clone.PreservedPromptAttributes);
                CopyElements(userCell.PreservedCells, clone.PreservedCells);
                target.Add(clone);
            }
        }

        private static void CopyShapeData(IEnumerable<VisioShapeDataRow> source, IList<VisioShapeDataRow> target, IDictionary<string, string> data) {
            foreach (VisioShapeDataRow row in source) {
                VisioShapeDataRow clone = new(row.Name, row.Value) {
                    ValueUnit = row.ValueUnit,
                    ValueFormula = row.ValueFormula,
                    Label = row.Label,
                    LabelFormula = row.LabelFormula,
                    Prompt = row.Prompt,
                    PromptFormula = row.PromptFormula,
                    Type = row.Type,
                    TypeFormula = row.TypeFormula,
                    Format = row.Format,
                    FormatFormula = row.FormatFormula,
                    SortKey = row.SortKey,
                    SortKeyFormula = row.SortKeyFormula,
                    Invisible = row.Invisible,
                    InvisibleFormula = row.InvisibleFormula,
                    Verify = row.Verify,
                    VerifyFormula = row.VerifyFormula,
                    DataLinked = row.DataLinked,
                    DataLinkedFormula = row.DataLinkedFormula,
                    Calendar = row.Calendar,
                    CalendarFormula = row.CalendarFormula,
                    LangId = row.LangId,
                    LangIdFormula = row.LangIdFormula,
                    LoadedValue = row.LoadedValue,
                    RowIndex = row.RowIndex
                };
                CopyAttributes(row.PreservedRowAttributes, clone.PreservedRowAttributes);
                CopyElements(row.PreservedCells, clone.PreservedCells);
                foreach (KeyValuePair<string, XElement> cell in row.PreservedKnownCells) {
                    clone.PreservedKnownCells[cell.Key] = new XElement(cell.Value);
                }

                foreach (string cellName in row.PreservedCellOrder) {
                    clone.PreservedCellOrder.Add(cellName);
                }

                target.Add(clone);
                if (clone.Value != null) {
                    data[clone.Name] = clone.Value;
                }
            }
        }

        private static void CopyProtection(VisioProtection source, VisioProtection target) {
            foreach (string cellName in VisioProtection.CellNames) {
                if (source.TryGetCellValue(cellName, out bool? value)) {
                    target.TrySetCellValue(cellName, value);
                }
            }
        }

        private static string ResolveDuplicatePageName(VisioDocument document, VisioPage sourcePage, string? requestedName) {
            if (!string.IsNullOrWhiteSpace(requestedName)) {
                return requestedName!;
            }

            string baseName = $"{sourcePage.Name} Copy";
            string candidate = baseName;
            int suffix = 2;
            while (document.Pages.Any(page => string.Equals(page.Name, candidate, StringComparison.OrdinalIgnoreCase))) {
                candidate = $"{baseName} {suffix.ToString(CultureInfo.InvariantCulture)}";
                suffix++;
            }

            return candidate;
        }

        private static void CopyPageSettings(VisioPage source, VisioPage target) {
            target.DefaultUnit = source.DefaultUnit;
            target.ScaleMeasurementUnit = source.ScaleMeasurementUnit;
            target.ApplyLoadedPageScale(source.GetEffectivePageScale());
            target.ApplyLoadedDrawingScale(source.GetEffectiveDrawingScale());
            target.ViewScale = source.ViewScale;
            target.ViewCenterX = source.ViewCenterX;
            target.ViewCenterY = source.ViewCenterY;
            target.GridVisible = source.GridVisible;
            target.Snap = source.Snap;
            target.PageLockReplace = source.PageLockReplace;
            target.PageLockDuplicate = source.PageLockDuplicate;
            target.DrawingSizeType = source.DrawingSizeType;
            target.AutoResizeDrawing = source.AutoResizeDrawing;
            target.AllowShapeSplitting = source.AllowShapeSplitting;
            target.UiVisibility = source.UiVisibility;
            target.PlacementStyle = source.PlacementStyle;
            target.PlacementDepth = source.PlacementDepth;
            target.PlacementFlip = source.PlacementFlip;
            target.MoveShapesAwayOnDrop = source.MoveShapesAwayOnDrop;
            target.ResizePageToFitLayout = source.ResizePageToFitLayout;
            target.EnableLayoutGrid = source.EnableLayoutGrid;
            target.ConnectorRouteStyle = source.ConnectorRouteStyle;
            target.ConnectorRouteAppearance = source.ConnectorRouteAppearance;
            target.LineJumpStyle = source.LineJumpStyle;
            target.LineJumpCode = source.LineJumpCode;
            target.HorizontalLineJumpDirection = source.HorizontalLineJumpDirection;
            target.VerticalLineJumpDirection = source.VerticalLineJumpDirection;
            target.PrintOrientation = source.PrintOrientation;

            if (source.HasExplicitMargins) {
                target.SetMargins(
                    source.LeftMargin.FromInches(source.MarginUnit),
                    source.RightMargin.FromInches(source.MarginUnit),
                    source.TopMargin.FromInches(source.MarginUnit),
                    source.BottomMargin.FromInches(source.MarginUnit),
                    source.MarginUnit);
            }

            if (source.HasConnectorSpacing) {
                target.SetLoadedConnectorSpacingInches(
                    source.LineToLineX,
                    source.LineToLineY,
                    source.LineToNodeX,
                    source.LineToNodeY,
                    source.ConnectorSpacingUnit);
            }

            if (source.HasLayoutGridSizing) {
                target.SetLoadedLayoutGridSizingInches(
                    source.LayoutBlockSizeX,
                    source.LayoutBlockSizeY,
                    source.LayoutAvenueSizeX,
                    source.LayoutAvenueSizeY,
                    source.LayoutGridUnit);
            }
        }

        private static void CopyLayers(VisioPage source, VisioPage target) {
            foreach (VisioLayer layer in source.Layers) {
                VisioLayer clone = new(layer.Name, layer.NameU) {
                    Color = layer.Color,
                    Status = layer.Status,
                    Visible = layer.Visible,
                    Print = layer.Print,
                    Active = layer.Active,
                    Lock = layer.Lock,
                    Snap = layer.Snap,
                    Glue = layer.Glue,
                    ColorTransparency = layer.ColorTransparency
                };
                CopyAttributes(layer.PreservedRowAttributes, clone.PreservedRowAttributes);
                foreach (KeyValuePair<string, XElement> cell in layer.PreservedKnownCells) {
                    clone.PreservedKnownCells[cell.Key] = new XElement(cell.Value);
                }
                CopyElements(layer.PreservedCells, clone.PreservedCells);
                target.Layers.Add(clone);
            }
        }

        private static void CopyPagePreservation(VisioPage source, VisioPage target) {
            CopyAttributes(source.PreservedPageAttributes, target.PreservedPageAttributes);
            CopyAttributes(source.PreservedPageContentAttributes, target.PreservedPageContentAttributes);
            CopyElements(source.PreservedPageContentElements, target.PreservedPageContentElements);
            CopyAttributes(source.PreservedShapesContainerAttributes, target.PreservedShapesContainerAttributes);
            CopyElements(source.PreservedShapesContainerElements, target.PreservedShapesContainerElements);
            CopyAttributes(source.PreservedConnectsAttributes, target.PreservedConnectsAttributes);
            CopyElements(source.PreservedConnectsElements, target.PreservedConnectsElements);
            CopyElements(source.PreservedPageSheetCells, target.PreservedPageSheetCells);
            CopyElements(source.PreservedPageSheetSections, target.PreservedPageSheetSections);
        }

        private static void CopyConnectorPreservation(VisioConnector source, VisioConnector target) {
            target.PreservedFromConnectionCell = source.PreservedFromConnectionCell;
            target.PreservedToConnectionCell = source.PreservedToConnectionCell;
            CopyAttributes(source.PreservedBeginConnectAttributes, target.PreservedBeginConnectAttributes);
            CopyAttributes(source.PreservedEndConnectAttributes, target.PreservedEndConnectAttributes);
            CopyNames(source.PreservedBeginConnectAttributeOrder, target.PreservedBeginConnectAttributeOrder);
            CopyNames(source.PreservedEndConnectAttributeOrder, target.PreservedEndConnectAttributeOrder);
        }

        private static void CopyStringSet(IEnumerable<string> source, ISet<string> target) {
            foreach (string value in source) {
                target.Add(value);
            }
        }

        private static void CopyStringList(IEnumerable<string> source, IList<string> target) {
            foreach (string value in source) {
                target.Add(value);
            }
        }

        private static void CopyDictionary(IEnumerable<KeyValuePair<string, string>> source, IDictionary<string, string> target) {
            foreach (KeyValuePair<string, string> pair in source) {
                target[pair.Key] = pair.Value;
            }
        }

        private static void CopyAttributes(IEnumerable<XAttribute> source, IList<XAttribute> target) {
            foreach (XAttribute attribute in source) {
                target.Add(new XAttribute(attribute));
            }
        }

        private static void CopyNames(IEnumerable<XName> source, IList<XName> target) {
            foreach (XName name in source) {
                target.Add(name);
            }
        }

        private static void CopyElements(IEnumerable<XElement> source, IList<XElement> target) {
            foreach (XElement element in source) {
                target.Add(new XElement(element));
            }
        }

        private sealed class IdAllocator {
            private readonly HashSet<int> _usedIds = new();

            public IdAllocator(VisioPage page, VisioPage? sourcePage = null) {
                foreach (VisioShape shape in page.AllShapes()) {
                    Reserve(shape.Id);
                }

                foreach (VisioConnector connector in page.Connectors) {
                    Reserve(connector.Id);
                }

                if (sourcePage != null) {
                    foreach (VisioShape shape in sourcePage.AllShapes()) {
                        Reserve(shape.Id);
                    }

                    foreach (VisioConnector connector in sourcePage.Connectors) {
                        Reserve(connector.Id);
                    }
                }
            }

            public string Next() {
                int id = 1;
                while (_usedIds.Contains(id)) {
                    id++;
                }

                _usedIds.Add(id);
                return id.ToString(CultureInfo.InvariantCulture);
            }

            private void Reserve(string? id) {
                if (int.TryParse(id, out int numericId) && numericId > 0) {
                    _usedIds.Add(numericId);
                }
            }
        }
    }
}
