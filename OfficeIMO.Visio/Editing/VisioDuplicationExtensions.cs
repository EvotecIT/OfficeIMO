using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for duplicating Visio content while keeping copied shapes independent.
    /// </summary>
    public static partial class VisioDuplicationExtensions {
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
            var sourceIdentifiers = VisioDocument.AssignPageElementIdentifiers(sourcePage);
            string duplicateName = ResolveDuplicatePageName(document, sourcePage, name);
            VisioPage clone = document.AddPage(duplicateName, sourcePage.Width, sourcePage.Height);
            clone.NameU = duplicateName;
            CopyPageSettings(sourcePage, clone);
            CopyLayers(sourcePage, clone);
            CopyPagePreservation(sourcePage, clone);

            IdAllocator ids = new(clone, sourcePage);
            Dictionary<VisioShape, VisioShape> shapeMap = new();
            Dictionary<VisioConnector, VisioConnector> connectorMap = new();
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap = new();

            foreach (VisioShape shape in sourcePage.Shapes) {
                VisioShape shapeClone = CloneShape(shape, ids, 0D, 0D, applyOffset: false, shapeMap, connectionPointMap, sourceDocument: document);
                clone.Shapes.Add(shapeClone);
            }

            RemapContainerMembership(shapeMap);

            foreach (VisioConnector connector in sourcePage.Connectors) {
                VisioShape? clonedFrom = null, clonedTo = null;
                if ((connector.From != null && !shapeMap.TryGetValue(connector.From, out clonedFrom)) ||
                    (connector.To != null && !shapeMap.TryGetValue(connector.To, out clonedTo))) {
                    continue;
                }

                VisioConnector connectorClone = CloneConnector(connector, ids, clonedFrom, clonedTo, 0D, 0D, connectionPointMap, sourceDocument: document);
                clone.Connectors.Add(connectorClone);
                connectorMap[connector] = connectorClone;
            }

            RemapCopiedFormulaReferences(clone, shapeMap, connectorMap, sourceIdentifiers, remapPageSheet: true);
            CopyComments(sourcePage, clone, shapeMap, connectorMap);

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

        private static void CopyComments(
            VisioPage sourcePage,
            VisioPage clone,
            IReadOnlyDictionary<VisioShape, VisioShape> shapeMap,
            IReadOnlyDictionary<VisioConnector, VisioConnector> connectorMap) {
            Dictionary<string, string> targetMap = new(StringComparer.Ordinal);
            foreach (KeyValuePair<VisioShape, VisioShape> pair in shapeMap) {
                targetMap[pair.Key.Id] = pair.Value.Id;
            }

            foreach (KeyValuePair<VisioConnector, VisioConnector> pair in connectorMap) {
                targetMap[pair.Key.Id] = pair.Value.Id;
            }

            foreach (VisioComment comment in sourcePage.Comments) {
                VisioComment copy = CloneComment(comment);
                if (!string.IsNullOrWhiteSpace(comment.ShapeId) &&
                    targetMap.TryGetValue(comment.ShapeId!, out string? clonedTargetId)) {
                    copy.ShapeId = clonedTargetId;
                }

                clone.Comments.Add(copy);
            }
        }

        private static VisioComment CloneComment(VisioComment comment) {
            return new VisioComment(comment.Text) {
                Id = comment.Id,
                AuthorName = comment.AuthorName,
                AuthorInitials = comment.AuthorInitials,
                AuthorResolutionId = comment.AuthorResolutionId,
                ShapeId = comment.ShapeId,
                CreatedAt = comment.CreatedAt,
                EditedAt = comment.EditedAt,
                Done = comment.Done,
                AutoCommentType = comment.AutoCommentType
            };
        }

        private static void CopyTargetedComments(
            VisioPage page,
            IReadOnlyDictionary<VisioShape, VisioShape> shapeMap,
            IReadOnlyDictionary<VisioConnector, VisioConnector> connectorMap) {
            Dictionary<string, string> targetMap = new(StringComparer.Ordinal);
            foreach (KeyValuePair<VisioShape, VisioShape> pair in shapeMap) {
                targetMap[pair.Key.Id] = pair.Value.Id;
            }

            foreach (KeyValuePair<VisioConnector, VisioConnector> pair in connectorMap) {
                targetMap[pair.Key.Id] = pair.Value.Id;
            }

            if (targetMap.Count == 0) {
                return;
            }

            HashSet<int> usedCommentIds = new(page.Comments.Select(comment => comment.Id));
            foreach (VisioComment comment in page.Comments.ToList()) {
                if (string.IsNullOrWhiteSpace(comment.ShapeId) ||
                    !targetMap.TryGetValue(comment.ShapeId!, out string? clonedTargetId)) {
                    continue;
                }

                VisioComment copy = CloneComment(comment);
                copy.Id = AllocateNextCommentId(usedCommentIds);
                copy.ShapeId = clonedTargetId;
                page.Comments.Add(copy);
            }
        }

        private static int AllocateNextCommentId(HashSet<int> usedCommentIds) {
            int nextId = 1;
            while (usedCommentIds.Contains(nextId)) {
                nextId++;
            }

            usedCommentIds.Add(nextId);
            return nextId;
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
        /// Duplicates shapes on the same page and optionally copies connectors whose attached shapes are all duplicated; fully free connectors are excluded.
        /// </summary>
        /// <param name="page">Page that owns the shapes.</param>
        /// <param name="shapes">Shapes to duplicate. Nested children are copied with their selected ancestor.</param>
        /// <param name="offsetX">Horizontal offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="offsetY">Vertical offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="includeInternalConnectors">Whether connectors whose attached shapes are all duplicated should also be copied, including partial connectors.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection DuplicateShapes(this VisioPage page, IEnumerable<VisioShape> shapes, double offsetX = DefaultDuplicateOffsetX, double offsetY = DefaultDuplicateOffsetY, bool includeInternalConnectors = true) {
            return DuplicateShapes(page, shapes, new VisioShapeDuplicationOptions {
                OffsetX = offsetX,
                OffsetY = offsetY,
                IncludeInternalConnectors = includeInternalConnectors
            });
        }

        /// <summary>
        /// Duplicates shapes on the same page and optionally copies connectors whose attached shapes are all duplicated; fully free connectors are excluded.
        /// </summary>
        /// <param name="page">Page that owns the shapes.</param>
        /// <param name="shapes">Shapes to duplicate. Nested children are copied with their selected ancestor.</param>
        /// <param name="options">Optional duplication settings.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection DuplicateShapes(this VisioPage page, IEnumerable<VisioShape> shapes, VisioShapeDuplicationOptions? options) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            VisioShapeDuplicationOptions effectiveOptions = options ?? new VisioShapeDuplicationOptions();
            EnsureFiniteOffset(effectiveOptions.OffsetX, nameof(effectiveOptions.OffsetX));
            EnsureFiniteOffset(effectiveOptions.OffsetY, nameof(effectiveOptions.OffsetY));

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

            var sourceIdentifiers = VisioDocument.AssignPageElementIdentifiers(page);
            IdAllocator ids = new(page);
            Dictionary<VisioShape, VisioShape> shapeMap = new();
            Dictionary<VisioConnector, VisioConnector> connectorMap = new();
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap = new();
            List<VisioShape> duplicatedRoots = new();

            foreach (VisioShape root in rootShapes) {
                VisioShape clone = CloneShape(root, ids, effectiveOptions.OffsetX, effectiveOptions.OffsetY, applyOffset: true, shapeMap, connectionPointMap, effectiveOptions, page.OwnerDocument);
                page.Shapes.Add(clone);
                duplicatedRoots.Add(clone);
            }

            RemapContainerMembership(shapeMap);

            if (effectiveOptions.IncludeInternalConnectors) {
                foreach (VisioConnector connector in page.Connectors.ToList()) {
                    VisioShape? clonedFrom = null, clonedTo = null;
                    if ((connector.From == null && connector.To == null) ||
                        (connector.From != null && !shapeMap.TryGetValue(connector.From, out clonedFrom)) ||
                        (connector.To != null && !shapeMap.TryGetValue(connector.To, out clonedTo))) {
                        continue;
                    }

                    VisioConnector clonedConnector = CloneConnector(connector, ids, clonedFrom, clonedTo, effectiveOptions.OffsetX, effectiveOptions.OffsetY, connectionPointMap, effectiveOptions, page.OwnerDocument);
                    page.Connectors.Add(clonedConnector);
                    connectorMap[connector] = clonedConnector;
                }
            }

            RemapCopiedFormulaReferences(page, shapeMap, connectorMap, sourceIdentifiers);
            CopyTargetedComments(page, shapeMap, connectorMap);

            return new VisioShapeSelection(duplicatedRoots, page);
        }

        /// <summary>
        /// Duplicates a page-backed selection on the same page.
        /// </summary>
        /// <param name="selection">Selection to duplicate.</param>
        /// <param name="offsetX">Horizontal offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="offsetY">Vertical offset for duplicated top-level shapes and page-coordinate routing points.</param>
        /// <param name="includeInternalConnectors">Whether connectors whose attached shapes are all duplicated should also be copied, including partial connectors.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection Duplicate(this VisioShapeSelection selection, double offsetX = DefaultDuplicateOffsetX, double offsetY = DefaultDuplicateOffsetY, bool includeInternalConnectors = true) {
            return Duplicate(selection, new VisioShapeDuplicationOptions {
                OffsetX = offsetX,
                OffsetY = offsetY,
                IncludeInternalConnectors = includeInternalConnectors
            });
        }

        /// <summary>
        /// Duplicates a page-backed selection on the same page.
        /// </summary>
        /// <param name="selection">Selection to duplicate.</param>
        /// <param name="options">Optional duplication settings.</param>
        /// <returns>A selection containing the duplicated root shapes.</returns>
        public static VisioShapeSelection Duplicate(this VisioShapeSelection selection, VisioShapeDuplicationOptions? options) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Use page.DuplicateShapes(selection, ...) instead.");
            }

            return selection.OwnerPage.DuplicateShapes(selection, options);
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

        private static VisioConnector CloneConnector(
            VisioConnector source,
            IdAllocator ids,
            VisioShape? clonedFrom,
            VisioShape? clonedTo,
            double offsetX,
            double offsetY,
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap,
            VisioShapeDuplicationOptions? options = null,
            VisioDocument? sourceDocument = null) {
            VisioConnector clone = new(ResolveConnectorCloneId(source, ids, options),
                new OfficeIMO.Drawing.OfficePoint(source.StartPoint.X + offsetX, source.StartPoint.Y + offsetY),
                new OfficeIMO.Drawing.OfficePoint(source.EndPoint.X + offsetX, source.EndPoint.Y + offsetY)) {
                From = clonedFrom, To = clonedTo,
                StartAttachment = source.StartAttachment, EndAttachment = source.EndAttachment,
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
                HasModeledParaSection = source.HasModeledParaSection,
                CharacterSectionSource = source.CharacterSectionSource?.Clone(),
                ParagraphSectionSource = source.ParagraphSectionSource?.Clone(),
                ShapeDataSectionName = source.ShapeDataSectionName,
                NativeStyleReferences = source.NativeStyleReferences,
                PreserveDynamicConnectorMaster = source.PreserveDynamicConnectorMaster,
                NativeCellMetadata = source.NativeCellMetadata?.Clone()
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
            clone.NativeLayerMembership = source.NativeLayerMembership?.Clone();
            CopyHyperlinks(source.Hyperlinks, clone.Hyperlinks, VisioHyperlinkRowNames.Inherited(source, sourceDocument));
            CopyDictionary(source.Data, clone.Data);
            CopyShapeData(source.ShapeData, clone.ShapeData, clone.Data);
            CopyProtection(source.Protection, clone.Protection);
            CopyElements(source.PreservedGeometrySections, clone.PreservedGeometrySections);
            CopyElements(source.PreservedCellElements, clone.PreservedCellElements);
            CopyElements(source.PreservedNonGeometrySections, clone.PreservedNonGeometrySections);
            CopyElements(source.PreservedDataRows, clone.PreservedDataRows);
            clone.ForeignResources.AddRange(source.ForeignResources);
            if (source.ForeignResources.Count > 0) {
                foreach (VisioConnector.PreservedShapeChildEntry entry in source.PreservedShapeChildren)
                    clone.PreservedShapeChildren.Add(entry.RawElement != null
                        ? new VisioConnector.PreservedShapeChildEntry(entry.RawElement)
                        : new VisioConnector.PreservedShapeChildEntry(entry.Token!));
            }
            CopyConnectorPreservation(source, clone);
            clone.NativeGeometry = source.NativeGeometry?.AppliesTo(source) == true ? source.NativeGeometry.CopyFor(clone) : null;
            return clone;
        }

        private static string ResolveShapeCloneId(VisioShape source, IdAllocator ids, VisioShapeDuplicationOptions? options) {
            string? preferred = options?.ShapeIdFactory?.Invoke(source);
            string? suffix = options?.IdSuffix;
            if (string.IsNullOrWhiteSpace(preferred) &&
                !string.IsNullOrEmpty(suffix) &&
                !string.IsNullOrWhiteSpace(source.Id)) {
                preferred = source.Id + suffix;
            }

            return ids.Next(preferred);
        }

        private static string ResolveConnectorCloneId(VisioConnector source, IdAllocator ids, VisioShapeDuplicationOptions? options) {
            string? preferred = options?.ConnectorIdFactory?.Invoke(source);
            string? suffix = options?.ConnectorIdSuffix ?? options?.IdSuffix;
            if (string.IsNullOrWhiteSpace(preferred) &&
                !string.IsNullOrEmpty(suffix) &&
                !string.IsNullOrWhiteSpace(source.Id)) {
                preferred = source.Id + suffix;
            }

            return ids.Next(preferred);
        }

        private static void EnsureFiniteOffset(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(parameterName, "Offset must be a finite number.");
            }
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

        private static void CopyHyperlinks(IList<VisioHyperlink> source, IList<VisioHyperlink> target, IList<VisioHyperlink>? inherited = null) {
            string?[] names = VisioHyperlinkRowNames.Create(source, inherited);
            for (int i = 0; i < source.Count; i++) {
                VisioHyperlink hyperlink = source[i];
                VisioHyperlink clone = new(hyperlink.Address, hyperlink.Description, hyperlink.SubAddress) {
                    RowName = names[i],
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

                clone.CopyValueAssignmentsFrom(hyperlink);
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
                clone.CopyValueAssignmentsFrom(userCell);
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
                    MirroredDataValue = row.MirroredDataValue,
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

                clone.CopyValueAssignmentsFrom(row);
                target.Add(clone);
                if (clone.Value != null && !data.ContainsKey(clone.Name)) {
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
                target.SetLoadedMargins(
                    source.LeftMargin,
                    source.RightMargin,
                    source.TopMargin,
                    source.BottomMargin,
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
                    ColorTransparency = layer.ColorTransparency,
                    SourceIndex = layer.SourceIndex
                };
                CopyAttributes(layer.PreservedRowAttributes, clone.PreservedRowAttributes);
                foreach (KeyValuePair<string, XElement> cell in layer.PreservedKnownCells) {
                    clone.PreservedKnownCells[cell.Key] = new XElement(cell.Value);
                }
                CopyElements(layer.PreservedCells, clone.PreservedCells);
                clone.CopyValueAssignmentsFrom(layer);
                target.Layers.Add(clone);
            }
        }

        private static void CopyPagePreservation(VisioPage source, VisioPage target) {
            target.NativePageSheetMetadata = source.NativePageSheetMetadata?.Clone();
            target.PageSheetLengthCells = source.PageSheetLengthCells?.Clone();
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
            private readonly HashSet<string> _usedTextIds = new(StringComparer.OrdinalIgnoreCase);

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

            public string Next(string? preferredId = null) {
                if (!string.IsNullOrWhiteSpace(preferredId)) {
                    return ReservePreferred(preferredId!.Trim());
                }

                int id = 1;
                while (_usedIds.Contains(id) || _usedTextIds.Contains(id.ToString(CultureInfo.InvariantCulture))) {
                    id++;
                }

                string value = id.ToString(CultureInfo.InvariantCulture);
                Reserve(value);
                return value;
            }

            private void Reserve(string? id) {
                if (string.IsNullOrWhiteSpace(id)) {
                    return;
                }

                string resolvedId = id!;
                _usedTextIds.Add(resolvedId);
                if (int.TryParse(resolvedId, out int numericId) && numericId > 0) {
                    _usedIds.Add(numericId);
                }
            }

            private string ReservePreferred(string preferredId) {
                string candidate = preferredId;
                int suffix = 2;
                while (_usedTextIds.Contains(candidate)) {
                    candidate = preferredId + "-" + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                }

                Reserve(candidate);
                return candidate;
            }
        }
    }
}
