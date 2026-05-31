using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Creates deterministic inspection snapshots for generated or loaded Visio documents.
    /// </summary>
    public static class VisioInspectionExtensions {
        /// <summary>
        /// Creates a stable, data-oriented snapshot of the document structure.
        /// </summary>
        public static VisioInspectionSnapshot CreateInspectionSnapshot(this VisioDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            IReadOnlyList<VisioInspectionMasterSnapshot> masters = document.Masters
                .OrderBy(master => master.NameU, StringComparer.OrdinalIgnoreCase)
                .ThenBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                .Select(master => new VisioInspectionMasterSnapshot(
                    master.Id,
                    master.NameU,
                    master.Shape.NameU,
                    master.Shape.Text,
                    master.Shape.Width,
                    master.Shape.Height,
                    master.IsPackageBacked,
                    master.StencilId,
                    master.StencilName,
                    master.StencilCategory,
                    master.StencilCatalogName,
                    master.StencilSourcePackagePath,
                    master.StencilKeywords,
                    master.StencilAliases,
                    master.StencilTags,
                    master.StencilIconNameU,
                    master.StencilDefaultWidth,
                    master.StencilDefaultHeight,
                    master.StencilDefaultUnit?.ToString(),
                    master.StencilPreviewImageRelationshipId,
                    master.StencilPreviewImageTarget,
                    master.StencilPreviewImageContentType,
                    master.StencilPreviewImageExtension,
                    master.StencilPreviewImageByteLength))
                .ToList()
                .AsReadOnly();

            IReadOnlyList<VisioInspectionPageSnapshot> pages = document.Pages
                .OrderBy(page => page.Id)
                .ThenBy(page => page.Name, StringComparer.OrdinalIgnoreCase)
                .Select(CreatePageSnapshot)
                .ToList()
                .AsReadOnly();

            return new VisioInspectionSnapshot(
                document.Title,
                document.Author,
                document.Theme != null ? document.Theme.GetType().Name : null,
                document.UseMastersByDefault,
                document.WriteMasterDeltasOnly,
                masters,
                pages);
        }

        private static VisioInspectionPageSnapshot CreatePageSnapshot(VisioPage page) {
            IReadOnlyList<VisioInspectionShapeSnapshot> shapes = page.AllShapes()
                .OrderBy(shape => shape.Id, StringComparer.OrdinalIgnoreCase)
                .Select(CreateShapeSnapshot)
                .ToList()
                .AsReadOnly();

            IReadOnlyList<VisioInspectionConnectorSnapshot> connectors = page.Connectors
                .OrderBy(connector => connector.Id, StringComparer.OrdinalIgnoreCase)
                .Select(CreateConnectorSnapshot)
                .ToList()
                .AsReadOnly();

            IReadOnlyList<string> layers = page.Layers
                .Select(layer => layer.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            return new VisioInspectionPageSnapshot(
                page.Id,
                page.Name,
                page.NameU,
                page.Width,
                page.Height,
                layers,
                shapes,
                connectors);
        }

        private static VisioInspectionShapeSnapshot CreateShapeSnapshot(VisioShape shape) {
            return new VisioInspectionShapeSnapshot(
                shape.Id,
                shape.Name,
                shape.NameU,
                shape.Type,
                shape.Master?.Id,
                shape.MasterNameU,
                shape.MasterShapeId,
                shape.Parent?.Id,
                shape.Text,
                shape.PinX,
                shape.PinY,
                shape.Width,
                shape.Height,
                shape.Angle,
                shape.LineColor.ToString(),
                shape.FillColor.ToString(),
                shape.LinePattern,
                shape.FillPattern,
                shape.LineWeight,
                shape.IsContainer,
                shape.IsCallout,
                shape.IsBackgroundSurface,
                shape.IsDiagramAdornment,
                shape.CalloutTargetId,
                SortStrings(shape.LayerNames),
                CreateShapeDataSnapshot(shape.ShapeData),
                CreateUserCellSnapshot(shape.UserCells),
                CreateDataSnapshot(shape.Data),
                CreateConnectionPointSnapshot(shape.ConnectionPoints),
                shape.Children.Select(child => child.Id).OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly());
        }

        private static VisioInspectionConnectorSnapshot CreateConnectorSnapshot(VisioConnector connector) {
            return new VisioInspectionConnectorSnapshot(
                connector.Id,
                connector.From.Id,
                connector.To.Id,
                connector.Kind.ToString(),
                connector.Label,
                connector.LabelPlacement != null,
                connector.LabelPlacement?.Position,
                connector.LabelPlacement?.OffsetX,
                connector.LabelPlacement?.OffsetY,
                connector.LabelPlacement?.PinX,
                connector.LabelPlacement?.PinY,
                connector.LabelPlacement?.Width,
                connector.LabelPlacement?.Height,
                connector.Waypoints
                    .Select(waypoint => new VisioInspectionWaypointSnapshot(waypoint.X, waypoint.Y))
                    .ToList()
                    .AsReadOnly(),
                connector.LineColor.ToString(),
                connector.LinePattern,
                connector.LineWeight,
                connector.BeginArrow?.ToString(),
                connector.EndArrow?.ToString(),
                SortStrings(connector.LayerNames),
                CreateShapeDataSnapshot(connector.ShapeData),
                CreateDataSnapshot(connector.Data));
        }

        private static IReadOnlyList<VisioInspectionShapeDataSnapshot> CreateShapeDataSnapshot(IEnumerable<VisioShapeDataRow> rows) {
            return rows
                .OrderBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
                .Select(row => new VisioInspectionShapeDataSnapshot(row.Name, row.Label, row.Value, row.Type?.ToString(), row.Format, row.Prompt))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<VisioInspectionUserCellSnapshot> CreateUserCellSnapshot(IEnumerable<VisioUserCell> rows) {
            return rows
                .OrderBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
                .Select(row => new VisioInspectionUserCellSnapshot(row.Name, row.Value, row.Formula, row.Prompt))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<VisioInspectionConnectionPointSnapshot> CreateConnectionPointSnapshot(IEnumerable<VisioConnectionPoint> points) {
            return points
                .Select((point, index) => new VisioInspectionConnectionPointSnapshot(index, point.SectionIndex, point.X, point.Y, point.DirX, point.DirY))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyDictionary<string, string> CreateDataSnapshot(IDictionary<string, string> data) {
            return new ReadOnlyDictionary<string, string>(
                data.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
        }

        private static IReadOnlyList<string> SortStrings(IEnumerable<string> values) {
            return values
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }
    }

    /// <summary>
    /// Deterministic structural and semantic snapshot of a Visio document.
    /// </summary>
    public sealed class VisioInspectionSnapshot {
        internal VisioInspectionSnapshot(
            string? title,
            string? author,
            string? themeType,
            bool useMastersByDefault,
            bool writeMasterDeltasOnly,
            IReadOnlyList<VisioInspectionMasterSnapshot> masters,
            IReadOnlyList<VisioInspectionPageSnapshot> pages) {
            Title = title;
            Author = author;
            ThemeType = themeType;
            UseMastersByDefault = useMastersByDefault;
            WriteMasterDeltasOnly = writeMasterDeltasOnly;
            Masters = masters;
            Pages = pages;
        }

        /// <summary>Document title.</summary>
        public string? Title { get; }

        /// <summary>Document author.</summary>
        public string? Author { get; }

        /// <summary>Document theme type, when present.</summary>
        public string? ThemeType { get; }

        /// <summary>Whether generated shapes use masters by default.</summary>
        public bool UseMastersByDefault { get; }

        /// <summary>Whether page instances write only master deltas.</summary>
        public bool WriteMasterDeltasOnly { get; }

        /// <summary>Registered masters.</summary>
        public IReadOnlyList<VisioInspectionMasterSnapshot> Masters { get; }

        /// <summary>Document pages.</summary>
        public IReadOnlyList<VisioInspectionPageSnapshot> Pages { get; }

        /// <summary>Gets the total number of shapes, including group children.</summary>
        public int ShapeCount => Pages.Sum(page => page.Shapes.Count);

        /// <summary>Gets the total number of connectors.</summary>
        public int ConnectorCount => Pages.Sum(page => page.Connectors.Count);

        /// <summary>
        /// Compares this snapshot to another snapshot.
        /// </summary>
        public VisioInspectionDiff Diff(VisioInspectionSnapshot other) {
            return VisioInspectionDiff.Compare(this, other);
        }

        /// <summary>
        /// Writes a stable line-oriented representation suitable for golden snapshots and review diffs.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            AppendLine(builder, "document.title", Title);
            AppendLine(builder, "document.author", Author);
            AppendLine(builder, "document.theme", ThemeType);
            AppendLine(builder, "document.useMastersByDefault", UseMastersByDefault);
            AppendLine(builder, "document.writeMasterDeltasOnly", WriteMasterDeltasOnly);
            AppendLine(builder, "document.masterCount", Masters.Count);
            AppendLine(builder, "document.pageCount", Pages.Count);
            AppendLine(builder, "document.shapeCount", ShapeCount);
            AppendLine(builder, "document.connectorCount", ConnectorCount);

            foreach (VisioInspectionMasterSnapshot master in Masters) {
                string prefix = "master[" + EscapeKey(master.Id) + "]";
                AppendLine(builder, prefix + ".nameU", master.NameU);
                AppendLine(builder, prefix + ".shapeNameU", master.ShapeNameU);
                AppendLine(builder, prefix + ".text", master.Text);
                AppendLine(builder, prefix + ".width", master.Width);
                AppendLine(builder, prefix + ".height", master.Height);
                AppendLine(builder, prefix + ".packageBacked", master.IsPackageBacked);
                AppendLine(builder, prefix + ".stencilId", master.StencilId);
                AppendLine(builder, prefix + ".stencilName", master.StencilName);
                AppendLine(builder, prefix + ".stencilCategory", master.StencilCategory);
                AppendLine(builder, prefix + ".stencilCatalog", master.StencilCatalogName);
                AppendLine(builder, prefix + ".stencilSourcePackagePath", master.StencilSourcePackagePath);
                AppendLine(builder, prefix + ".stencilKeywords", string.Join(",", master.StencilKeywords));
                AppendLine(builder, prefix + ".stencilAliases", string.Join(",", master.StencilAliases));
                AppendLine(builder, prefix + ".stencilTags", string.Join(",", master.StencilTags));
                AppendLine(builder, prefix + ".stencilIconNameU", master.StencilIconNameU);
                AppendLine(builder, prefix + ".stencilDefaultWidth", master.StencilDefaultWidth);
                AppendLine(builder, prefix + ".stencilDefaultHeight", master.StencilDefaultHeight);
                AppendLine(builder, prefix + ".stencilDefaultUnit", master.StencilDefaultUnit);
                AppendLine(builder, prefix + ".stencilPreviewImageRelationshipId", master.StencilPreviewImageRelationshipId);
                AppendLine(builder, prefix + ".stencilPreviewImageTarget", master.StencilPreviewImageTarget);
                AppendLine(builder, prefix + ".stencilPreviewImageContentType", master.StencilPreviewImageContentType);
                AppendLine(builder, prefix + ".stencilPreviewImageExtension", master.StencilPreviewImageExtension);
                AppendLine(builder, prefix + ".stencilPreviewImageByteLength", master.StencilPreviewImageByteLength);
            }

            foreach (VisioInspectionPageSnapshot page in Pages) {
                page.AppendText(builder);
            }

            return builder.ToString();
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }

        internal static void AppendLine(StringBuilder builder, string key, object? value) {
            builder.Append(key);
            builder.Append('=');
            builder.Append(FormatValue(value));
            builder.AppendLine();
        }

        internal static string FormatValue(object? value) {
            if (value == null) {
                return string.Empty;
            }

            if (value is double doubleValue) {
                return doubleValue.ToString("0.######", CultureInfo.InvariantCulture);
            }

            if (value is bool boolValue) {
                return boolValue ? "true" : "false";
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        internal static string EscapeKey(string? value) {
            return string.IsNullOrEmpty(value)
                ? string.Empty
                : value!.Replace("\\", "\\\\").Replace("]", "\\]");
        }
    }

    /// <summary>
    /// Snapshot of a registered Visio master.
    /// </summary>
    public sealed class VisioInspectionMasterSnapshot {
        internal VisioInspectionMasterSnapshot(
            string id,
            string nameU,
            string? shapeNameU,
            string? text,
            double width,
            double height,
            bool isPackageBacked,
            string? stencilId,
            string? stencilName,
            string? stencilCategory,
            string? stencilCatalogName,
            string? stencilSourcePackagePath,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            string? stencilIconNameU,
            double? stencilDefaultWidth,
            double? stencilDefaultHeight,
            string? stencilDefaultUnit,
            string? stencilPreviewImageRelationshipId,
            string? stencilPreviewImageTarget,
            string? stencilPreviewImageContentType,
            string? stencilPreviewImageExtension,
            long? stencilPreviewImageByteLength) {
            Id = id;
            NameU = nameU;
            ShapeNameU = shapeNameU;
            Text = text;
            Width = width;
            Height = height;
            IsPackageBacked = isPackageBacked;
            StencilId = stencilId;
            StencilName = stencilName;
            StencilCategory = stencilCategory;
            StencilCatalogName = stencilCatalogName;
            StencilSourcePackagePath = stencilSourcePackagePath;
            StencilKeywords = stencilKeywords;
            StencilAliases = stencilAliases;
            StencilTags = stencilTags;
            StencilIconNameU = stencilIconNameU;
            StencilDefaultWidth = stencilDefaultWidth;
            StencilDefaultHeight = stencilDefaultHeight;
            StencilDefaultUnit = stencilDefaultUnit;
            StencilPreviewImageRelationshipId = stencilPreviewImageRelationshipId;
            StencilPreviewImageTarget = stencilPreviewImageTarget;
            StencilPreviewImageContentType = stencilPreviewImageContentType;
            StencilPreviewImageExtension = stencilPreviewImageExtension;
            StencilPreviewImageByteLength = stencilPreviewImageByteLength;
        }

        /// <summary>Master identifier.</summary>
        public string Id { get; }

        /// <summary>Master universal name.</summary>
        public string NameU { get; }

        /// <summary>Universal name of the master shape.</summary>
        public string? ShapeNameU { get; }

        /// <summary>Text stored on the master shape.</summary>
        public string? Text { get; }

        /// <summary>Master shape width.</summary>
        public double Width { get; }

        /// <summary>Master shape height.</summary>
        public double Height { get; }

        /// <summary>Whether the master came from a package-backed stencil or document.</summary>
        public bool IsPackageBacked { get; }

        /// <summary>OfficeIMO stencil identifier, when known.</summary>
        public string? StencilId { get; }

        /// <summary>OfficeIMO stencil display name, when known.</summary>
        public string? StencilName { get; }

        /// <summary>OfficeIMO stencil category, when known.</summary>
        public string? StencilCategory { get; }

        /// <summary>Stencil catalog name, when known.</summary>
        public string? StencilCatalogName { get; }

        /// <summary>Source package path, when known.</summary>
        public string? StencilSourcePackagePath { get; }

        /// <summary>Searchable stencil keywords.</summary>
        public IReadOnlyList<string> StencilKeywords { get; }

        /// <summary>Stencil lookup aliases.</summary>
        public IReadOnlyList<string> StencilAliases { get; }

        /// <summary>Semantic stencil tags.</summary>
        public IReadOnlyList<string> StencilTags { get; }

        /// <summary>Preview icon master universal name, when known.</summary>
        public string? StencilIconNameU { get; }

        /// <summary>Source stencil default width, when known.</summary>
        public double? StencilDefaultWidth { get; }

        /// <summary>Source stencil default height, when known.</summary>
        public double? StencilDefaultHeight { get; }

        /// <summary>Source stencil default size unit, when known.</summary>
        public string? StencilDefaultUnit { get; }

        /// <summary>Preview image relationship id, when known.</summary>
        public string? StencilPreviewImageRelationshipId { get; }

        /// <summary>Preview image relationship target, when known.</summary>
        public string? StencilPreviewImageTarget { get; }

        /// <summary>Preview image content type, when known.</summary>
        public string? StencilPreviewImageContentType { get; }

        /// <summary>Preview image extension, when known.</summary>
        public string? StencilPreviewImageExtension { get; }

        /// <summary>Preview image byte length, when known.</summary>
        public long? StencilPreviewImageByteLength { get; }
    }

    /// <summary>
    /// Snapshot of a Visio page.
    /// </summary>
    public sealed class VisioInspectionPageSnapshot {
        internal VisioInspectionPageSnapshot(
            int id,
            string name,
            string? nameU,
            double width,
            double height,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeSnapshot> shapes,
            IReadOnlyList<VisioInspectionConnectorSnapshot> connectors) {
            Id = id;
            Name = name;
            NameU = nameU;
            Width = width;
            Height = height;
            Layers = layers;
            Shapes = shapes;
            Connectors = connectors;
        }

        /// <summary>Page identifier.</summary>
        public int Id { get; }

        /// <summary>Page display name.</summary>
        public string Name { get; }

        /// <summary>Page universal name.</summary>
        public string? NameU { get; }

        /// <summary>Page width in inches.</summary>
        public double Width { get; }

        /// <summary>Page height in inches.</summary>
        public double Height { get; }

        /// <summary>Layer names used on the page.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape snapshots on the page, including group children.</summary>
        public IReadOnlyList<VisioInspectionShapeSnapshot> Shapes { get; }

        /// <summary>Connector snapshots on the page.</summary>
        public IReadOnlyList<VisioInspectionConnectorSnapshot> Connectors { get; }

        internal void AppendText(StringBuilder builder) {
            string prefix = "page[" + Escape + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".id", Id);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".nameU", NameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".width", Width);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".height", Height);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".shapeCount", Shapes.Count);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".connectorCount", Connectors.Count);

            foreach (VisioInspectionShapeSnapshot shape in Shapes) {
                shape.AppendText(builder, prefix);
            }

            foreach (VisioInspectionConnectorSnapshot connector in Connectors) {
                connector.AppendText(builder, prefix);
            }
        }

        private string Escape => VisioInspectionSnapshot.EscapeKey(Name);
    }

    /// <summary>
    /// Snapshot of a Visio shape.
    /// </summary>
    public sealed class VisioInspectionShapeSnapshot {
        internal VisioInspectionShapeSnapshot(
            string id,
            string? name,
            string? nameU,
            string? type,
            string? masterId,
            string? masterNameU,
            string? masterShapeId,
            string? parentId,
            string? text,
            double pinX,
            double pinY,
            double width,
            double height,
            double angle,
            string lineColor,
            string fillColor,
            int linePattern,
            int fillPattern,
            double lineWeight,
            bool isContainer,
            bool isCallout,
            bool isBackgroundSurface,
            bool isDiagramAdornment,
            string? calloutTargetId,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeData,
            IReadOnlyList<VisioInspectionUserCellSnapshot> userCells,
            IReadOnlyDictionary<string, string> data,
            IReadOnlyList<VisioInspectionConnectionPointSnapshot> connectionPoints,
            IReadOnlyList<string> childIds) {
            Id = id;
            Name = name;
            NameU = nameU;
            Type = type;
            MasterId = masterId;
            MasterNameU = masterNameU;
            MasterShapeId = masterShapeId;
            ParentId = parentId;
            Text = text;
            PinX = pinX;
            PinY = pinY;
            Width = width;
            Height = height;
            Angle = angle;
            LineColor = lineColor;
            FillColor = fillColor;
            LinePattern = linePattern;
            FillPattern = fillPattern;
            LineWeight = lineWeight;
            IsContainer = isContainer;
            IsCallout = isCallout;
            IsBackgroundSurface = isBackgroundSurface;
            IsDiagramAdornment = isDiagramAdornment;
            CalloutTargetId = calloutTargetId;
            Layers = layers;
            ShapeData = shapeData;
            UserCells = userCells;
            Data = data;
            ConnectionPoints = connectionPoints;
            ChildIds = childIds;
        }

        /// <summary>Shape identifier.</summary>
        public string Id { get; }

        /// <summary>Shape display name.</summary>
        public string? Name { get; }

        /// <summary>Shape universal name.</summary>
        public string? NameU { get; }

        /// <summary>Visio shape type, such as Group, when available.</summary>
        public string? Type { get; }

        /// <summary>Referenced master identifier.</summary>
        public string? MasterId { get; }

        /// <summary>Referenced master universal name.</summary>
        public string? MasterNameU { get; }

        /// <summary>Referenced master shape identifier.</summary>
        public string? MasterShapeId { get; }

        /// <summary>Parent group shape identifier.</summary>
        public string? ParentId { get; }

        /// <summary>Shape text.</summary>
        public string? Text { get; }

        /// <summary>Shape pin X coordinate.</summary>
        public double PinX { get; }

        /// <summary>Shape pin Y coordinate.</summary>
        public double PinY { get; }

        /// <summary>Shape width.</summary>
        public double Width { get; }

        /// <summary>Shape height.</summary>
        public double Height { get; }

        /// <summary>Shape rotation angle in radians.</summary>
        public double Angle { get; }

        /// <summary>Line color as a stable OfficeIMO color string.</summary>
        public string LineColor { get; }

        /// <summary>Fill color as a stable OfficeIMO color string.</summary>
        public string FillColor { get; }

        /// <summary>Visio line pattern value.</summary>
        public int LinePattern { get; }

        /// <summary>Visio fill pattern value.</summary>
        public int FillPattern { get; }

        /// <summary>Shape line weight.</summary>
        public double LineWeight { get; }

        /// <summary>Whether the shape is marked as a Visio container.</summary>
        public bool IsContainer { get; }

        /// <summary>Whether the shape is marked as an OfficeIMO callout.</summary>
        public bool IsCallout { get; }

        /// <summary>Whether the shape is marked as a background surface.</summary>
        public bool IsBackgroundSurface { get; }

        /// <summary>Whether the shape is marked as generated diagram adornment.</summary>
        public bool IsDiagramAdornment { get; }

        /// <summary>Callout target shape identifier, when present.</summary>
        public string? CalloutTargetId { get; }

        /// <summary>Layer names assigned to the shape.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape Data rows attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionShapeDataSnapshot> ShapeData { get; }

        /// <summary>User cell rows attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionUserCellSnapshot> UserCells { get; }

        /// <summary>Arbitrary data attached to the shape.</summary>
        public IReadOnlyDictionary<string, string> Data { get; }

        /// <summary>Connection points attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionConnectionPointSnapshot> ConnectionPoints { get; }

        /// <summary>Number of connection points attached to the shape.</summary>
        public int ConnectionPointCount => ConnectionPoints.Count;

        /// <summary>Child shape identifiers when this shape is a group.</summary>
        public IReadOnlyList<string> ChildIds { get; }

        internal void AppendText(StringBuilder builder, string pagePrefix) {
            string prefix = pagePrefix + ".shape[" + VisioInspectionSnapshot.EscapeKey(Id) + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".name", Name);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".nameU", NameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".type", Type);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterId", MasterId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterNameU", MasterNameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterShapeId", MasterShapeId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".parentId", ParentId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".text", Text);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".pinX", PinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".pinY", PinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".width", Width);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".height", Height);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".angle", Angle);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineColor", LineColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".fillColor", FillColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".linePattern", LinePattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".fillPattern", FillPattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineWeight", LineWeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isContainer", IsContainer);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isCallout", IsCallout);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isBackgroundSurface", IsBackgroundSurface);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isDiagramAdornment", IsDiagramAdornment);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".calloutTargetId", CalloutTargetId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".connectionPointCount", ConnectionPointCount);
            AppendConnectionPoints(builder, prefix, ConnectionPoints);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".children", string.Join(",", ChildIds));
            AppendShapeData(builder, prefix, ShapeData);
            AppendUserCells(builder, prefix, UserCells);
            AppendData(builder, prefix, Data);
        }

        internal static void AppendConnectionPoints(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionConnectionPointSnapshot> points) {
            foreach (VisioInspectionConnectionPointSnapshot point in points) {
                string pointPrefix = prefix + ".connectionPoint[" + point.Index.ToString(CultureInfo.InvariantCulture) + "]";
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".sectionIndex", point.SectionIndex);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".x", point.X);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".y", point.Y);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".dirX", point.DirX);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".dirY", point.DirY);
            }
        }

        internal static void AppendShapeData(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionShapeDataSnapshot> rows) {
            foreach (VisioInspectionShapeDataSnapshot row in rows) {
                string rowPrefix = prefix + ".shapeData[" + VisioInspectionSnapshot.EscapeKey(row.Name) + "]";
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".label", row.Label);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".value", row.Value);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".type", row.Type);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".format", row.Format);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".prompt", row.Prompt);
            }
        }

        internal static void AppendUserCells(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionUserCellSnapshot> rows) {
            foreach (VisioInspectionUserCellSnapshot row in rows) {
                string rowPrefix = prefix + ".user[" + VisioInspectionSnapshot.EscapeKey(row.Name) + "]";
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".value", row.Value);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".formula", row.Formula);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".prompt", row.Prompt);
            }
        }

        internal static void AppendData(StringBuilder builder, string prefix, IReadOnlyDictionary<string, string> data) {
            foreach (KeyValuePair<string, string> pair in data.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".data[" + VisioInspectionSnapshot.EscapeKey(pair.Key) + "]", pair.Value);
            }
        }
    }

    /// <summary>
    /// Snapshot of a Visio connector.
    /// </summary>
    public sealed class VisioInspectionConnectorSnapshot {
        internal VisioInspectionConnectorSnapshot(
            string id,
            string fromId,
            string toId,
            string kind,
            string? label,
            bool hasLabelPlacement,
            double? labelPosition,
            double? labelOffsetX,
            double? labelOffsetY,
            double? labelPinX,
            double? labelPinY,
            double? labelWidth,
            double? labelHeight,
            IReadOnlyList<VisioInspectionWaypointSnapshot> waypoints,
            string lineColor,
            int linePattern,
            double lineWeight,
            string? beginArrow,
            string? endArrow,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeData,
            IReadOnlyDictionary<string, string> data) {
            Id = id;
            FromId = fromId;
            ToId = toId;
            Kind = kind;
            Label = label;
            HasLabelPlacement = hasLabelPlacement;
            LabelPosition = labelPosition;
            LabelOffsetX = labelOffsetX;
            LabelOffsetY = labelOffsetY;
            LabelPinX = labelPinX;
            LabelPinY = labelPinY;
            LabelWidth = labelWidth;
            LabelHeight = labelHeight;
            Waypoints = waypoints;
            LineColor = lineColor;
            LinePattern = linePattern;
            LineWeight = lineWeight;
            BeginArrow = beginArrow;
            EndArrow = endArrow;
            Layers = layers;
            ShapeData = shapeData;
            Data = data;
        }

        /// <summary>Connector identifier.</summary>
        public string Id { get; }

        /// <summary>Source shape identifier.</summary>
        public string FromId { get; }

        /// <summary>Target shape identifier.</summary>
        public string ToId { get; }

        /// <summary>Connector kind.</summary>
        public string Kind { get; }

        /// <summary>Connector label.</summary>
        public string? Label { get; }

        /// <summary>Whether explicit label placement exists.</summary>
        public bool HasLabelPlacement { get; }

        /// <summary>Relative label position along the connector path, when explicit placement exists.</summary>
        public double? LabelPosition { get; }

        /// <summary>Relative label X offset, when explicit placement exists.</summary>
        public double? LabelOffsetX { get; }

        /// <summary>Relative label Y offset, when explicit placement exists.</summary>
        public double? LabelOffsetY { get; }

        /// <summary>Absolute label X coordinate, when the label is pinned to the page.</summary>
        public double? LabelPinX { get; }

        /// <summary>Absolute label Y coordinate, when the label is pinned to the page.</summary>
        public double? LabelPinY { get; }

        /// <summary>Explicit label width, when explicit placement exists.</summary>
        public double? LabelWidth { get; }

        /// <summary>Explicit label height, when explicit placement exists.</summary>
        public double? LabelHeight { get; }

        /// <summary>Explicit connector waypoints.</summary>
        public IReadOnlyList<VisioInspectionWaypointSnapshot> Waypoints { get; }

        /// <summary>Line color as a stable OfficeIMO color string.</summary>
        public string LineColor { get; }

        /// <summary>Visio line pattern value.</summary>
        public int LinePattern { get; }

        /// <summary>Connector line weight.</summary>
        public double LineWeight { get; }

        /// <summary>Begin arrow value.</summary>
        public string? BeginArrow { get; }

        /// <summary>End arrow value.</summary>
        public string? EndArrow { get; }

        /// <summary>Layer names assigned to the connector.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape Data rows attached to the connector.</summary>
        public IReadOnlyList<VisioInspectionShapeDataSnapshot> ShapeData { get; }

        /// <summary>Arbitrary data attached to the connector.</summary>
        public IReadOnlyDictionary<string, string> Data { get; }

        internal void AppendText(StringBuilder builder, string pagePrefix) {
            string prefix = pagePrefix + ".connector[" + VisioInspectionSnapshot.EscapeKey(Id) + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".from", FromId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".to", ToId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".kind", Kind);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".label", Label);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".hasLabelPlacement", HasLabelPlacement);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPosition", LabelPosition);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelOffsetX", LabelOffsetX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelOffsetY", LabelOffsetY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPinX", LabelPinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPinY", LabelPinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelWidth", LabelWidth);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelHeight", LabelHeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineColor", LineColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".linePattern", LinePattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineWeight", LineWeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".beginArrow", BeginArrow);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".endArrow", EndArrow);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));

            for (int i = 0; i < Waypoints.Count; i++) {
                string waypointPrefix = prefix + ".waypoint[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                VisioInspectionSnapshot.AppendLine(builder, waypointPrefix + ".x", Waypoints[i].X);
                VisioInspectionSnapshot.AppendLine(builder, waypointPrefix + ".y", Waypoints[i].Y);
            }

            VisioInspectionShapeSnapshot.AppendShapeData(builder, prefix, ShapeData);
            VisioInspectionShapeSnapshot.AppendData(builder, prefix, Data);
        }
    }

    /// <summary>
    /// Snapshot of one connector waypoint.
    /// </summary>
    public sealed class VisioInspectionWaypointSnapshot {
        internal VisioInspectionWaypointSnapshot(double x, double y) {
            X = x;
            Y = y;
        }

        /// <summary>Waypoint X coordinate.</summary>
        public double X { get; }

        /// <summary>Waypoint Y coordinate.</summary>
        public double Y { get; }
    }

    /// <summary>
    /// Snapshot of a Shape Data row.
    /// </summary>
    public sealed class VisioInspectionShapeDataSnapshot {
        internal VisioInspectionShapeDataSnapshot(string name, string? label, string? value, string? type, string? format, string? prompt) {
            Name = name;
            Label = label;
            Value = value;
            Type = type;
            Format = format;
            Prompt = prompt;
        }

        /// <summary>Shape Data row name.</summary>
        public string Name { get; }

        /// <summary>Shape Data display label.</summary>
        public string? Label { get; }

        /// <summary>Shape Data value.</summary>
        public string? Value { get; }

        /// <summary>Shape Data type.</summary>
        public string? Type { get; }

        /// <summary>Shape Data format string.</summary>
        public string? Format { get; }

        /// <summary>Shape Data prompt.</summary>
        public string? Prompt { get; }
    }

    /// <summary>
    /// Snapshot of a User cell row.
    /// </summary>
    public sealed class VisioInspectionUserCellSnapshot {
        internal VisioInspectionUserCellSnapshot(string name, string? value, string? formula, string? prompt) {
            Name = name;
            Value = value;
            Formula = formula;
            Prompt = prompt;
        }

        /// <summary>User cell row name.</summary>
        public string Name { get; }

        /// <summary>User cell value.</summary>
        public string? Value { get; }

        /// <summary>User cell formula.</summary>
        public string? Formula { get; }

        /// <summary>User cell prompt.</summary>
        public string? Prompt { get; }
    }

    /// <summary>
    /// Snapshot of one Visio shape connection point.
    /// </summary>
    public sealed class VisioInspectionConnectionPointSnapshot {
        internal VisioInspectionConnectionPointSnapshot(int index, int? sectionIndex, double x, double y, double dirX, double dirY) {
            Index = index;
            SectionIndex = sectionIndex;
            X = x;
            Y = y;
            DirX = dirX;
            DirY = dirY;
        }

        /// <summary>Zero-based position in the shape connection point collection.</summary>
        public int Index { get; }

        /// <summary>Original Visio Connection section row index, when loaded or assigned.</summary>
        public int? SectionIndex { get; }

        /// <summary>X coordinate relative to the shape.</summary>
        public double X { get; }

        /// <summary>Y coordinate relative to the shape.</summary>
        public double Y { get; }

        /// <summary>Directional X component.</summary>
        public double DirX { get; }

        /// <summary>Directional Y component.</summary>
        public double DirY { get; }
    }

    /// <summary>
    /// Line-oriented difference between two inspection snapshots.
    /// </summary>
    public sealed class VisioInspectionDiff {
        private VisioInspectionDiff(IReadOnlyList<VisioInspectionDifference> differences) {
            Differences = differences;
        }

        /// <summary>Snapshot differences.</summary>
        public IReadOnlyList<VisioInspectionDifference> Differences { get; }

        /// <summary>Whether any snapshot line changed.</summary>
        public bool HasDifferences => Differences.Count > 0;

        /// <summary>Compares two snapshots.</summary>
        public static VisioInspectionDiff Compare(VisioInspectionSnapshot expected, VisioInspectionSnapshot actual) {
            if (expected == null) {
                throw new ArgumentNullException(nameof(expected));
            }

            if (actual == null) {
                throw new ArgumentNullException(nameof(actual));
            }

            SortedDictionary<string, string> expectedLines = ToLineMap(expected.ToText());
            SortedDictionary<string, string> actualLines = ToLineMap(actual.ToText());
            SortedSet<string> keys = new(expectedLines.Keys, StringComparer.Ordinal);
            keys.UnionWith(actualLines.Keys);

            List<VisioInspectionDifference> differences = new();
            foreach (string key in keys) {
                bool hasExpected = expectedLines.TryGetValue(key, out string? expectedValue);
                bool hasActual = actualLines.TryGetValue(key, out string? actualValue);
                if (!hasExpected && hasActual) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Added, key, null, actualValue));
                } else if (hasExpected && !hasActual) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Removed, key, expectedValue, null));
                } else if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal)) {
                    differences.Add(new VisioInspectionDifference(VisioInspectionDifferenceKind.Changed, key, expectedValue, actualValue));
                }
            }

            return new VisioInspectionDiff(differences.AsReadOnly());
        }

        /// <summary>Writes a stable text representation of the diff.</summary>
        public string ToText() {
            StringBuilder builder = new();
            foreach (VisioInspectionDifference difference in Differences) {
                builder.Append(difference.Kind);
                builder.Append(' ');
                builder.Append(difference.Path);
                builder.Append(" expected=");
                builder.Append(VisioInspectionSnapshot.FormatValue(difference.Expected));
                builder.Append(" actual=");
                builder.Append(VisioInspectionSnapshot.FormatValue(difference.Actual));
                builder.AppendLine();
            }

            return builder.ToString();
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }

        private static SortedDictionary<string, string> ToLineMap(string text) {
            SortedDictionary<string, string> map = new(StringComparer.Ordinal);
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string line in lines) {
                if (line.Length == 0) {
                    continue;
                }

                int separator = line.IndexOf('=');
                string key = separator >= 0 ? line.Substring(0, separator) : line;
                string value = separator >= 0 ? line.Substring(separator + 1) : string.Empty;
                map[key] = value;
            }

            return map;
        }
    }

    /// <summary>
    /// Kind of inspection snapshot difference.
    /// </summary>
    public enum VisioInspectionDifferenceKind {
        /// <summary>A path exists only in the actual snapshot.</summary>
        Added,

        /// <summary>A path exists only in the expected snapshot.</summary>
        Removed,

        /// <summary>A path exists in both snapshots with a different value.</summary>
        Changed
    }

    /// <summary>
    /// One inspection snapshot difference.
    /// </summary>
    public sealed class VisioInspectionDifference {
        internal VisioInspectionDifference(VisioInspectionDifferenceKind kind, string path, string? expected, string? actual) {
            Kind = kind;
            Path = path;
            Expected = expected;
            Actual = actual;
        }

        /// <summary>Difference kind.</summary>
        public VisioInspectionDifferenceKind Kind { get; }

        /// <summary>Stable snapshot path that changed.</summary>
        public string Path { get; }

        /// <summary>Expected value.</summary>
        public string? Expected { get; }

        /// <summary>Actual value.</summary>
        public string? Actual { get; }
    }
}
