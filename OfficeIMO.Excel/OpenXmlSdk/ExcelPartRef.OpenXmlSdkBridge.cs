using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel.OpenXml {
    /// <summary>The kinds of part drawing, chart and image code adds to a package.</summary>
    internal enum ExcelPartKind {
        Other,
        Workbook,
        Worksheet,
        Chartsheet,
        Drawings,
        Chart,
        ExtendedChart,
        ChartStyle,
        ChartColorStyle,
        ChartDrawing,
        Image,
        EmbeddedPackage,
        Theme,
        ThemeOverride,
        Pivot,
        Table
    }

    /// <summary>
    /// A part of the workbook package that drawing, chart and image code reads and edits: its root element, its
    /// relationships, the parts it relates to and its bytes. Code outside this file names no SDK type for parts. While
    /// the workbook is held by the Open XML SDK the part is an SDK part, passed as <see cref="object"/>; this is the only
    /// place drawing code touches SDK parts, and it goes when charts are written to OfficeIMO's own package.
    /// </summary>
    internal readonly struct ExcelPartRef : IEquatable<ExcelPartRef> {
        private readonly OpenXmlPart? _part;

        private ExcelPartRef(OpenXmlPart part) => _part = part;

        /// <summary>An SDK part, passed as <see cref="object"/> so callers name no SDK type.</summary>
        internal static ExcelPartRef FromSdk(object? part) => part is SpreadsheetDocument workbook ? FromSdk(workbook.WorkbookPart) : part is OpenXmlPart sdk ? new ExcelPartRef(sdk) : default;

        internal static ExcelPartRef None => default;

        internal bool IsNone => _part == null;
        internal bool Exists => _part != null;

        /// <summary>The SDK part, or null.</summary>
        internal object? SdkPart => _part;

        // ---- Parts of the workbook --------------------------------------------------------------------------------

        /// <summary>The worksheet part of a sheet.</summary>
        internal static ExcelPartRef Worksheet(ExcelSheet sheet) => new(sheet.WorksheetPart);

        /// <summary>The worksheet part of a sheet, without the preparation <see cref="Worksheet"/> runs first.</summary>
        internal static ExcelPartRef WorksheetForMetadata(ExcelSheet sheet) => new(sheet.DeferredMetadataWorksheetPart);

        /// <summary>The workbook part.</summary>
        internal static ExcelPartRef Workbook(ExcelDocument document) => new(document.WorkbookPartRoot);

        /// <summary>The worksheet parts of the workbook, with their relationship ids.</summary>
        internal static IEnumerable<ExcelPartRef> WorksheetParts(ExcelDocument document) {
            foreach (WorksheetPart part in document.WorkbookPartRoot.WorksheetParts) yield return new ExcelPartRef(part);
        }

        /// <summary>The chartsheet parts of the workbook.</summary>
        internal static IEnumerable<ExcelPartRef> ChartsheetParts(ExcelDocument document) {
            foreach (ChartsheetPart part in document.WorkbookPartRoot.ChartsheetParts) yield return new ExcelPartRef(part);
        }

        /// <summary>Whether the sheet's element is still in the workbook's sheet list.</summary>
        internal static bool IsSheetAttached(ExcelSheet sheet) => sheet.SheetElement.Parent != null;

        /// <summary>The sheet id the workbook gives the sheet, or null.</summary>
        internal static uint? SheetId(ExcelSheet sheet) => sheet.SheetElement.SheetId?.Value;

        // ---- The part ---------------------------------------------------------------------------------------------

        /// <summary>The part name, such as <c>/xl/charts/chart1.xml</c>.</summary>
        internal string Name => _part?.Uri.OriginalString ?? string.Empty;

        /// <summary>The part URI, the part's identity across relationships.</summary>
        internal Uri? Uri => _part?.Uri;

        internal string ContentType => _part?.ContentType ?? string.Empty;
        internal string? RelationshipType => _part?.RelationshipType;
        internal bool IsReadOnly => _part?.OpenXmlPackage.FileOpenAccess == FileAccess.Read;
        internal bool ProcessesMarkupCompatibility => _part?.OpenXmlPackage.MarkupCompatibilityProcessSettings.ProcessMode != MarkupCompatibilityProcessMode.NoProcess;
        internal bool IsRootLoaded => _part?.IsRootElementLoaded == true;
        internal ExcelPartKind Kind => _part switch {
            WorkbookPart => ExcelPartKind.Workbook,
            WorksheetPart => ExcelPartKind.Worksheet,
            ChartsheetPart => ExcelPartKind.Chartsheet,
            DrawingsPart => ExcelPartKind.Drawings,
            ChartPart => ExcelPartKind.Chart,
            ExtendedChartPart => ExcelPartKind.ExtendedChart,
            ChartStylePart => ExcelPartKind.ChartStyle,
            ChartColorStylePart => ExcelPartKind.ChartColorStyle,
            ChartDrawingPart => ExcelPartKind.ChartDrawing,
            ImagePart => ExcelPartKind.Image,
            EmbeddedPackagePart => ExcelPartKind.EmbeddedPackage,
            ThemePart => ExcelPartKind.Theme,
            ThemeOverridePart => ExcelPartKind.ThemeOverride,
            PivotTablePart => ExcelPartKind.Pivot,
            TableDefinitionPart => ExcelPartKind.Table,
            _ => ExcelPartKind.Other
        };

        /// <summary>The root element, read from the part on first use; none for a part without XML content.</summary>
        internal OfficeXmlNode Root => OfficeXmlNode.FromSdk(_part?.RootElement);

        /// <summary>Writes the root element back to the part.</summary>
        internal void SaveRoot() {
            if (_part?.RootElement is OpenXmlPartRootElement root) root.Save();
        }

        /// <summary>Replaces the part's root element with a new one of this name, written back to the part.</summary>
        internal OfficeXmlNode SetRootMarkup(byte[] markup) {
            if (_part == null) throw new InvalidOperationException("There is no part.");
            using (var stream = new MemoryStream(markup, writable: false)) _part.FeedData(stream);
            return Root;
        }

        /// <summary>Opens the part's bytes for reading.</summary>
        internal Stream OpenRead() => _part!.GetStream(FileMode.Open, FileAccess.Read);

        /// <summary>The part's bytes.</summary>
        internal byte[] ReadAllBytes() {
            using Stream stream = _part!.GetStream(FileMode.Open, FileAccess.Read);
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return copy.ToArray();
        }

        /// <summary>Replaces the part's bytes.</summary>
        internal void Feed(byte[] data) {
            using var stream = new MemoryStream(data, writable: false);
            _part!.FeedData(stream);
        }

        // ---- Relationships ----------------------------------------------------------------------------------------

        /// <summary>The internal part a relationship id targets, or none.</summary>
        internal ExcelPartRef GetPartById(string? id) =>
            _part != null && !string.IsNullOrEmpty(id) && _part.TryGetPartById(id!, out OpenXmlPart? related) ? new ExcelPartRef(related!) : default;

        /// <summary>The relationship id under which this part relates to <paramref name="child"/>, or null.</summary>
        internal string? GetIdOfPart(ExcelPartRef child) {
            if (_part == null || child._part == null) return null;
            foreach (IdPartPair pair in _part.Parts) {
                if (ReferenceEquals(pair.OpenXmlPart, child._part)) return pair.RelationshipId;
            }
            return null;
        }

        /// <summary>The parts this part relates to, with their relationship ids, in relationship order.</summary>
        internal IEnumerable<(string Id, ExcelPartRef Part)> Parts {
            get {
                if (_part == null) yield break;
                foreach (IdPartPair pair in _part.Parts) yield return (pair.RelationshipId, new ExcelPartRef(pair.OpenXmlPart));
            }
        }

        /// <summary>The related parts of one kind, in relationship order.</summary>
        internal IEnumerable<ExcelPartRef> PartsOfKind(ExcelPartKind kind) {
            if (_part == null) yield break;
            foreach (IdPartPair pair in _part.Parts) {
                var part = new ExcelPartRef(pair.OpenXmlPart);
                if (part.Kind == kind) yield return part;
            }
        }

        /// <summary>The first related part of one kind, or none.</summary>
        internal ExcelPartRef FirstPartOfKind(ExcelPartKind kind) {
            foreach (ExcelPartRef part in PartsOfKind(kind)) return part;
            return default;
        }

        /// <summary>The drawing part of a worksheet or chartsheet, or none.</summary>
        internal ExcelPartRef Drawings => _part switch {
            WorksheetPart worksheet => new ExcelPartRef(worksheet.DrawingsPart!),
            ChartsheetPart chartsheet => new ExcelPartRef(chartsheet.DrawingsPart!),
            _ => default
        };

        /// <summary>The theme part of the workbook, or none.</summary>
        internal ExcelPartRef Theme => _part is WorkbookPart workbook ? new ExcelPartRef(workbook.ThemePart!) : default;

        /// <summary>Adds a new part of a kind that has a fixed content type, related under a new or given id.</summary>
        internal ExcelPartRef AddNewPart(ExcelPartKind kind, string? id = null) {
            if (_part == null) throw new InvalidOperationException("There is no part to add a part to.");
            OpenXmlPart added = kind switch {
                ExcelPartKind.Drawings => id == null ? _part.AddNewPart<DrawingsPart>() : _part.AddNewPart<DrawingsPart>(id),
                ExcelPartKind.Chart => id == null ? _part.AddNewPart<ChartPart>() : _part.AddNewPart<ChartPart>(id),
                ExcelPartKind.ExtendedChart => id == null ? _part.AddNewPart<ExtendedChartPart>() : _part.AddNewPart<ExtendedChartPart>(id),
                ExcelPartKind.ChartStyle => id == null ? _part.AddNewPart<ChartStylePart>() : _part.AddNewPart<ChartStylePart>(id),
                ExcelPartKind.ChartColorStyle => id == null ? _part.AddNewPart<ChartColorStylePart>() : _part.AddNewPart<ChartColorStylePart>(id),
                _ => throw new NotSupportedException("Parts of kind " + kind + " are added with a content type.")
            };
            return new ExcelPartRef(added);
        }

        /// <summary>Adds an image part with the given content type to a drawing part.</summary>
        internal ExcelPartRef AddImagePart(string contentType, string? id = null) {
            switch (_part) {
                case DrawingsPart drawings:
                    return new ExcelPartRef(id == null ? drawings.AddImagePart(contentType) : drawings.AddImagePart(contentType, id));
                case ChartPart chart:
                    return new ExcelPartRef(id == null ? chart.AddImagePart(contentType) : chart.AddImagePart(contentType, id));
                default:
                    throw new NotSupportedException("Image parts are added to drawing and chart parts.");
            }
        }



        /// <summary>Removes the relationship to a related part, deleting it when nothing else relates to it.</summary>
        internal void DeletePart(ExcelPartRef child) {
            if (_part != null && child._part != null) _part.DeletePart(child._part);
        }


        /// <summary>
        /// The target of an external relationship, or of an internal hyperlink relationship, with this id; null when
        /// the part has none.
        /// </summary>
        internal Uri? GetHyperlinkUri(string id, bool externalOnly) {
            if (_part == null) return null;
            foreach (HyperlinkRelationship relationship in _part.HyperlinkRelationships) {
                if (relationship.Id == id && (!externalOnly || relationship.IsExternal)) return relationship.Uri;
            }
            return null;
        }


        /// <summary>Whether the part has a hyperlink relationship with this id.</summary>
        internal bool HasHyperlink(string id) {
            if (_part == null) return false;
            foreach (HyperlinkRelationship relationship in _part.HyperlinkRelationships) {
                if (relationship.Id == id) return true;
            }
            return false;
        }

        /// <summary>Removes the hyperlink relationship with this id.</summary>
        internal void RemoveHyperlink(string id) {
            if (_part == null) return;
            foreach (HyperlinkRelationship relationship in _part.HyperlinkRelationships) {
                if (relationship.Id == id) {
                    _part.DeleteReferenceRelationship(relationship);
                    return;
                }
            }
        }

        /// <summary>The target of an external relationship with this id, or null.</summary>
        internal Uri? GetExternalRelationshipUri(string id) {
            if (_part == null) return null;
            foreach (ExternalRelationship relationship in _part.ExternalRelationships) {
                if (relationship.Id == id) return relationship.Uri;
            }
            return null;
        }

        // ---- Chart utilities (group C) ----------------------------------------------------------------------------

        /// <summary>The workbook part of the spreadsheet package this part belongs to, or none.</summary>
        internal ExcelPartRef PackageWorkbook =>
            _part?.OpenXmlPackage is SpreadsheetDocument document && document.WorkbookPart is WorkbookPart workbook ? new ExcelPartRef(workbook) : default;
        // ---- Worksheet drawings and images (group D) -------------------------------------------------------------

        /// <summary>
        /// The sheet's <c>drawing</c> element in its worksheet, or none. The worksheet is read the way the sheet loads it,
        /// so this doesn't bypass the sheet's own loader.
        /// </summary>
        internal static OfficeXmlNode WorksheetDrawingElement(ExcelSheet sheet) =>
            OfficeXmlNode.FromSdk(sheet.WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>());

        /// <summary>Links the worksheet to its drawing part: appends <c>drawing r:id</c> to the worksheet element.</summary>
        internal static void LinkWorksheetDrawing(ExcelSheet sheet, string relationshipId) =>
            sheet.WorksheetRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = relationshipId });

        /// <summary>Writes the sheet's worksheet element back to its part.</summary>
        internal static void SaveWorksheetElement(ExcelSheet sheet) => sheet.WorksheetRoot.Save();

        /// <summary>Whether the workbook package was opened for reading only.</summary>
        internal static bool IsReadOnlyPackage(ExcelDocument document) =>
            document._spreadSheetDocument.FileOpenAccess == FileAccess.Read;

        /// <summary>The first worksheet part this part is related from, or none.</summary>
        internal ExcelPartRef ParentWorksheet {
            get {
                if (_part == null) return default;
                foreach (OpenXmlPart parent in _part.GetParentParts()) {
                    if (parent is WorksheetPart worksheet) return new ExcelPartRef(worksheet);
                }
                return default;
            }
        }

        /// <summary>The column and row geometry of this worksheet part; the default geometry for none.</summary>
        internal Utilities.ExcelWorksheetGeometryIndex CreateWorksheetGeometry() =>
            Utilities.ExcelWorksheetGeometryIndex.Create(_part as WorksheetPart);

        /// <summary>Adds an image part with this content type and file extension to a drawing part.</summary>
        internal ExcelPartRef AddImagePartWithExtension(string contentType, string extension) {
            if (_part is not DrawingsPart drawings) throw new NotSupportedException("Image parts with an extension are added to drawing parts.");
            return new ExcelPartRef(drawings.AddImagePart(new PartTypeInfo(contentType, extension)));
        }

        /// <summary>
        /// The id of an external hyperlink relationship whose target is written exactly as <paramref name="uri"/>, added
        /// when there is none.
        /// </summary>
        internal string GetOrAddExternalHyperlinkWithSameText(Uri uri) {
            foreach (HyperlinkRelationship relationship in _part!.HyperlinkRelationships) {
                if (relationship.IsExternal && string.Equals(relationship.Uri.OriginalString, uri.OriginalString, StringComparison.Ordinal)) {
                    return relationship.Id;
                }
            }
            return _part.AddHyperlinkRelationship(uri, true).Id;
        }

        /// <summary>This chart part as the shared chart reader and writer take it.</summary>
        internal IOfficeChartPart AsChartPart() =>
            _part is ChartPart chart ? new OfficeSdkChartPart(chart) : throw new InvalidOperationException("The part is not a chart part.");

        /// <summary>The column and row geometry of a worksheet part; the default geometry for none.</summary>
        internal static Utilities.ExcelWorksheetGeometryIndex CreateGeometryIndex(ExcelPartRef worksheet) =>
            Utilities.ExcelWorksheetGeometryIndex.Create(worksheet._part as WorksheetPart);

        /// <summary>The maximum digit width of a worksheet part's default font.</summary>
        internal static double GetMaximumDigitWidth(ExcelPartRef worksheet) =>
            Utilities.ExcelColumnGeometry.GetMaximumDigitWidth(worksheet._part as WorksheetPart);

        /// <summary>The hidden rows and columns of a sheet, for charts that plot visible cells only.</summary>
        internal static OfficeOpenXmlChartHiddenCells CreateHiddenCells(ExcelSheet sheet) => new(sheet.WorksheetPart.Worksheet!);

        /// <summary>The largest <c>xdr:cNvPr</c> id in a worksheet drawing, read through the SDK's typed elements.</summary>
        internal uint MaxSdkDrawingId() {
            uint max = 0;
            if ((_part as DrawingsPart)?.WorksheetDrawing is { } drawing) {
                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties properties in
                    drawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties>()) {
                    if (properties.Id?.Value is uint id && id > max) max = id;
                }
            }
            return max;
        }

        // ---- Identity ---------------------------------------------------------------------------------------------

        public bool Equals(ExcelPartRef other) => ReferenceEquals(_part, other._part);
        public override bool Equals(object? obj) => obj is ExcelPartRef other && Equals(other);
        public override int GetHashCode() => _part == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_part);
        public static bool operator ==(ExcelPartRef left, ExcelPartRef right) => left.Equals(right);
        public static bool operator !=(ExcelPartRef left, ExcelPartRef right) => !left.Equals(right);

        // ---- Chart core (group A1) ---------------------------------------------------------------------------------

        /// <summary>
        /// The chart part a relationship id targets. Throws as the SDK's lookup and cast do when the id is unknown or
        /// targets another kind of part.
        /// </summary>
        internal ExcelPartRef GetChartPartById(string id) => new((ChartPart)_part!.GetPartById(id));

        // ---- Chart image export (group B) -------------------------------------------------------------------------

        /// <summary>The styles part of the workbook, or none.</summary>
        internal ExcelPartRef WorkbookStyles => _part is WorkbookPart workbook ? new ExcelPartRef(workbook.WorkbookStylesPart!) : default;
    }
}

namespace OfficeIMO.Excel {
    // Chart utilities (group C): the SDK-typed entry the package bridge calls when it opens a package.
    internal static partial class ExcelChartAxisIdGenerator {
        /// <summary>
        /// Reads the relationships the load-time axis scan used to read, without parsing any chart, so a malformed
        /// relationship part still fails the load itself. Loading a chart root reads every part's relationships,
        /// which is why a workbook with charts reads them all here too.
        /// </summary>
        internal static void ReadChartRelationships(SpreadsheetDocument document) {
            if (document.WorkbookPart == null) return;
            if (HasWorksheetCharts(OpenXml.ExcelPartRef.FromSdk(document.WorkbookPart))) {
                _ = document.StrictRelationshipFound;
            }
        }
    }

    public partial class ExcelDocument {
        /// <summary>Snapshots live drawing identity before package replacement or a transactional mutation.</summary>
        internal IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> CaptureDrawingHandlesForReload(
            IReadOnlyDictionary<uint, WorksheetPart>? activeParts = null) {
            activeParts ??= CaptureActiveSheetPartsForReload();
            var parts = new Dictionary<uint, OpenXml.ExcelPartRef>(activeParts.Count);
            foreach (KeyValuePair<uint, WorksheetPart> pair in activeParts) parts.Add(pair.Key, OpenXml.ExcelPartRef.FromSdk(pair.Value));
            return CaptureDrawingHandles(parts);
        }
    }
}
