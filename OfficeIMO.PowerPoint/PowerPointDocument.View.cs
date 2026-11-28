using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int ViewPml = OfficeXmlNamespaces.PresentationML;
        private const string ViewRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "viewProps";
        private const string ViewContentType = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
        private static class ViewChildOrder {
            internal static readonly OfficeXmlChildOrder Root = new(
                new[] { ViewPml, ViewPml, ViewPml, ViewPml, ViewPml, ViewPml, ViewPml, ViewPml },
                new[] { "normalViewPr", "slideViewPr", "outlineViewPr", "notesTextViewPr", "sorterViewPr", "notesViewPr", "gridSpacing", "extLst" });
            internal static readonly OfficeXmlChildOrder Slide = new(
                new[] { ViewPml, ViewPml }, new[] { "cSldViewPr", "extLst" });
            internal static readonly OfficeXmlChildOrder Common = new(
                new[] { ViewPml, ViewPml, ViewPml }, new[] { "cViewPr", "guideLst", "extLst" });
        }

        /// <summary>
        ///     Gets or sets whether slide view snapping uses the grid.
        /// </summary>
        public bool SnapToGrid {
            get {
                ThrowIfDisposed();
                string? value = CommonView(ViewPart.Root).Attribute("snapToGrid");
                return value != null && XmlConvert.ToBoolean(value);
            }
            set {
                ThrowIfDisposed();
                PowerPointPartRef part = EnsureViewPart();
                EnsureCommonView(part.Root).SetBooleanAttribute("snapToGrid", value);
                part.MarkEdited();
            }
        }

        /// <summary>
        ///     Horizontal grid spacing.
        /// </summary>
        public OfficeLength GridSpacingX {
            get {
                ThrowIfDisposed();
                return ReadGridSpacing("cx");
            }
            set {
                ThrowIfDisposed();
                int native = EnsureInt32(value.ToEmus(), nameof(GridSpacingX));
                PowerPointPartRef part = EnsureViewPart();
                part.Root.GetOrAddChild(ViewPml, "gridSpacing", ViewChildOrder.Root).SetInt32Attribute("cx", native);
                part.MarkEdited();
            }
        }

        /// <summary>
        ///     Vertical grid spacing.
        /// </summary>
        public OfficeLength GridSpacingY {
            get {
                ThrowIfDisposed();
                return ReadGridSpacing("cy");
            }
            set {
                ThrowIfDisposed();
                int native = EnsureInt32(value.ToEmus(), nameof(GridSpacingY));
                PowerPointPartRef part = EnsureViewPart();
                part.Root.GetOrAddChild(ViewPml, "gridSpacing", ViewChildOrder.Root).SetInt32Attribute("cy", native);
                part.MarkEdited();
            }
        }

        /// <summary>
        ///     Sets horizontal and vertical grid spacing with explicit lengths.
        /// </summary>
        public void SetGridSpacing(OfficeLength horizontal, OfficeLength vertical) {
            ThrowIfDisposed();
            int x = EnsureInt32(horizontal.ToEmus(), nameof(horizontal));
            int y = EnsureInt32(vertical.ToEmus(), nameof(vertical));
            PowerPointPartRef part = EnsureViewPart();
            OfficeXmlNode spacing = part.Root.GetOrAddChild(ViewPml, "gridSpacing", ViewChildOrder.Root);
            spacing.SetInt32Attribute("cx", x);
            spacing.SetInt32Attribute("cy", y);
            part.MarkEdited();
        }




        /// <summary>
        ///     Returns the current guides defined for slide view.
        /// </summary>
        public IReadOnlyList<PowerPointGuideInfo> GetGuides() {
            ThrowIfDisposed();
            var guides = new List<PowerPointGuideInfo>();
            OfficeXmlNode list = CommonView(ViewPart.Root).Child(ViewPml, "guideLst");
            foreach (OfficeXmlNode guide in list.Children(ViewPml, "guide")) {
                PowerPointGuideOrientation orientation = guide.Attribute("orient") switch {
                    "vert" => PowerPointGuideOrientation.Vertical,
                    "horz" or null => PowerPointGuideOrientation.Horizontal,
                    _ => throw new FormatException("The guide orientation is invalid.")
                };
                string? position = guide.Attribute("pos");
                guides.Add(new PowerPointGuideInfo(orientation, OfficeLength.Emus(position == null ? 0 : XmlConvert.ToInt32(position))));
            }

            return guides;
        }

        /// <summary>
        ///     Clears all guides from slide view.
        /// </summary>
        public void ClearGuides() {
            ThrowIfDisposed();
            PowerPointPartRef part = ViewPart;
            OfficeXmlNode list = CommonView(part.Root).Child(ViewPml, "guideLst");
            if (list.IsNone) return;
            list.RemoveChildren(ViewPml, "guide");
            part.MarkEdited();
        }

        /// <summary>
        ///     Sets the guide list to the provided collection.
        /// </summary>
        public void SetGuides(IEnumerable<PowerPointGuideInfo> guides) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(guides);
            var validated = guides.Select(guide => (guide.Orientation,
                Position: EnsureInt32(guide.Position.ToEmus(), nameof(guide.Position)))).ToArray();

            PowerPointPartRef part = EnsureViewPart();
            OfficeXmlNode list = EnsureCommonView(part.Root).GetOrAddChild(ViewPml, "guideLst", ViewChildOrder.Common);
            list.RemoveChildren(ViewPml, "guide");
            part.MarkEdited();

            foreach ((PowerPointGuideOrientation orientation, int position) in validated) AppendGuide(list, orientation, position);
        }

        /// <summary>
        ///     Adds a guide to slide view.
        /// </summary>
        public PowerPointGuideInfo AddGuide(PowerPointGuideOrientation orientation, OfficeLength position) {
            ThrowIfDisposed();
            int nativePosition = EnsureInt32(position.ToEmus(), nameof(PowerPointGuideInfo.Position));
            PowerPointPartRef part = EnsureViewPart();
            OfficeXmlNode list = EnsureCommonView(part.Root).GetOrAddChild(ViewPml, "guideLst", ViewChildOrder.Common);
            var guide = new PowerPointGuideInfo(orientation, position);
            AppendGuide(list, orientation, nativePosition);
            part.MarkEdited();
            return guide;
        }




        /// <summary>
        ///     Adds vertical column guides based on a grid definition.
        /// </summary>
        public IReadOnlyList<PowerPointGuideInfo> AddColumnGuides(int columnCount, OfficeLength margin, OfficeLength gutter, bool includeOuterEdges = true,
            bool clearExisting = false) {
            ThrowIfDisposed();
            if (clearExisting) {
                ClearGuides();
            }

            PowerPointLayoutBox[] columns = SlideSize.GetColumns(columnCount,margin,gutter);
            var positions = new SortedSet<long>();
            foreach (PowerPointLayoutBox column in columns) {
                positions.Add(column.Left.ToEmus());
                positions.Add(column.Right.ToEmus());
            }

            if (!includeOuterEdges && positions.Count >= 2) {
                positions.Remove(positions.Min);
                positions.Remove(positions.Max);
            }

            var added = new List<PowerPointGuideInfo>(positions.Count);
            foreach (long position in positions) {
                added.Add(AddGuide(PowerPointGuideOrientation.Vertical, OfficeLength.Emus(position)));
            }
            return added;
        }


        /// <summary>
        ///     Adds horizontal row guides based on a grid definition.
        /// </summary>
        public IReadOnlyList<PowerPointGuideInfo> AddRowGuides(int rowCount, OfficeLength margin, OfficeLength gutter, bool includeOuterEdges = true,
            bool clearExisting = false) {
            ThrowIfDisposed();
            if (clearExisting) {
                ClearGuides();
            }

            PowerPointLayoutBox[] rows = SlideSize.GetRows(rowCount,margin,gutter);
            var positions = new SortedSet<long>();
            foreach (PowerPointLayoutBox row in rows) {
                positions.Add(row.Top.ToEmus());
                positions.Add(row.Bottom.ToEmus());
            }

            if (!includeOuterEdges && positions.Count >= 2) {
                positions.Remove(positions.Min);
                positions.Remove(positions.Max);
            }

            var added = new List<PowerPointGuideInfo>(positions.Count);
            foreach (long position in positions) {
                added.Add(AddGuide(PowerPointGuideOrientation.Horizontal, OfficeLength.Emus(position)));
            }
            return added;
        }


        private static void AppendGuide(OfficeXmlNode list, PowerPointGuideOrientation orientation, int position) {
            OfficeXmlNode element = list.Append(list.CreateElement(ViewPml, "guide"));
            element.SetAttribute("orient", orientation == PowerPointGuideOrientation.Vertical ? "vert" : "horz");
            element.SetInt32Attribute("pos", position);
        }

        private static int EnsureInt32(long value, string paramName) {
            if (value < int.MinValue || value > int.MaxValue) {
                throw new ArgumentOutOfRangeException(paramName,
                    $"Value must be between {int.MinValue} and {int.MaxValue}.");
            }
            return (int)value;
        }

        private PowerPointPartRef ViewPart {
            get {
                PowerPointPartRef presentation = PresentationPartRef;
                PowerPointPartRef view = presentation.RelatedPart(ViewRelationshipType);
                if (presentation.IsSdk || ReadablePackage is not { } package) return view;
                bool supported = view.IsNone
                    ? !HasRelationshipOfType(package, presentation.Name, ViewRelationshipType)
                    : string.Equals(package.GetContentType(view.Name), ViewContentType, StringComparison.OrdinalIgnoreCase)
                        && view.Root.Is(ViewPml, "viewPr");
                if (supported) return view.Current;
                OpenSdkForPackageCompatibility();
                return PresentationPartRef.RelatedPart(ViewRelationshipType);
            }
        }

        private PowerPointPartRef EnsureViewPart() {
            PowerPointPartRef existing = ViewPart;
            if (existing.Root.Exists) return existing.Current;

            PowerPointPartRef presentation = PresentationPartRef;
            if (!presentation.IsSdk && EditablePackage is { } package && existing.IsNone) {
                string path = UniquePartName(package, ViewContentType,
                    presentation.Name.Substring(0, presentation.Name.LastIndexOf('/') + 1) + "viewProps.xml");
                package.CreateXmlPart(path, ViewContentType, ViewPml, "viewPr", 16);
                try {
                    package.AddRelationship(presentation.Name, ViewRelationshipType, path, id: NewRelationshipId());
                } catch {
                    package.DeletePart(path);
                    throw;
                }
                InvalidatePackageRelationships();
                MarkPackageEdited();
                return PowerPointPartRef.Package(this, path);
            }

            return EnsureSdkViewPart();
        }

        private OfficeLength ReadGridSpacing(string name) {
            string? value = ViewPart.Root.Child(ViewPml, "gridSpacing").Attribute(name);
            return OfficeLength.Emus(value == null ? 0L : XmlConvert.ToInt64(value));
        }

        private static OfficeXmlNode CommonView(OfficeXmlNode root) =>
            root.Child(ViewPml, "slideViewPr").Child(ViewPml, "cSldViewPr");

        private static OfficeXmlNode EnsureCommonView(OfficeXmlNode root) =>
            root.GetOrAddChild(ViewPml, "slideViewPr", ViewChildOrder.Root).GetOrAddChild(ViewPml, "cSldViewPr", ViewChildOrder.Slide);
    }
}
