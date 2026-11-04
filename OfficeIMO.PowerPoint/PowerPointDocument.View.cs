using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Core.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        ///     Gets or sets whether slide view snapping uses the grid.
        /// </summary>
        public bool SnapToGrid {
            get {
                ThrowIfDisposed();
                CommonSlideViewProperties? common = GetCommonSlideViewProperties();
                return common?.SnapToGrid?.Value == true;
            }
            set {
                ThrowIfDisposed();
                CommonSlideViewProperties common = EnsureCommonSlideViewProperties();
                common.SnapToGrid = value;
            }
        }

        /// <summary>
        ///     Horizontal grid spacing.
        /// </summary>
        public OfficeLength GridSpacingX {
            get {
                ThrowIfDisposed();
                return OfficeLength.Emus(GetGridSpacing()?.Cx?.Value ?? 0L);
            }
            set {
                ThrowIfDisposed();
                GridSpacing spacing = EnsureGridSpacing();
                spacing.Cx = EnsureInt32(value.ToEmus(), nameof(GridSpacingX));
            }
        }

        /// <summary>
        ///     Vertical grid spacing.
        /// </summary>
        public OfficeLength GridSpacingY {
            get {
                ThrowIfDisposed();
                return OfficeLength.Emus(GetGridSpacing()?.Cy?.Value ?? 0L);
            }
            set {
                ThrowIfDisposed();
                GridSpacing spacing = EnsureGridSpacing();
                spacing.Cy = EnsureInt32(value.ToEmus(), nameof(GridSpacingY));
            }
        }

        /// <summary>
        ///     Sets horizontal and vertical grid spacing with explicit lengths.
        /// </summary>
        public void SetGridSpacing(OfficeLength horizontal, OfficeLength vertical) {
            ThrowIfDisposed();
            _ = EnsureInt32(horizontal.ToEmus(), nameof(horizontal));
            _ = EnsureInt32(vertical.ToEmus(), nameof(vertical));
            GridSpacingX = horizontal;
            GridSpacingY = vertical;
        }




        /// <summary>
        ///     Returns the current guides defined for slide view.
        /// </summary>
        public IReadOnlyList<PowerPointGuideInfo> GetGuides() {
            ThrowIfDisposed();
            CommonSlideViewProperties? common = GetCommonSlideViewProperties();
            GuideList? guideList = common?.GuideList;
            if (guideList == null) {
                return Array.Empty<PowerPointGuideInfo>();
            }

            List<PowerPointGuideInfo> guides = new();
            foreach (Guide guide in guideList.Elements<Guide>()) {
                DirectionValues? direction = guide.Orientation?.Value;
                PowerPointGuideOrientation orientation = direction == DirectionValues.Vertical
                    ? PowerPointGuideOrientation.Vertical
                    : PowerPointGuideOrientation.Horizontal;
                guides.Add(new PowerPointGuideInfo(orientation, OfficeLength.Emus(guide.Position?.Value ?? 0)));
            }

            return guides;
        }

        /// <summary>
        ///     Clears all guides from slide view.
        /// </summary>
        public void ClearGuides() {
            ThrowIfDisposed();
            CommonSlideViewProperties? common = GetCommonSlideViewProperties();
            common?.GuideList?.RemoveAllChildren<Guide>();
        }

        /// <summary>
        ///     Sets the guide list to the provided collection.
        /// </summary>
        public void SetGuides(IEnumerable<PowerPointGuideInfo> guides) {
            ThrowIfDisposed();
            if (guides == null) {
                throw new ArgumentNullException(nameof(guides));
            }

            CommonSlideViewProperties common = EnsureCommonSlideViewProperties();
            GuideList guideList = common.GuideList ??= new GuideList();
            guideList.RemoveAllChildren<Guide>();

            foreach (PowerPointGuideInfo guide in guides) {
                guideList.Append(CreateGuide(guide));
            }
        }

        /// <summary>
        ///     Adds a guide to slide view.
        /// </summary>
        public PowerPointGuideInfo AddGuide(PowerPointGuideOrientation orientation, OfficeLength position) {
            ThrowIfDisposed();
            CommonSlideViewProperties common = EnsureCommonSlideViewProperties();
            GuideList guideList = common.GuideList ??= new GuideList();
            var guide = new PowerPointGuideInfo(orientation, position);
            guideList.Append(CreateGuide(guide));
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


        private static Guide CreateGuide(PowerPointGuideInfo guide) {
            DirectionValues orientation = guide.Orientation == PowerPointGuideOrientation.Vertical
                ? DirectionValues.Vertical
                : DirectionValues.Horizontal;
            return new Guide {
                Orientation = orientation,
                Position = EnsureInt32(guide.Position.ToEmus(), nameof(guide.Position))
            };
        }

        private static int EnsureInt32(long value, string paramName) {
            if (value < int.MinValue || value > int.MaxValue) {
                throw new ArgumentOutOfRangeException(paramName,
                    $"Value must be between {int.MinValue} and {int.MaxValue}.");
            }
            return (int)value;
        }

        private ViewProperties EnsureViewProperties() {
            ViewPropertiesPart viewPart = _presentationPart.ViewPropertiesPart
                ?? _presentationPart.AddNewPart<ViewPropertiesPart>();
            viewPart.ViewProperties ??= new ViewProperties();
            return viewPart.ViewProperties;
        }

        private GridSpacing? GetGridSpacing() {
            return _presentationPart.ViewPropertiesPart?.ViewProperties?.GridSpacing;
        }

        private GridSpacing EnsureGridSpacing() {
            ViewProperties viewProperties = EnsureViewProperties();
            GridSpacing spacing = viewProperties.GridSpacing ??= new GridSpacing();
            return spacing;
        }

        private CommonSlideViewProperties? GetCommonSlideViewProperties() {
            return _presentationPart.ViewPropertiesPart?
                .ViewProperties?
                .GetFirstChild<SlideViewProperties>()?
                .GetFirstChild<CommonSlideViewProperties>();
        }

        private CommonSlideViewProperties EnsureCommonSlideViewProperties() {
            ViewProperties viewProperties = EnsureViewProperties();
            SlideViewProperties slideView = viewProperties.GetFirstChild<SlideViewProperties>()
                ?? viewProperties.AppendChild(new SlideViewProperties());
            CommonSlideViewProperties common = slideView.GetFirstChild<CommonSlideViewProperties>()
                ?? slideView.AppendChild(new CommonSlideViewProperties());
            return common;
        }

    }
}
