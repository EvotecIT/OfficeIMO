using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
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
        ///     Horizontal grid spacing in EMUs.
        /// </summary>
        public long GridSpacingXEmus {
            get {
                ThrowIfDisposed();
                return GetGridSpacing()?.Cx?.Value ?? 0L;
            }
            set {
                ThrowIfDisposed();
                GridSpacing spacing = EnsureGridSpacing();
                spacing.Cx = EnsureInt32(value, nameof(GridSpacingXEmus));
            }
        }

        /// <summary>
        ///     Vertical grid spacing in EMUs.
        /// </summary>
        public long GridSpacingYEmus {
            get {
                ThrowIfDisposed();
                return GetGridSpacing()?.Cy?.Value ?? 0L;
            }
            set {
                ThrowIfDisposed();
                GridSpacing spacing = EnsureGridSpacing();
                spacing.Cy = EnsureInt32(value, nameof(GridSpacingYEmus));
            }
        }

        /// <summary>
        ///     Sets grid spacing in EMUs.
        /// </summary>
        public void SetGridSpacing(long xEmus, long yEmus) {
            ThrowIfDisposed();
            GridSpacingXEmus = xEmus;
            GridSpacingYEmus = yEmus;
        }

        /// <summary>
        ///     Sets grid spacing in centimeters.
        /// </summary>
        public void SetGridSpacingCm(double xCm, double yCm) {
            SetGridSpacing(PowerPointUnits.FromCentimeters(xCm), PowerPointUnits.FromCentimeters(yCm));
        }

        /// <summary>
        ///     Sets grid spacing in inches.
        /// </summary>
        public void SetGridSpacingInches(double xInches, double yInches) {
            SetGridSpacing(PowerPointUnits.FromInches(xInches), PowerPointUnits.FromInches(yInches));
        }

        /// <summary>
        ///     Sets grid spacing in points.
        /// </summary>
        public void SetGridSpacingPoints(double xPoints, double yPoints) {
            SetGridSpacing(PowerPointUnits.FromPoints(xPoints), PowerPointUnits.FromPoints(yPoints));
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
                guides.Add(new PowerPointGuideInfo(orientation, guide.Position?.Value ?? 0));
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
        public void AddGuide(PowerPointGuideOrientation orientation, long positionEmus) {
            ThrowIfDisposed();
            CommonSlideViewProperties common = EnsureCommonSlideViewProperties();
            GuideList guideList = common.GuideList ??= new GuideList();
            guideList.Append(CreateGuide(new PowerPointGuideInfo(orientation, positionEmus)));
        }

        /// <summary>
        ///     Adds a guide using centimeter measurements.
        /// </summary>
        public void AddGuideCm(PowerPointGuideOrientation orientation, double positionCm) {
            AddGuide(orientation, PowerPointUnits.FromCentimeters(positionCm));
        }

        /// <summary>
        ///     Adds a guide using inch measurements.
        /// </summary>
        public void AddGuideInches(PowerPointGuideOrientation orientation, double positionInches) {
            AddGuide(orientation, PowerPointUnits.FromInches(positionInches));
        }

        /// <summary>
        ///     Adds a guide using point measurements.
        /// </summary>
        public void AddGuidePoints(PowerPointGuideOrientation orientation, double positionPoints) {
            AddGuide(orientation, PowerPointUnits.FromPoints(positionPoints));
        }

        /// <summary>
        ///     Adds vertical column guides based on a grid definition.
        /// </summary>
        public void AddColumnGuides(int columnCount, long marginEmus, long gutterEmus, bool includeOuterEdges = true,
            bool clearExisting = false) {
            ThrowIfDisposed();
            if (clearExisting) {
                ClearGuides();
            }

            PowerPointLayoutBox[] columns = SlideSize.GetColumns(columnCount, marginEmus, gutterEmus);
            var positions = new SortedSet<long>();
            foreach (PowerPointLayoutBox column in columns) {
                positions.Add(column.Left);
                positions.Add(column.Right);
            }

            if (!includeOuterEdges && positions.Count >= 2) {
                positions.Remove(positions.Min);
                positions.Remove(positions.Max);
            }

            foreach (long position in positions) {
                AddGuide(PowerPointGuideOrientation.Vertical, position);
            }
        }

        /// <summary>
        ///     Adds vertical column guides using centimeters.
        /// </summary>
        public void AddColumnGuidesCm(int columnCount, double marginCm, double gutterCm, bool includeOuterEdges = true,
            bool clearExisting = false) {
            AddColumnGuides(columnCount,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterCm),
                includeOuterEdges,
                clearExisting);
        }

        /// <summary>
        ///     Adds horizontal row guides based on a grid definition.
        /// </summary>
        public void AddRowGuides(int rowCount, long marginEmus, long gutterEmus, bool includeOuterEdges = true,
            bool clearExisting = false) {
            ThrowIfDisposed();
            if (clearExisting) {
                ClearGuides();
            }

            PowerPointLayoutBox[] rows = SlideSize.GetRows(rowCount, marginEmus, gutterEmus);
            var positions = new SortedSet<long>();
            foreach (PowerPointLayoutBox row in rows) {
                positions.Add(row.Top);
                positions.Add(row.Bottom);
            }

            if (!includeOuterEdges && positions.Count >= 2) {
                positions.Remove(positions.Min);
                positions.Remove(positions.Max);
            }

            foreach (long position in positions) {
                AddGuide(PowerPointGuideOrientation.Horizontal, position);
            }
        }

        /// <summary>
        ///     Adds horizontal row guides using centimeters.
        /// </summary>
        public void AddRowGuidesCm(int rowCount, double marginCm, double gutterCm, bool includeOuterEdges = true,
            bool clearExisting = false) {
            AddRowGuides(rowCount,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterCm),
                includeOuterEdges,
                clearExisting);
        }

        private static Guide CreateGuide(PowerPointGuideInfo guide) {
            DirectionValues orientation = guide.Orientation == PowerPointGuideOrientation.Vertical
                ? DirectionValues.Vertical
                : DirectionValues.Horizontal;
            return new Guide {
                Orientation = orientation,
                Position = EnsureInt32(guide.PositionEmus, nameof(guide.PositionEmus))
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
