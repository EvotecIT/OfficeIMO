using System;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for adding semantic callouts and annotation leader lines.
    /// </summary>
    public static class VisioCalloutExtensions {
        /// <summary>
        /// Adds a callout near a target shape using an automatically assigned shape identifier.
        /// </summary>
        /// <param name="page">Page that receives the callout.</param>
        /// <param name="target">Shape being annotated.</param>
        /// <param name="text">Callout text.</param>
        /// <param name="pinX">Callout pin X coordinate.</param>
        /// <param name="pinY">Callout pin Y coordinate.</param>
        /// <param name="options">Optional callout options.</param>
        public static VisioShape AddCallout(this VisioPage page, VisioShape target, string text, double pinX, double pinY, VisioCalloutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            VisioCalloutOptions effectiveOptions = options ?? new VisioCalloutOptions();
            VisioCalloutOptions pageOptions = ConvertOptionsToPageUnits(effectiveOptions, page.DefaultUnit);
            EnsureTargetBelongsToPage(page, target);
            VisioShape callout = page.AddRectangle(pinX, pinY, effectiveOptions.Width, effectiveOptions.Height, text, page.DefaultUnit);
            ConfigureCallout(page, target, callout, pageOptions);
            return callout;
        }

        /// <summary>
        /// Adds a callout on a target side using an automatically assigned shape identifier.
        /// </summary>
        /// <param name="page">Page that receives the callout.</param>
        /// <param name="target">Shape being annotated.</param>
        /// <param name="text">Callout text.</param>
        /// <param name="placement">Side of the target where the callout should be placed.</param>
        /// <param name="gap">Distance between the target edge and callout edge, in page units.</param>
        /// <param name="options">Optional callout options.</param>
        public static VisioShape AddCallout(this VisioPage page, VisioShape target, string text, VisioSide placement, double gap = 0.35D, VisioCalloutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            VisioCalloutOptions effectiveOptions = options ?? new VisioCalloutOptions();
            VisioCalloutOptions pageOptions = ConvertOptionsToPageUnits(effectiveOptions, page.DefaultUnit);
            CalculatePin(target, pageOptions, placement, gap.ToInches(page.DefaultUnit), out double pinX, out double pinY);
            return AddCalloutInPageCoordinates(page, target, text, pinX, pinY, pageOptions);
        }

        /// <summary>
        /// Adds a callout near a target shape using an explicit shape identifier.
        /// </summary>
        /// <param name="page">Page that receives the callout.</param>
        /// <param name="target">Shape being annotated.</param>
        /// <param name="id">Callout shape identifier.</param>
        /// <param name="text">Callout text.</param>
        /// <param name="pinX">Callout pin X coordinate.</param>
        /// <param name="pinY">Callout pin Y coordinate.</param>
        /// <param name="options">Optional callout options.</param>
        public static VisioShape AddCallout(this VisioPage page, VisioShape target, string id, string text, double pinX, double pinY, VisioCalloutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Callout id cannot be empty.", nameof(id));
            }

            if (page.FindShapeById(id) != null || page.Connectors.Any(connector => string.Equals(connector.Id, id, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidOperationException("A shape or connector with the callout id already exists on the page.");
            }

            VisioCalloutOptions effectiveOptions = options ?? new VisioCalloutOptions();
            VisioCalloutOptions pageOptions = ConvertOptionsToPageUnits(effectiveOptions, page.DefaultUnit);
            EnsureTargetBelongsToPage(page, target);
            VisioShape callout = new VisioShape(id, pinX.ToInches(page.DefaultUnit), pinY.ToInches(page.DefaultUnit), pageOptions.Width, pageOptions.Height, text) {
                Name = "Callout",
                NameU = "Rectangle"
            };
            page.Shapes.Add(callout);
            ConfigureCallout(page, target, callout, pageOptions);
            return callout;
        }

        /// <summary>
        /// Adds a callout on a target side using an explicit shape identifier.
        /// </summary>
        /// <param name="page">Page that receives the callout.</param>
        /// <param name="target">Shape being annotated.</param>
        /// <param name="id">Callout shape identifier.</param>
        /// <param name="text">Callout text.</param>
        /// <param name="placement">Side of the target where the callout should be placed.</param>
        /// <param name="gap">Distance between the target edge and callout edge, in page units.</param>
        /// <param name="options">Optional callout options.</param>
        public static VisioShape AddCallout(this VisioPage page, VisioShape target, string id, string text, VisioSide placement, double gap = 0.35D, VisioCalloutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Callout id cannot be empty.", nameof(id));
            }

            VisioCalloutOptions effectiveOptions = options ?? new VisioCalloutOptions();
            VisioCalloutOptions pageOptions = ConvertOptionsToPageUnits(effectiveOptions, page.DefaultUnit);
            CalculatePin(target, pageOptions, placement, gap.ToInches(page.DefaultUnit), out double pinX, out double pinY);
            return AddCalloutInPageCoordinates(page, target, id, text, pinX, pinY, pageOptions);
        }

        private static VisioShape AddCalloutInPageCoordinates(VisioPage page, VisioShape target, string text, double pinX, double pinY, VisioCalloutOptions options) {
            VisioShape callout = page.AddRectangle(pinX, pinY, options.Width, options.Height, text, VisioMeasurementUnit.Inches);
            ConfigureCallout(page, target, callout, options);
            return callout;
        }

        private static VisioShape AddCalloutInPageCoordinates(VisioPage page, VisioShape target, string id, string text, double pinX, double pinY, VisioCalloutOptions options) {
            if (page.FindShapeById(id) != null || page.Connectors.Any(connector => string.Equals(connector.Id, id, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidOperationException("A shape or connector with the callout id already exists on the page.");
            }

            VisioShape callout = new VisioShape(id, pinX, pinY, options.Width, options.Height, text) {
                Name = "Callout",
                NameU = "Rectangle"
            };
            page.Shapes.Add(callout);
            ConfigureCallout(page, target, callout, options);
            return callout;
        }

        private static void ConfigureCallout(VisioPage page, VisioShape target, VisioShape callout, VisioCalloutOptions options) {
            EnsureTargetBelongsToPage(page, target);

            callout.Name = string.IsNullOrWhiteSpace(callout.Name) ? "Callout" : callout.Name;
            callout.NameU = string.IsNullOrWhiteSpace(callout.NameU) ? "Rectangle" : callout.NameU;
            options.GetShapeStyle().ApplyTo(callout);
            callout.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.CalloutKind, "STR", prompt: "OfficeIMO semantic kind");
            callout.SetUserCell(VisioSemanticUserCells.CalloutTargetId, target.Id, "STR", prompt: "OfficeIMO callout target shape id");

            VisioSide targetSide = options.TargetSide == VisioSide.Auto ? ChooseSide(target, callout) : options.TargetSide;
            VisioSide calloutSide = options.CalloutSide == VisioSide.Auto ? Opposite(targetSide) : options.CalloutSide;
            VisioConnector leader = page.AddConnector(callout, target, options.LeaderKind, calloutSide, targetSide);
            options.GetLeaderStyle().ApplyTo(leader);

            if (options.RouteLeader && leader.Kind == ConnectorKind.RightAngle) {
                leader.RouteOrthogonal(offset: options.RouteOffset);
            }

            callout.SetUserCell(VisioSemanticUserCells.CalloutLeaderId, leader.Id, "STR", prompt: "OfficeIMO callout leader connector id");

            if (!string.IsNullOrWhiteSpace(options.LayerName)) {
                page.AddToLayer(options.LayerName!, callout);
                page.AddToLayer(options.LayerName!, leader);
            }
        }

        private static void CalculatePin(VisioShape target, VisioCalloutOptions options, VisioSide placement, double gap, out double pinX, out double pinY) {
            if (double.IsNaN(gap) || double.IsInfinity(gap) || gap < 0D) {
                throw new ArgumentOutOfRangeException(nameof(gap), "Gap must be a finite non-negative number.");
            }

            if (double.IsNaN(options.Width) || double.IsInfinity(options.Width) || options.Width <= 0D) {
                throw new ArgumentOutOfRangeException(nameof(options.Width), "Callout width must be a finite positive number.");
            }

            if (double.IsNaN(options.Height) || double.IsInfinity(options.Height) || options.Height <= 0D) {
                throw new ArgumentOutOfRangeException(nameof(options.Height), "Callout height must be a finite positive number.");
            }

            VisioSide effectivePlacement = placement == VisioSide.Auto ? VisioSide.Right : placement;
            pinX = target.PinX;
            pinY = target.PinY;
            switch (effectivePlacement) {
                case VisioSide.Left:
                    pinX = target.PinX - (target.Width / 2D) - gap - (options.Width / 2D);
                    break;
                case VisioSide.Right:
                    pinX = target.PinX + (target.Width / 2D) + gap + (options.Width / 2D);
                    break;
                case VisioSide.Bottom:
                    pinY = target.PinY - (target.Height / 2D) - gap - (options.Height / 2D);
                    break;
                case VisioSide.Top:
                    pinY = target.PinY + (target.Height / 2D) + gap + (options.Height / 2D);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(placement));
            }
        }

        private static void EnsureTargetBelongsToPage(VisioPage page, VisioShape target) {
            if (!page.AllShapes().Contains(target)) {
                throw new InvalidOperationException("The target shape must belong to the page.");
            }
        }

        private static VisioCalloutOptions ConvertOptionsToPageUnits(VisioCalloutOptions options, VisioMeasurementUnit unit) {
            return new VisioCalloutOptions {
                Width = options.Width.ToInches(unit),
                Height = options.Height.ToInches(unit),
                LayerName = options.LayerName,
                TargetSide = options.TargetSide,
                CalloutSide = options.CalloutSide,
                LeaderKind = options.LeaderKind,
                RouteLeader = options.RouteLeader,
                RouteOffset = options.RouteOffset.ToInches(unit),
                ShapeStyle = options.ShapeStyle?.Clone(),
                LeaderStyle = options.LeaderStyle?.Clone()
            };
        }

        private static VisioSide ChooseSide(VisioShape target, VisioShape callout) {
            double dx = callout.PinX - target.PinX;
            double dy = callout.PinY - target.PinY;
            if (Math.Abs(dx) >= Math.Abs(dy)) {
                return dx < 0 ? VisioSide.Left : VisioSide.Right;
            }

            return dy < 0 ? VisioSide.Bottom : VisioSide.Top;
        }

        private static VisioSide Opposite(VisioSide side) {
            switch (side) {
                case VisioSide.Left:
                    return VisioSide.Right;
                case VisioSide.Right:
                    return VisioSide.Left;
                case VisioSide.Top:
                    return VisioSide.Bottom;
                case VisioSide.Bottom:
                    return VisioSide.Top;
                default:
                    return VisioSide.Auto;
            }
        }
    }
}
