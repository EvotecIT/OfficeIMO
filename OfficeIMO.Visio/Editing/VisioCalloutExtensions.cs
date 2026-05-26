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
            VisioShape callout = page.AddRectangle(pinX, pinY, effectiveOptions.Width, effectiveOptions.Height, text);
            ConfigureCallout(page, target, callout, effectiveOptions);
            return callout;
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
            VisioShape callout = new VisioShape(id, pinX, pinY, effectiveOptions.Width, effectiveOptions.Height, text) {
                Name = "Callout",
                NameU = "Rectangle"
            };
            page.Shapes.Add(callout);
            ConfigureCallout(page, target, callout, effectiveOptions);
            return callout;
        }

        private static void ConfigureCallout(VisioPage page, VisioShape target, VisioShape callout, VisioCalloutOptions options) {
            if (!page.AllShapes().Contains(target)) {
                throw new InvalidOperationException("The target shape must belong to the page.");
            }

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
