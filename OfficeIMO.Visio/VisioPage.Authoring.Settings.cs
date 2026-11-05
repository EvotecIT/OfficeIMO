using System;
using Color = OfficeIMO.OfficeColor;
namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>Sets print orientation.</summary>
        /// <param name="orientation">Print orientation.</param>
        public VisioPage SetPrintOrientation(VisioPagePrintOrientation orientation) {
            PrintOrientation = orientation;
            return this;
        }

        /// <summary>Locks or unlocks page replacement.</summary>
        /// <param name="locked">Whether replacement is locked.</param>
        public VisioPage SetLockReplace(bool locked = true) {
            PageLockReplace = locked;
            return this;
        }

        /// <summary>Locks or unlocks page duplication.</summary>
        /// <param name="locked">Whether duplication is locked.</param>
        public VisioPage SetLockDuplicate(bool locked = true) {
            PageLockDuplicate = locked;
            return this;
        }

        /// <summary>Sets how Visio determines the drawing page size.</summary>
        /// <param name="sizeType">Drawing size behavior.</param>
        public VisioPage SetDrawingSize(VisioDrawingSizeType sizeType) {
            DrawingSizeType = sizeType;
            return this;
        }

        /// <summary>Enables or disables automatic page resizing to fit the diagram.</summary>
        /// <param name="enabled">Whether automatic drawing resize is enabled.</param>
        public VisioPage SetAutoResizeDrawing(bool enabled = true) {
            AutoResizeDrawing = enabled;
            return this;
        }

        /// <summary>Enables or disables automatic shape splitting on this page.</summary>
        /// <param name="enabled">Whether automatic shape splitting is enabled.</param>
        public VisioPage SetShapeSplitting(bool enabled = true) {
            AllowShapeSplitting = enabled;
            return this;
        }

        /// <summary>Shows or hides the page name in Visio UI page lists.</summary>
        /// <param name="visibility">Page UI visibility.</param>
        public VisioPage SetUiVisibility(VisioPageUiVisibility visibility) {
            UiVisibility = visibility;
            return this;
        }

        /// <summary>Sets Visio's page-level placement style.</summary>
        /// <param name="style">Placement style.</param>
        public VisioPage SetPlacementStyle(VisioPlacementStyle style) {
            PlacementStyle = style;
            return this;
        }

        /// <summary>Sets Visio's page-level placement analysis depth.</summary>
        /// <param name="depth">Placement depth.</param>
        public VisioPage SetPlacementDepth(VisioPlacementDepth depth) {
            PlacementDepth = depth;
            return this;
        }

        /// <summary>Sets Visio's page-level placement flip behavior.</summary>
        /// <param name="flip">Placement flip flags.</param>
        public VisioPage SetPlacementFlip(VisioPlacementFlip flip) {
            PlacementFlip = flip;
            return this;
        }

        /// <summary>Enables or disables moving nearby shapes away on drop.</summary>
        /// <param name="enabled">Whether nearby placeable shapes move away on drop.</param>
        public VisioPage SetMoveShapesAwayOnDrop(bool enabled = true) {
            MoveShapesAwayOnDrop = enabled;
            return this;
        }

        /// <summary>Enables or disables enlarging the page after Visio lays out shapes.</summary>
        /// <param name="enabled">Whether Visio should resize the page after layout.</param>
        public VisioPage SetResizePageToFitLayout(bool enabled = true) {
            ResizePageToFitLayout = enabled;
            return this;
        }

        /// <summary>Enables or disables Visio's internal layout grid for page layout.</summary>
        /// <param name="enabled">Whether Visio should use the internal layout grid.</param>
        public VisioPage SetEnableLayoutGrid(bool enabled = true) {
            EnableLayoutGrid = enabled;
            return this;
        }

        /// <summary>Sets Visio's page-level connector routing style.</summary>
        /// <param name="style">Routing style to use for connectors without local routing.</param>
        public VisioPage SetConnectorRouteStyle(VisioPageRouteStyle style) {
            ConnectorRouteStyle = style;
            return this;
        }

        /// <summary>Sets the default appearance for routed connectors on this page.</summary>
        /// <param name="appearance">Default routed connector appearance.</param>
        public VisioPage SetConnectorRouteAppearance(VisioLineRouteExtension appearance) {
            ConnectorRouteAppearance = appearance;
            return this;
        }

        /// <summary>Sets the default line jump behavior for this page.</summary>
        /// <param name="style">Line jump style.</param>
        /// <param name="code">Which connectors receive line jumps.</param>
        /// <param name="horizontalDirection">Default jump direction for horizontal dynamic connectors.</param>
        /// <param name="verticalDirection">Default jump direction for vertical dynamic connectors.</param>
        public VisioPage SetLineJumps(
            VisioLineJumpStyle style,
            VisioLineJumpCode code,
            VisioHorizontalLineJumpDirection horizontalDirection = VisioHorizontalLineJumpDirection.Default,
            VisioVerticalLineJumpDirection verticalDirection = VisioVerticalLineJumpDirection.Default) {
            LineJumpStyle = style;
            LineJumpCode = code;
            HorizontalLineJumpDirection = horizontalDirection;
            VerticalLineJumpDirection = verticalDirection;
            return this;
        }
    }
}
