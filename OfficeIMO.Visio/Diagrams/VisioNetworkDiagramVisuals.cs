using System;

namespace OfficeIMO.Visio.Diagrams {
    internal static class VisioNetworkDiagramVisuals {
        internal static void GetNodeShape(VisioNetworkNodeKind kind, double nodeWidth, double nodeHeight, out string masterNameU, out double width, out double height) {
            width = nodeWidth;
            height = nodeHeight;
            switch (kind) {
                case VisioNetworkNodeKind.User:
                case VisioNetworkNodeKind.Internet:
                case VisioNetworkNodeKind.Wireless:
                    masterNameU = "Circle";
                    width = 0.9D;
                    height = 0.9D;
                    break;
                case VisioNetworkNodeKind.Firewall:
                case VisioNetworkNodeKind.Router:
                    masterNameU = "Decision";
                    width = nodeWidth * 0.95D;
                    height = nodeHeight * 1.2D;
                    break;
                case VisioNetworkNodeKind.Storage:
                case VisioNetworkNodeKind.Database:
                    masterNameU = "Data";
                    break;
                case VisioNetworkNodeKind.Switch:
                    masterNameU = "Rectangle";
                    width = nodeWidth * 1.15D;
                    height = nodeHeight * 0.75D;
                    break;
                case VisioNetworkNodeKind.Note:
                    masterNameU = "Rectangle";
                    width = nodeWidth * 1.55D;
                    height = nodeHeight * 1.15D;
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        internal static VisioShapeStyle GetNodeStyle(VisioStyleTheme theme, VisioNetworkNodeKind kind) {
            switch (kind) {
                case VisioNetworkNodeKind.User:
                case VisioNetworkNodeKind.Wireless:
                    return theme.Marker;
                case VisioNetworkNodeKind.Firewall:
                case VisioNetworkNodeKind.Router:
                    return theme.Emphasis;
                case VisioNetworkNodeKind.Storage:
                case VisioNetworkNodeKind.Database:
                    return theme.Success;
                case VisioNetworkNodeKind.Note:
                    return theme.Container;
                case VisioNetworkNodeKind.Internet:
                    return theme.Decision;
                default:
                    return theme.Primary;
            }
        }

        internal static VisioConnectorStyle GetConnectorStyle(VisioStyleTheme theme, VisioNetworkLinkKind kind) {
            switch (kind) {
                case VisioNetworkLinkKind.Management:
                    return theme.ControlConnector;
                case VisioNetworkLinkKind.Trunk:
                    return theme.DataConnector;
                case VisioNetworkLinkKind.Wireless:
                    return theme.ControlConnector;
                default:
                    return theme.Connector;
            }
        }

        internal static VisioShape CreateBackgroundZone(
            VisioDocument document,
            string id,
            double pinX,
            double pinY,
            double width,
            double height,
            string text,
            VisioStyleTheme theme) {
            VisioShape shape = new(id, pinX, pinY, width, height, text) {
                NameU = "Rectangle",
            };
            theme.Container.ApplyTo(shape);
            shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
            return shape;
        }

        internal static void ResolveSides(VisioShape from, VisioShape to, out VisioSide fromSide, out VisioSide toSide) {
            double dx = to.PinX - from.PinX;
            double dy = to.PinY - from.PinY;
            if (Math.Abs(dx) >= Math.Abs(dy)) {
                if (dx >= 0D) {
                    fromSide = VisioSide.Right;
                    toSide = VisioSide.Left;
                } else {
                    fromSide = VisioSide.Left;
                    toSide = VisioSide.Right;
                }

                return;
            }

            if (dy >= 0D) {
                fromSide = VisioSide.Top;
                toSide = VisioSide.Bottom;
            } else {
                fromSide = VisioSide.Bottom;
                toSide = VisioSide.Top;
            }
        }
    }
}
