using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Connection sites of DrawingML preset geometries and the transforms that route a connector between two
    ///     points. Site order and angles follow the preset shape definitions in ECMA-376 Part 1, Annex D.
    /// </summary>
    internal static class PowerPointConnectorGeometry {
        private readonly struct LocalSite {
            internal LocalSite(double x, double y, int angle) {
                X = x;
                Y = y;
                Angle = angle;
            }

            internal double X { get; }
            internal double Y { get; }
            internal int Angle { get; }
        }

        private const double EllipseInset = 0.1464466094067262D; // (1 - cos 45°) / 2

        private static readonly LocalSite[] BoxSites = {
            new(0.5D, 0D, 270), new(0D, 0.5D, 180), new(0.5D, 1D, 90), new(1D, 0.5D, 0)
        };

        private static readonly LocalSite[] EllipseSites = {
            new(0.5D, 0D, 270), new(EllipseInset, EllipseInset, 270), new(0D, 0.5D, 180),
            new(EllipseInset, 1D - EllipseInset, 90), new(0.5D, 1D, 90), new(1D - EllipseInset, 1D - EllipseInset, 90),
            new(1D, 0.5D, 0), new(1D - EllipseInset, EllipseInset, 270)
        };

        // Isosceles triangle with its default apex adjustment (50%).
        private static readonly LocalSite[] TriangleSites = {
            new(0.5D, 0D, 270), new(0.25D, 0.5D, 180), new(0D, 1D, 90), new(0.5D, 1D, 90), new(1D, 1D, 90), new(0.75D, 0.5D, 0)
        };

        private static readonly Dictionary<A.ShapeTypeValues, LocalSite[]> PresetSites = new() {
            [A.ShapeTypeValues.Rectangle] = BoxSites,
            [A.ShapeTypeValues.RoundRectangle] = BoxSites,
            [A.ShapeTypeValues.Diamond] = BoxSites,
            [A.ShapeTypeValues.FlowChartProcess] = BoxSites,
            [A.ShapeTypeValues.FlowChartAlternateProcess] = BoxSites,
            [A.ShapeTypeValues.FlowChartDecision] = BoxSites,
            [A.ShapeTypeValues.FlowChartPredefinedProcess] = BoxSites,
            [A.ShapeTypeValues.FlowChartInternalStorage] = BoxSites,
            [A.ShapeTypeValues.FlowChartTerminator] = BoxSites,
            [A.ShapeTypeValues.Ellipse] = EllipseSites,
            [A.ShapeTypeValues.FlowChartConnector] = EllipseSites,
            [A.ShapeTypeValues.Triangle] = TriangleSites
        };

        /// <summary>Sites in slide coordinates (EMU), with shape flips and rotation applied.</summary>
        internal static bool TryGetSites(PowerPointShape shape, out List<PowerPointConnectionSite> sites) {
            sites = new List<PowerPointConnectionSite>();
            if (shape is PowerPointConnectionShape) return false;
            // A shape that follows a layout placeholder has no local position, so its sites cannot be placed.
            if (!shape.TryGetBoundsPoints(out _, out _, out double boundsWidth, out double boundsHeight)
                || (boundsWidth <= 0D && boundsHeight <= 0D)) {
                return false;
            }

            ShapeProperties? properties = GetShapeProperties(shape.Element);
            IReadOnlyList<LocalSite>? local = null;
            A.CustomGeometry? custom = properties?.GetFirstChild<A.CustomGeometry>();
            if (custom != null) {
                local = ReadCustomSites(custom, shape.Width.ToEmus(), shape.Height.ToEmus());
            } else {
                A.ShapeTypeValues preset = properties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value
                    ?? (shape.Element is Picture or GraphicFrame ? A.ShapeTypeValues.Rectangle : default);
                if (properties?.GetFirstChild<A.PresetGeometry>() == null && shape.Element is not (Picture or GraphicFrame)) return false;
                if (PresetSites.TryGetValue(preset, out LocalSite[]? known)) local = known;
            }

            if (local == null || local.Count == 0) return false;

            double left = shape.Left.ToEmus(), top = shape.Top.ToEmus();
            double width = shape.Width.ToEmus(), height = shape.Height.ToEmus();
            double centerX = left + width / 2D, centerY = top + height / 2D;
            double rotation = shape.Rotation ?? 0D;
            bool flipH = shape.HorizontalFlip == true, flipV = shape.VerticalFlip == true;
            double radians = rotation * Math.PI / 180D;
            double cos = Math.Cos(radians), sin = Math.Sin(radians);

            for (int i = 0; i < local.Count; i++) {
                double fx = flipH ? 1D - local[i].X : local[i].X;
                double fy = flipV ? 1D - local[i].Y : local[i].Y;
                double angle = local[i].Angle;
                if (flipH) angle = 180D - angle;
                if (flipV) angle = -angle;
                double dx = (fx - 0.5D) * width, dy = (fy - 0.5D) * height;
                double x = centerX + dx * cos - dy * sin;
                double y = centerY + dx * sin + dy * cos;
                angle = ((angle + rotation) % 360D + 360D) % 360D;
                sites.Add(new PowerPointConnectionSite(i, OfficeLength.Emus((long)Math.Round(x)),
                    OfficeLength.Emus((long)Math.Round(y)), angle));
            }

            return true;
        }

        internal static bool TryGetSite(PowerPointShape shape, int index, out PowerPointConnectionSite site) {
            site = default;
            if (!TryGetSites(shape, out List<PowerPointConnectionSite> sites) || index < 0 || index >= sites.Count) return false;
            site = sites[index];
            return true;
        }

        internal static PowerPointConnectionSite GetSite(PowerPointShape shape, int index, string parameterName) {
            if (!shape.TryGetBoundsPoints(out _, out _, out _, out _)) {
                throw new NotSupportedException("Shape '" + shape.Name + "' has no position of its own (it follows a layout "
                    + "placeholder), so connectors cannot be glued to it.");
            }

            if (!TryGetSites(shape, out List<PowerPointConnectionSite> sites)) {
                throw new NotSupportedException("Shape '" + shape.Name + "' has no known connection sites. Glue connectors to "
                    + "rectangles, rounded rectangles, ellipses, triangles, diamonds, common flowchart shapes, pictures, "
                    + "tables, charts or custom geometry with connection sites.");
            }

            if (index < 0 || index >= sites.Count) {
                throw new ArgumentOutOfRangeException(parameterName, "Shape '" + shape.Name + "' has " + sites.Count
                    + " connection sites (0 to " + (sites.Count - 1) + ").");
            }

            return sites[index];
        }

        /// <summary>Picks the closest pair of sites that face each other.</summary>
        internal static (int Start, int End) ChooseSites(PowerPointShape from, PowerPointShape to) {
            PowerPointConnectionSite[] starts = GetAll(from, nameof(from));
            PowerPointConnectionSite[] ends = GetAll(to, nameof(to));
            double best = double.MaxValue;
            (int, int) choice = (0, 0);
            foreach (PowerPointConnectionSite start in starts) {
                foreach (PowerPointConnectionSite end in ends) {
                    double dx = end.X.ToEmus() - start.X.ToEmus(), dy = end.Y.ToEmus() - start.Y.ToEmus();
                    double distance = Math.Sqrt(dx * dx + dy * dy);
                    // Prefer sites whose outward direction points toward the other end.
                    double penalty = Facing(start.Angle, dx, dy) ? 1D : 1.5D;
                    penalty *= Facing(end.Angle, -dx, -dy) ? 1D : 1.5D;
                    if (distance * penalty < best) {
                        best = distance * penalty;
                        choice = (start.Index, end.Index);
                    }
                }
            }

            return choice;
        }

        /// <summary>Writes preset, transform and flips so the connector runs from start to end.</summary>
        internal static void Route(ConnectionShape connector, PowerPointConnectorKind kind, long startX, long startY,
            long endX, long endY, double? startAngle, double? endAngle) {
            bool startVertical = startAngle.HasValue ? IsVertical(startAngle.Value) : Math.Abs(endY - startY) > Math.Abs(endX - startX);
            bool endVertical = endAngle.HasValue ? IsVertical(endAngle.Value) : startVertical;
            A.ShapeTypeValues preset = kind switch {
                PowerPointConnectorKind.Elbow => startVertical == endVertical ? A.ShapeTypeValues.BentConnector3 : A.ShapeTypeValues.BentConnector2,
                PowerPointConnectorKind.Curved => startVertical == endVertical ? A.ShapeTypeValues.CurvedConnector3 : A.ShapeTypeValues.CurvedConnector2,
                _ => A.ShapeTypeValues.StraightConnector1
            };
            // Bent and curved connectors start with a horizontal leg; a 90° rotation makes it vertical.
            bool rotate = kind != PowerPointConnectorKind.Straight && startVertical;

            long dx = Math.Abs(endX - startX), dy = Math.Abs(endY - startY);
            double centerX = (startX + endX) / 2D, centerY = (startY + endY) / 2D;
            long cx = rotate ? dy : dx, cy = rotate ? dx : dy;
            bool flipH, flipV;
            if (rotate) {
                flipH = startY > endY;
                flipV = startX < endX;
            } else {
                flipH = endX < startX;
                flipV = endY < startY;
            }

            ShapeProperties properties = connector.ShapeProperties ??= new ShapeProperties();
            A.Transform2D transform = properties.Transform2D ??= new A.Transform2D();
            transform.Offset = new A.Offset {
                X = (long)Math.Round(centerX - cx / 2D),
                Y = (long)Math.Round(centerY - cy / 2D)
            };
            transform.Extents = new A.Extents { Cx = cx, Cy = cy };
            transform.Rotation = rotate ? 5400000 : null;
            transform.HorizontalFlip = flipH ? true : null;
            transform.VerticalFlip = flipV ? true : null;

            A.PresetGeometry geometry = properties.GetFirstChild<A.PresetGeometry>() ?? new A.PresetGeometry();
            geometry.Preset = preset;
            geometry.AdjustValueList ??= new A.AdjustValueList();
            geometry.AdjustValueList.RemoveAllChildren();
            if (geometry.Parent == null) properties.InsertAfter(geometry, transform);
        }

        internal static PowerPointConnectorKind? KindOf(A.ShapeTypeValues? preset) {
            if (preset == null) return null;
            if (preset == A.ShapeTypeValues.StraightConnector1 || preset == A.ShapeTypeValues.Line) return PowerPointConnectorKind.Straight;
            if (preset == A.ShapeTypeValues.BentConnector2 || preset == A.ShapeTypeValues.BentConnector3
                || preset == A.ShapeTypeValues.BentConnector4 || preset == A.ShapeTypeValues.BentConnector5) return PowerPointConnectorKind.Elbow;
            if (preset == A.ShapeTypeValues.CurvedConnector2 || preset == A.ShapeTypeValues.CurvedConnector3
                || preset == A.ShapeTypeValues.CurvedConnector4 || preset == A.ShapeTypeValues.CurvedConnector5) return PowerPointConnectorKind.Curved;
            return null;
        }

        private static PowerPointConnectionSite[] GetAll(PowerPointShape shape, string parameterName) {
            GetSite(shape, 0, parameterName);
            TryGetSites(shape, out List<PowerPointConnectionSite> sites);
            return sites.ToArray();
        }

        private static bool Facing(double angle, double dx, double dy) {
            double radians = angle * Math.PI / 180D;
            return Math.Cos(radians) * dx + Math.Sin(radians) * dy > 0D;
        }

        private static bool IsVertical(double angle) {
            double normalized = (angle % 180D + 180D) % 180D;
            return normalized > 45D && normalized < 135D;
        }

        private static List<LocalSite>? ReadCustomSites(A.CustomGeometry geometry, long width, long height) {
            A.ConnectionSiteList? list = geometry.ConnectionSiteList;
            if (list == null || width <= 0 || height <= 0) return null;
            var sites = new List<LocalSite>();
            foreach (A.ConnectionSite site in list.Elements<A.ConnectionSite>()) {
                // Only literal coordinates can be resolved without evaluating the shape guide formulas.
                if (!long.TryParse(site.Position?.X?.Value, out long x) || !long.TryParse(site.Position?.Y?.Value, out long y)
                    || !int.TryParse(site.Angle?.Value, out int angle)) {
                    return null;
                }

                sites.Add(new LocalSite(x / (double)width, y / (double)height, angle / 60000));
            }

            return sites;
        }

        private static ShapeProperties? GetShapeProperties(OpenXmlElement element) => element switch {
            Shape shape => shape.ShapeProperties,
            Picture picture => picture.ShapeProperties,
            _ => null
        };
    }
}
