using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Connection sites of DrawingML preset geometries and the transforms that route a connector between two
    ///     points. Site order and angles follow the preset shape definitions in ECMA-376 Part 1, Annex D.
    /// </summary>
    internal static class PowerPointConnectorGeometry {
        private const int Presentation = OfficeXmlNamespaces.PresentationML;
        private const int Drawing = OfficeXmlNamespaces.DrawingML;

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

        private static readonly Dictionary<OfficePresetShapeType, LocalSite[]> PresetSites = new() {
            [OfficePresetShapeType.Rectangle] = BoxSites,
            [OfficePresetShapeType.RoundRectangle] = BoxSites,
            [OfficePresetShapeType.Diamond] = BoxSites,
            [OfficePresetShapeType.FlowChartProcess] = BoxSites,
            [OfficePresetShapeType.FlowChartAlternateProcess] = BoxSites,
            [OfficePresetShapeType.FlowChartDecision] = BoxSites,
            [OfficePresetShapeType.FlowChartPredefinedProcess] = BoxSites,
            [OfficePresetShapeType.FlowChartInternalStorage] = BoxSites,
            [OfficePresetShapeType.FlowChartTerminator] = BoxSites,
            [OfficePresetShapeType.Ellipse] = EllipseSites,
            [OfficePresetShapeType.FlowChartConnector] = EllipseSites,
            [OfficePresetShapeType.Triangle] = TriangleSites
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

            OfficeXmlNode element = shape.XmlElement;
            OfficeXmlNode properties = element.Is(Presentation, "sp") || element.Is(Presentation, "pic")
                ? element.Child(Presentation, "spPr") : default;
            OfficeXmlNode custom = properties.Child(Drawing, "custGeom");
            IReadOnlyList<LocalSite>? local = null;
            if (custom.Exists) {
                local = ReadCustomSites(custom, shape.Width.ToEmus(), shape.Height.ToEmus());
            } else {
                OfficeXmlNode geometry = properties.Child(Drawing, "prstGeom");
                bool pictureOrFrame = element.Is(Presentation, "pic") || element.Is(Presentation, "graphicFrame");
                if (geometry.IsNone && !pictureOrFrame) return false;
                OfficePresetShapeType preset = pictureOrFrame ? OfficePresetShapeType.Rectangle : default;
                if (geometry.Attribute("prst") is { } token &&
                    !PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out preset)) return false;
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
        internal static void Route(PowerPointConnectionShape connector, PowerPointConnectorKind kind, long startX, long startY,
            long endX, long endY, double? startAngle, double? endAngle) {
            bool startVertical = startAngle.HasValue ? IsVertical(startAngle.Value) : Math.Abs(endY - startY) > Math.Abs(endX - startX);
            bool endVertical = endAngle.HasValue ? IsVertical(endAngle.Value) : startVertical;
            OfficePresetShapeType preset = kind switch {
                PowerPointConnectorKind.Elbow => startVertical == endVertical ? OfficePresetShapeType.BentConnector3 : OfficePresetShapeType.BentConnector2,
                PowerPointConnectorKind.Curved => startVertical == endVertical ? OfficePresetShapeType.CurvedConnector3 : OfficePresetShapeType.CurvedConnector2,
                _ => OfficePresetShapeType.StraightConnector1
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

            OfficeXmlNode element = connector.XmlElement;
            OfficeXmlNode propertyNode = element.Child(Presentation, "spPr");
            if (propertyNode.IsNone) {
                propertyNode = element.CreateElement(Presentation, "spPr");
                OfficeXmlNode nonVisual = element.Child(Presentation, "nvCxnSpPr");
                propertyNode = nonVisual.Exists ? nonVisual.InsertAfterSelf(propertyNode) : element.InsertBefore(propertyNode, element.FirstChild);
            }
            var properties = new AShapePropertiesView(propertyNode);
            ATransform2DView transform = properties.GetOrAddXfrm();
            transform.Element.RemoveChildren(Drawing, "off");
            transform.Element.RemoveChildren(Drawing, "ext");
            APoint2DView offset = transform.GetOrAddOff();
            offset.Element.SetInt64Attribute("x", (long)Math.Round(centerX - cx / 2D));
            offset.Element.SetInt64Attribute("y", (long)Math.Round(centerY - cy / 2D));
            APositiveSize2DView extent = transform.GetOrAddExt();
            extent.Cx = cx;
            extent.Cy = cy;
            transform.Rot = rotate ? 5400000 : null;
            transform.FlipH = flipH ? true : null;
            transform.FlipV = flipV ? true : null;

            APresetGeometry2DView geometry = properties.PrstGeom ?? new APresetGeometry2DView(
                transform.Element.InsertAfterSelf(propertyNode.CreateElement(Drawing, "prstGeom")));
            geometry.Prst = preset.ToOpenXmlToken();
            geometry.GetOrAddAvLst().Element.RemoveAllChildren();
            connector.OwnerPartRef.MarkEdited();
        }

        internal static PowerPointConnectorKind? KindOf(OfficePresetShapeType? preset) {
            if (preset == null) return null;
            if (preset == OfficePresetShapeType.StraightConnector1 || preset == OfficePresetShapeType.Line) return PowerPointConnectorKind.Straight;
            if (preset == OfficePresetShapeType.BentConnector2 || preset == OfficePresetShapeType.BentConnector3
                || preset == OfficePresetShapeType.BentConnector4 || preset == OfficePresetShapeType.BentConnector5) return PowerPointConnectorKind.Elbow;
            if (preset == OfficePresetShapeType.CurvedConnector2 || preset == OfficePresetShapeType.CurvedConnector3
                || preset == OfficePresetShapeType.CurvedConnector4 || preset == OfficePresetShapeType.CurvedConnector5) return PowerPointConnectorKind.Curved;
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

        private static List<LocalSite>? ReadCustomSites(OfficeXmlNode geometry, long width, long height) {
            OfficeXmlNode list = geometry.Child(Drawing, "cxnLst");
            if (list.IsNone || width <= 0 || height <= 0) return null;
            var sites = new List<LocalSite>();
            foreach (OfficeXmlNode site in list.Children(Drawing, "cxn")) {
                // Only literal coordinates can be resolved without evaluating the shape guide formulas.
                OfficeXmlNode position = site.Child(Drawing, "pos");
                if (position.Int64Attribute("x") is not long x || position.Int64Attribute("y") is not long y ||
                    site.Int32Attribute("ang") is not int angle) {
                    return null;
                }

                sites.Add(new LocalSite(x / (double)width, y / (double)height, angle / 60000));
            }

            return sites;
        }
    }
}
