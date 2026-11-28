using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private readonly record struct MapViewport(MapBounds Bounds, bool Mercator, double Left, double Top, double Width,
        double Height, double Scale, double ProjectedNorth) {
        internal OfficePoint Project(MapCoordinate value) {
            double x = Left + (value.Longitude - Bounds.West) * Scale;
            double y = Top + (ProjectedNorth - MapLatitude(value.Latitude, Mercator)) * Scale;
            return new OfficePoint(Math.Max(Left, Math.Min(Left + Width, x)), Math.Max(Top, Math.Min(Top + Height, y)));
        }

        internal static bool TryCreate(MapBounds bounds, bool mercator, double left, double top, double right,
            double bottom, out MapViewport view) {
            view = default;
            double longitudeSpan = bounds.East - bounds.West;
            double projectedNorth = MapLatitude(bounds.North, mercator);
            double latitudeSpan = projectedNorth - MapLatitude(bounds.South, mercator);
            if (longitudeSpan < 1e-8 || latitudeSpan < 1e-8 || right <= left || bottom <= top) return false;
            double scale = Math.Min((right - left) / longitudeSpan, (bottom - top) / latitudeSpan);
            double width = longitudeSpan * scale;
            double height = latitudeSpan * scale;
            view = new MapViewport(bounds, mercator, (left + right - width) / 2, (top + bottom - height) / 2,
                width, height, scale, projectedNorth);
            return true;
        }

        private static double MapLatitude(double latitude, bool mercator) {
            if (!mercator) return latitude;
            const double degrees = 180 / Math.PI;
            return Math.Log(Math.Tan(Math.PI / 4 + latitude / (2 * degrees))) * degrees;
        }
    }

    private static bool PaintMapPolygon(OfficeDrawing drawing, MapViewport view, MapPolygon polygon, OfficeColor? fill,
        OfficeColor? border, double borderWidth) {
        if ((fill is not OfficeColor { A: > 0 } && (border is not OfficeColor { A: > 0 } || borderWidth <= 0)) ||
            !view.Bounds.Intersects(polygon.Bounds)) return false;

        bool clip = !view.Bounds.Contains(polygon.Bounds);
        var paths = new List<OfficePathCommand>();
        bool painted = false;
        if (!clip || fill.HasValue) {
            foreach (MapCoordinate[] ring in polygon.Rings) {
                IReadOnlyList<MapCoordinate> visible = clip ? ClipMapRing(ring, view.Bounds) : ring;
                if (visible.Count < 3) continue;
                AppendMapRing(paths, view, visible);
            }
            if (paths.Count > 0) {
                PaintMapPath(drawing, paths, fill, clip ? null : border, borderWidth);
                painted = true;
            }
        }
        if (!clip || !border.HasValue || borderWidth <= 0) return painted;

        // Stroke the original visible edges so a crop boundary does not become a coastline or region boundary.
        paths.Clear();
        foreach (MapCoordinate[] ring in polygon.Rings) AppendClippedMapOutline(paths, view, ring);
        if (paths.Count == 0) return painted;
        PaintMapPath(drawing, paths, null, border, borderWidth);
        return true;
    }

    private static void AppendMapRing(List<OfficePathCommand> paths, MapViewport view, IReadOnlyList<MapCoordinate> ring) {
        OfficePoint first = view.Project(ring[0]);
        OfficePoint previous = first;
        double area = 0;
        int from = paths.Count;
        paths.Add(OfficePathCommand.MoveTo(first));
        for (int index = 1; index < ring.Count; index++) {
            OfficePoint point = view.Project(ring[index]);
            area += previous.X * point.Y - point.X * previous.Y;
            paths.Add(OfficePathCommand.LineTo(point));
            previous = point;
        }
        area += previous.X * first.Y - first.X * previous.Y;
        if (Math.Abs(area) < 1e-10) paths.RemoveRange(from, paths.Count - from);
        else paths.Add(OfficePathCommand.Close());
    }

    private static void AppendClippedMapOutline(List<OfficePathCommand> paths, MapViewport view, MapCoordinate[] ring) {
        MapCoordinate previous = ring[ring.Length - 1];
        OfficePoint? last = null;
        foreach (MapCoordinate current in ring) {
            if (TryClipMapSegment(view.Bounds, previous, current, out MapCoordinate a, out MapCoordinate b)) {
                OfficePoint start = view.Project(a);
                OfficePoint end = view.Project(b);
                if (start != end) {
                    if (start != last) paths.Add(OfficePathCommand.MoveTo(start));
                    paths.Add(OfficePathCommand.LineTo(end));
                    last = end;
                }
            } else {
                last = null;
            }
            previous = current;
        }
    }

    private static void PaintMapPath(OfficeDrawing drawing, List<OfficePathCommand> paths, OfficeColor? fill, OfficeColor? border,
        double borderWidth) {
        OfficeShape shape = OfficeShape.Path(drawing.Width, drawing.Height, paths);
        shape.FillColor = fill;
        shape.FillRule = OfficeFillRule.EvenOdd;
        if (borderWidth > 0) {
            shape.StrokeColor = border;
            shape.StrokeWidth = borderWidth;
            shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;
        }
        drawing.AddShape(shape, 0, 0);
    }

    private static IReadOnlyList<MapCoordinate> ClipMapRing(MapCoordinate[] ring, MapBounds view) {
        IReadOnlyList<MapCoordinate> current = ring;
        for (int side = 0; side < 4 && current.Count > 2; side++) {
            var next = new List<MapCoordinate>(current.Count);
            MapCoordinate previous = current[current.Count - 1];
            double last = GetMapEdgeDistance(view, previous, side);
            foreach (MapCoordinate point in current) {
                double distance = GetMapEdgeDistance(view, point, side);
                if ((last >= 0) != (distance >= 0))
                    next.Add(MixMapPosition(previous, point, last / (last - distance)));
                if (distance >= 0) next.Add(point);
                previous = point;
                last = distance;
            }
            current = next;
        }
        return current;
    }

    private static double GetMapEdgeDistance(MapBounds view, MapCoordinate point, int edge) => edge switch {
        0 => point.Longitude - view.West,
        1 => view.East - point.Longitude,
        2 => point.Latitude - view.South,
        _ => view.North - point.Latitude
    };

    private static bool TryClipMapSegment(MapBounds view, MapCoordinate a, MapCoordinate b, out MapCoordinate start, out MapCoordinate end) {
        start = default;
        end = default;
        double longitude = b.Longitude - a.Longitude;
        double latitude = b.Latitude - a.Latitude;
        double first = 0;
        double last = 1;
        if (!ClipMapEdge(-longitude, a.Longitude - view.West, ref first, ref last) ||
            !ClipMapEdge(longitude, view.East - a.Longitude, ref first, ref last) ||
            !ClipMapEdge(-latitude, a.Latitude - view.South, ref first, ref last) ||
            !ClipMapEdge(latitude, view.North - a.Latitude, ref first, ref last)) return false;
        start = MixMapPosition(a, b, first);
        end = MixMapPosition(a, b, last);
        return true;
    }

    private static bool ClipMapEdge(double step, double distance, ref double first, ref double last) {
        if (step == 0) return distance >= 0;
        double at = distance / step;
        if (step < 0) first = Math.Max(first, at);
        else last = Math.Min(last, at);
        return first <= last;
    }

    private static MapCoordinate MixMapPosition(MapCoordinate a, MapCoordinate b, double t) =>
        new(a.Longitude + (b.Longitude - a.Longitude) * t, a.Latitude + (b.Latitude - a.Latitude) * t);

    private static bool TryFindMapRegionAnchor(MapViewport view, MapFeature feature, out MapCoordinate anchor) {
        anchor = feature.LabelPosition;
        foreach (MapPolygon polygon in feature.Polygons) {
            if (!view.Bounds.Intersects(polygon.Bounds)) continue;
            if (view.Bounds.Contains(anchor) && MapPolygonContains(polygon, anchor)) return true;
            var bounds = new MapBounds(Math.Max(view.Bounds.West, polygon.Bounds.West), Math.Max(view.Bounds.South, polygon.Bounds.South),
                Math.Min(view.Bounds.East, polygon.Bounds.East), Math.Min(view.Bounds.North, polygon.Bounds.North));
            MapCoordinate center = bounds.Center;
            if (MapPolygonContains(polygon, center)) {
                anchor = center;
                return true;
            }
            double horizontal = (bounds.East - bounds.West) / 10;
            double vertical = (bounds.North - bounds.South) / 10;
            for (int row = 1; row <= 9; row += 2) {
                for (int column = 1; column <= 9; column += 2) {
                    MapCoordinate candidate = new(bounds.West + column * horizontal, bounds.South + row * vertical);
                    if (!MapPolygonContains(polygon, candidate)) continue;
                    anchor = candidate;
                    return true;
                }
            }
        }
        return false;
    }

    private static bool MapPolygonContains(MapPolygon polygon, MapCoordinate point) {
        bool inside = false;
        foreach (MapCoordinate[] ring in polygon.Rings) {
            MapCoordinate previous = ring[ring.Length - 1];
            foreach (MapCoordinate current in ring) {
                if ((current.Latitude > point.Latitude) != (previous.Latitude > point.Latitude) &&
                    point.Longitude < (previous.Longitude - current.Longitude) * (point.Latitude - current.Latitude) /
                        (previous.Latitude - current.Latitude) + current.Longitude) inside = !inside;
                previous = current;
            }
        }
        return inside;
    }
}
