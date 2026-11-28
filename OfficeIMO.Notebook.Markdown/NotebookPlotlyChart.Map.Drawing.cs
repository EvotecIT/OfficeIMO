using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static OfficeDrawing? DrawMap(MapLayout layout, List<MapScatter> scatter, MapChoropleth? regions,
        CancellationToken cancellationToken, List<NotebookChartMark>? marks) {
        MapSettings settings = layout.Settings;
        MapAppearance style = settings.Appearance;
        bool colorbar = regions?.Scale.Show == true;
        int legendCount = layout.ShowLegend ? scatter.Count(trace => trace.Legend) : 0;
        bool bottomLegend = legendCount != 0 && (layout.Legend == OfficeChartLegendPosition.Bottom || layout.Width < 360 || colorbar);
        double right = layout.Width - (colorbar ? 64 : legendCount != 0 && !bottomLegend ? 92 : 20);
        double bottom = layout.Height - 22 - (bottomLegend ? ((legendCount + 2) / 3) * 15 : 0);
        double top = string.IsNullOrWhiteSpace(layout.Title) ? 15 : 40;
        MapBounds bounds = GetMapViewportBounds(settings, scatter, regions);
        if (!MapViewport.TryCreate(bounds, settings.Mercator, 20, top, right, bottom, out MapViewport view)) return null;

        var drawing = new OfficeDrawing(layout.Width, layout.Height);
        var page = OfficeShape.Rectangle(layout.Width, layout.Height);
        page.FillColor = layout.Style.BackgroundColor;
        drawing.AddShape(page, 0, 0);
        var panel = OfficeShape.Rectangle(view.Width, view.Height);
        panel.FillColor = style.Background;
        drawing.AddShape(panel, view.Left, view.Top);

        if (!string.IsNullOrWhiteSpace(layout.Title))
            drawing.AddText(layout.Title!, 12, 10, layout.Width - 24, 24, new OfficeFontInfo(null, 16, OfficeFontStyle.Bold),
                layout.TextColor, OfficeTextAlignment.Center, shrinkToFit: true);

        if (style.Land.HasValue || style.Coast.HasValue) {
            foreach (MapPolygon polygon in GeographicWorld.Value.Land) {
                cancellationToken.ThrowIfCancellationRequested();
                PaintMapPolygon(drawing, view, polygon, style.Land, style.Coast, style.CoastWidth);
            }
        }
        if (style.Borders.HasValue) {
            foreach (MapFeature country in GeographicWorld.Value.Countries) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (MapPolygon polygon in country.Polygons)
                    PaintMapPolygon(drawing, view, polygon, null, style.Borders, style.BorderWidth);
            }
        }
        PaintMapGrid(drawing, view, settings.Longitude, longitude: true);
        PaintMapGrid(drawing, view, settings.Latitude, longitude: false);
        if (regions != null) DrawMapRegions(drawing, view, regions, cancellationToken, marks);
        foreach (MapScatter trace in scatter)
            DrawMapScatter(drawing, view, trace, layout.TextColor, cancellationToken, marks);

        if (style.Frame.HasValue && style.FrameWidth > 0) {
            var frame = OfficeShape.Rectangle(view.Width, view.Height);
            frame.StrokeColor = style.Frame;
            frame.StrokeWidth = style.FrameWidth;
            drawing.AddShape(frame, view.Left, view.Top);
        }
        if (colorbar) PaintSceneColorbar(drawing, regions!.Scale, layout.TextColor, regions.ScaleTitle);
        if (legendCount > 0) PaintMapLegend(drawing, layout, scatter, legendCount, bottomLegend, bottom);
        return drawing;
    }

    private static MapBounds GetMapViewportBounds(MapSettings settings, List<MapScatter> scatter, MapChoropleth? regions) {
        double latitudeLimit = settings.Mercator ? 85.0511 : 90;
        var world = new MapBounds(-180, -latitudeLimit, 180, latitudeLimit);
        MapBounds bounds = world;
        if (settings.Fit != MapFit.World) {
            MapBounds fit = MapBounds.Empty;
            foreach (MapScatter trace in scatter) {
                foreach (MapCoordinate? position in trace.Coordinates) {
                    if (position is MapCoordinate coordinate) fit = fit.Include(coordinate);
                }
            }
            if (settings.Fit == MapFit.GeoJson && regions?.GeoJsonBounds is MapBounds geometry) {
                fit = fit.Include(geometry);
            } else if (regions != null) {
                foreach (MapRegion region in regions.Regions) fit = fit.Include(region.Feature.Bounds);
            }
            if (fit.West <= fit.East && fit.South <= fit.North) {
                double paddingX = Math.Max(0.04, (fit.East - fit.West) * 0.10);
                double paddingY = Math.Max(0.04, (fit.North - fit.South) * 0.10);
                bounds = new MapBounds(Math.Max(world.West, fit.West - paddingX), Math.Max(world.South, fit.South - paddingY),
                    Math.Min(world.East, fit.East + paddingX), Math.Min(world.North, fit.North + paddingY));
            }
        }
        if (settings.Longitude.Range is double[] lon) bounds = bounds with { West = lon[0], East = lon[1] };
        if (settings.Latitude.Range is double[] lat)
            bounds = bounds with { South = Math.Max(world.South, lat[0]), North = Math.Min(world.North, lat[1]) };
        return bounds;
    }

    private static void PaintMapGrid(OfficeDrawing drawing, MapViewport view, MapAxis axis, bool longitude) {
        if (!axis.ShowGrid || axis.GridWidth <= 0) return;
        double from = longitude ? view.Bounds.West : view.Bounds.South;
        double to = longitude ? view.Bounds.East : view.Bounds.North;
        double step = axis.Step > 0 ? axis.Step : MapGridStep((to - from) / 6);
        double first = axis.Origin + Math.Ceiling((from - axis.Origin) / step) * step;
        for (int index = 0; index < 64; index++) {
            double value = first + index * step;
            if (value >= to) break;
            if (value <= from) continue;
            MapCoordinate a = longitude ? new MapCoordinate(value, view.Bounds.South) : new MapCoordinate(view.Bounds.West, value);
            MapCoordinate b = longitude ? new MapCoordinate(value, view.Bounds.North) : new MapCoordinate(view.Bounds.East, value);
            PaintSceneLine(drawing, view.Project(a), view.Project(b), axis.GridColor, axis.GridWidth);
        }
    }

    private static double MapGridStep(double target) {
        double unit = Math.Pow(10, Math.Floor(Math.Log10(target)));
        double scaled = target / unit;
        return unit * (scaled < 2 ? 1 : scaled < 5 ? 2 : 5);
    }

    private static void DrawMapRegions(OfficeDrawing drawing, MapViewport view, MapChoropleth data,
        CancellationToken cancellationToken, List<NotebookChartMark>? marks) {
        foreach (MapRegion region in data.Regions) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeColor fill = WithMapOpacity(data.Scale.At(region.Value), data.Opacity);
            OfficeColor border = WithMapOpacity(data.Border, data.Opacity);
            bool painted = false;
            foreach (MapPolygon polygon in region.Feature.Polygons)
                painted |= PaintMapPolygon(drawing, view, polygon, fill, border, data.BorderWidth);
            if (painted && marks != null && fill.A > 0 && TryFindMapRegionAnchor(view, region.Feature, out MapCoordinate anchor)) {
                string name = region.Label;
                string measure = data.ScaleTitle ?? "Value";
                AddMapInspectMark(marks, view.Project(anchor), 12, name + " • " + measure + ": " + region.Value.ToString("G", CultureInfo.InvariantCulture));
            }
        }
    }

    private static void DrawMapScatter(OfficeDrawing drawing, MapViewport view, MapScatter data, OfficeColor textColor,
        CancellationToken cancellationToken, List<NotebookChartMark>? marks) {
        if (data.Lines) {
            MapCoordinate? previous = null;
            foreach (MapCoordinate? coordinate in data.Coordinates) {
                cancellationToken.ThrowIfCancellationRequested();
                if (coordinate is not MapCoordinate position) {
                    if (!data.ConnectGaps) previous = null;
                    continue;
                }
                if (previous is MapCoordinate start) PaintMapRoute(drawing, view, data.Style, start, position);
                previous = position;
            }
        }

        double size = data.Style.MarkerSize;
        for (int index = 0; index < data.Coordinates.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (data.Coordinates[index] is not MapCoordinate position || !view.Bounds.Contains(position)) continue;
            OfficePoint location = view.Project(position);
            if (data.Markers) {
                if (data.Style.OutlineWidth > 0)
                    PaintMapMarker(drawing, view, location, data.Style.OutlineColor, size + 2 * data.Style.OutlineWidth, data.Style.Shape);
                PaintMapMarker(drawing, view, location, data.Style.MarkerColor, size, data.Style.Shape);
            }
            if (data.Text && data.Labels?[index] is string caption)
                PaintSceneLabel(drawing, caption, location.X, location.Y - (data.Markers ? size / 2 + 8 : 3), Math.Min(116, drawing.Width), 9, textColor);
            if (marks != null) {
                string label = data.Hover?[index] ?? data.Name ?? "Location";
                string coordinate = "longitude " + position.Longitude.ToString("G", CultureInfo.InvariantCulture) +
                    ", latitude " + position.Latitude.ToString("G", CultureInfo.InvariantCulture);
                AddMapInspectMark(marks, location, data.Markers ? Math.Max(size, 12) : 12, label + " • " + coordinate);
            }
        }
    }

    private static void PaintMapMarker(OfficeDrawing drawing, MapViewport view, OfficePoint at, OfficeColor color,
        double diameter, OfficeChartMarkerShape shape) {
        if (color.A == 0) return;
        double radius = diameter / 2;
        double left = at.X - radius;
        double top = at.Y - radius;
        double clipLeft = Math.Max(left, view.Left);
        double clipTop = Math.Max(top, view.Top);
        double clipRight = Math.Min(at.X + radius, view.Left + view.Width);
        double clipBottom = Math.Min(at.Y + radius, view.Top + view.Height);
        if (clipLeft == left && clipTop == top && clipRight == at.X + radius && clipBottom == at.Y + radius) {
            PaintSceneMarker(drawing, at, color, diameter, shape);
        } else {
            var piece = new OfficeDrawing(diameter, diameter);
            PaintSceneMarker(piece, new OfficePoint(radius, radius), color, diameter, shape);
            drawing.AddClippedDrawing(piece, clipLeft, clipTop, OfficeClipPath.Rectangle(clipRight - clipLeft, clipBottom - clipTop),
                left - clipLeft, top - clipTop);
        }
    }

    private static void PaintMapRoute(OfficeDrawing drawing, MapViewport view, MapScatterStyle style, MapCoordinate start, MapCoordinate end) {
        double delta = end.Longitude - start.Longitude;
        // Draw both visible pieces of a shortest longitude segment at the antimeridian.
        if (Math.Abs(delta) > 180) {
            double wrap = delta > 0 ? 360 : -360;
            DrawSegment(start, end with { Longitude = end.Longitude - wrap });
            DrawSegment(start with { Longitude = start.Longitude + wrap }, end);
        } else {
            DrawSegment(start, end);
        }

        void DrawSegment(MapCoordinate a, MapCoordinate b) {
            if (TryClipMapSegment(view.Bounds, a, b, out MapCoordinate left, out MapCoordinate right))
                PaintSceneLine(drawing, view.Project(left), view.Project(right), style.LineColor, style.LineWidth, style.Dash);
        }
    }

    private static void PaintMapLegend(OfficeDrawing drawing, MapLayout layout, List<MapScatter> traces, int count,
        bool bottom, double panelBottom) {
        int index = 0;
        double width = bottom ? (drawing.Width - 30) / Math.Min(count, 3) : 75;
        foreach (MapScatter trace in traces) {
            if (!trace.Legend) continue;
            double x = bottom ? 15 + index % 3 * width : drawing.Width - 79;
            double y = bottom ? panelBottom + 13 + index / 3 * 15 : 52 + index * 12;
            if (y + 6 > drawing.Height) break;
            if (trace.Lines) PaintSceneLine(drawing, new OfficePoint(x, y), new OfficePoint(x + 10, y), trace.Style.LineColor, 1.5);
            if (trace.Markers) PaintSceneMarker(drawing, new OfficePoint(x + 5, y), trace.Style.MarkerColor, 5, trace.Style.Shape);
            string label = trace.Name ?? "Trace " + (index + 1).ToString(CultureInfo.InvariantCulture);
            drawing.AddText(label, x + 13, y - 6, width - 13, 12, new OfficeFontInfo(null, 9), layout.TextColor, shrinkToFit: true);
            index++;
        }
    }

    private static void AddMapInspectMark(List<NotebookChartMark> marks, OfficePoint position, double size, string label) =>
        marks.Add(new NotebookChartMark(position.X - size / 2, position.Y - size / 2, size, size, label));
}
