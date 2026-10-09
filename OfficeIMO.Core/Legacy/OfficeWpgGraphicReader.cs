using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Reads the bounded WPG1 basic-vector profile into the shared drawing model.</summary>
internal static class OfficeWpgGraphicReader {
    internal static OfficeDrawing Read(byte[] data, Action record, Action item) {
        if (data.Length < 16 || data[0] != 0xff || data[1] != 'W' || data[2] != 'P' || data[3] != 'C' ||
            data[9] != 0x16 || data[10] != 1 || U16(data, 12) != 0)
            throw new NotSupportedException("Only unencrypted WPG1 graphics are in the basic-vector profile.");
        uint offset = (uint)(data[4] | data[5] << 8 | data[6] << 16 | data[7] << 24);
        if (offset < 16 || offset > data.Length) throw new InvalidDataException("The WPG record offset is outside its data.");
        var palette = new Dictionary<int, OfficeColor>();
        // The first sixteen WPG1 default palette entries. Other entries require an explicit color map.
        int[] colors = { 0x000000, 0x00007f, 0x007f00, 0x007f7f, 0x7f0000, 0x7f007f, 0x7f3f00, 0x7f7f7f,
            0xc0c0c0, 0x0000ff, 0x00ff00, 0x00ffff, 0xff0000, 0xff00ff, 0xffff00, 0xffffff };
        for (int i = 0; i < colors.Length; i++) palette[i] = Rgb(colors[i]);
        OfficeColor? fill = Rgb(0), stroke = Rgb(0);
        double strokeWidth = 0;
        OfficeDrawing? drawing = null;
        for (int at = (int)offset; at < data.Length;) {
            record(); int type = data[at++];
            int size = Variable(data, ref at);
            if (size > data.Length - at) throw new InvalidDataException("A WPG record is truncated.");
            int end = at + size;
            void Need(int minimum) { if (size < minimum) throw new InvalidDataException("A WPG record's fields are truncated."); }
            OfficeColor Color(int index) => palette.TryGetValue(index, out OfficeColor color) ? color :
                throw new NotSupportedException("The WPG color index requires an explicit palette.");
            if (type == 0x0f) {
                Need(6);
                if (drawing != null || U16(data, at + 2) == 0 || U16(data, at + 4) == 0)
                    throw new InvalidDataException("The WPG canvas is invalid or repeated.");
                drawing = new OfficeDrawing(Points(U16(data, at + 2)), Points(U16(data, at + 4)));
            } else if (type == 0x10) {
                if (drawing == null || end != data.Length) throw new InvalidDataException("The WPG end record is misplaced.");
                return drawing;
            } else if (type == 0x19 || type == 0x0a || type == 0x12) {
                // Editor page/grid preferences, comments and output-device preferences do not draw content.
            } else {
                if (drawing == null) throw new InvalidDataException("WPG content precedes its canvas.");
                if (type == 1) {
                    Need(2);
                    if (data[at] > 1) throw new NotSupportedException("Patterned WPG fills are outside the basic-vector profile.");
                    fill = data[at] == 0 ? null : Color(data[at + 1]);
                } else if (type == 2) {
                    Need(4);
                    if (data[at] > 1) throw new NotSupportedException("Dashed WPG strokes are outside the basic-vector profile.");
                    stroke = data[at] == 0 ? null : Color(data[at + 1]); strokeWidth = Points(U16(data, at + 2));
                } else if (type == 0x0e) {
                    Need(4); int first = U16(data, at), count = U16(data, at + 2);
                    if (first + count > 256 || count * 3 > size - 4) throw new InvalidDataException("The WPG palette is malformed.");
                    for (int i = 0; i < count; i++) { record(); int p = at + 4 + i * 3; palette[first + i] = Rgb(data[p] << 16 | data[p + 1] << 8 | data[p + 2]); }
                } else {
                    item(); OfficeShape shape; double x, y;
                    if (type == 5 || type == 6 || type == 8) {
                        Need(type == 5 ? 8 : 2);
                        int count = type == 5 ? 2 : U16(data, at), start = type == 5 ? at : at + 2;
                        if (count < (type == 8 ? 3 : 2) || count * 4 > end - start) throw new InvalidDataException("The WPG point directory is malformed.");
                        var points = new List<OfficePoint>(count);
                        for (int i = 0; i < count; i++) { record(); item(); points.Add(new OfficePoint(Points(S16(data, start + i * 4)), drawing.Height - Points(S16(data, start + i * 4 + 2)))); }
                        x = points.Min(point => point.X); y = points.Min(point => point.Y);
                        if (type == 8) shape = OfficeShape.Polygon(points);
                        else if (count == 2) shape = OfficeShape.Line(points[0], points[1]);
                        else {
                            double width = points.Max(point => point.X) - x, height = points.Max(point => point.Y) - y;
                            if (width == 0 || height == 0) {
                                if (points.Any(point => point.X < 0 || point.X > drawing.Width || point.Y < 0 || point.Y > drawing.Height))
                                    throw new NotSupportedException("WPG polyline clipping is outside the basic-vector profile.");
                                // A stroked path can occupy one axis while retaining a positive canvas.
                                x = 0; y = 0; width = drawing.Width; height = drawing.Height;
                            }
                            var commands = points.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point.X - x, point.Y - y) : OfficePathCommand.LineTo(point.X - x, point.Y - y)).ToArray();
                            shape = OfficeShape.Path(width, height, commands);
                        }
                        if (type != 8) shape.FillColor = null;
                    } else if (type == 7) {
                        Need(8); x = Points(S16(data, at)); double width = Points(S16(data, at + 4)), height = Points(S16(data, at + 6));
                        y = drawing.Height - Points(S16(data, at + 2)) - height;
                        shape = OfficeShape.Rectangle(width, height);
                    } else if (type == 9) {
                        Need(10); double rx = Points(S16(data, at + 4)), ry = Points(S16(data, at + 6));
                        if (S16(data, at + 8) != 0) throw new NotSupportedException("Rotated WPG ellipses are outside the basic-vector profile.");
                        x = Points(S16(data, at)) - rx; y = drawing.Height - Points(S16(data, at + 2)) - ry;
                        shape = OfficeShape.Ellipse(rx * 2, ry * 2);
                    } else throw new NotSupportedException("WPG record 0x" + type.ToString("X2") + " is outside the basic-vector profile.");
                    if (type == 7 || type == 8 || type == 9) shape.FillColor = fill;
                    shape.StrokeColor = stroke; shape.StrokeWidth = strokeWidth;
                    if (shape.Width < 0 || shape.Height < 0 || x < 0 || y < 0 || x + shape.Width > drawing.Width || y + shape.Height > drawing.Height)
                        throw new NotSupportedException("WPG geometry extending outside the canvas requires a separately qualified clipping profile.");
                    drawing.AddShape(shape, x, y);
                }
            }
            at = end;
        }
        throw new InvalidDataException("The WPG graphic has no end record.");
    }

    private static int Variable(byte[] data, ref int at) {
        if (at >= data.Length) throw new InvalidDataException("The WPG record length is missing.");
        int length = data[at++];
        if (length != 255) return length;
        length = U16(data, at); at += 2;
        if ((length & 0x8000) == 0) return length;
        int low = U16(data, at); at += 2;
        return checked((length & 0x7fff) * 65536 + low);
    }
    private static int U16(byte[] data, int at) {
        if (at < 0 || at > data.Length - 2) throw new InvalidDataException("WPG data is truncated.");
        return data[at] | data[at + 1] << 8;
    }
    private static int S16(byte[] data, int at) => (short)U16(data, at);
    private static double Points(int value) => value * 72d / 1200d;
    private static OfficeColor Rgb(int color) => OfficeColor.FromRgb((byte)(color >> 16), (byte)(color >> 8), (byte)color);
}
