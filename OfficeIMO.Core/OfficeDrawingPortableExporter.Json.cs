using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingPortableExporter {
    private static string Serialize(double width, double height, List<Element> elements, int limit) {
        var output = new StringBuilder();
        output.Append("{\"version\":1,\"width\":").Append(Number(width)).Append(",\"height\":").Append(Number(height)).Append(",\"elements\":");
        WriteElements(output, elements, limit);
        output.Append('}'); CheckJson(output, limit); return output.ToString();
    }
    private static void WriteElements(StringBuilder output, List<Element> elements, int limit) {
        output.Append('[');
        for (int i = 0; i < elements.Count; i++) {
            if (i != 0) output.Append(',');
            Element item = elements[i];
            output.Append("{\"kind\":"); Quoted(output, item.Kind, limit);
            if (item.Kind == "group") {
                if (item.Clip.HasValue) {
                    Rect clip = item.Clip.Value;
                    output.Append(",\"clip\":{\"x\":").Append(Number(clip.X)).Append(",\"y\":").Append(Number(clip.Y))
                        .Append(",\"width\":").Append(Number(clip.Width)).Append(",\"height\":").Append(Number(clip.Height)).Append('}');
                }
                output.Append(",\"elements\":"); WriteElements(output, item.Children!, limit);
            } else if (item.Kind == "text") {
                output.Append(",\"text\":"); Quoted(output, item.Text!, limit);
                output.Append(",\"x\":").Append(Number(item.X)).Append(",\"y\":").Append(Number(item.Y)).Append(",\"font\":"); Quoted(output, item.Font!, limit);
                output.Append(",\"fontSize\":").Append(Number(item.FontSize)).Append(",\"color\":"); Quoted(output, item.Color!, limit);
            } else {
                output.Append(",\"commands\":[");
                for (int c = 0; c < item.Commands!.Count; c++) {
                    if (c != 0) output.Append(',');
                    OfficePathCommand command = item.Commands[c];
                    if (command.Kind == OfficePathCommandKind.Close) output.Append("{\"op\":\"close\"}");
                    else {
                        output.Append("{\"op\":\"").Append(command.Kind == OfficePathCommandKind.MoveTo ? "move" :
                            command.Kind == OfficePathCommandKind.LineTo ? "line" : "cubic").Append('"');
                        if (command.Kind == OfficePathCommandKind.CubicBezierTo)
                            output.Append(",\"x1\":").Append(Number(command.ControlPoint1.X)).Append(",\"y1\":").Append(Number(command.ControlPoint1.Y))
                                .Append(",\"x2\":").Append(Number(command.ControlPoint2.X)).Append(",\"y2\":").Append(Number(command.ControlPoint2.Y));
                        output.Append(",\"x\":").Append(Number(command.Point.X)).Append(",\"y\":").Append(Number(command.Point.Y)).Append('}');
                    }
                    CheckJson(output, limit);
                }
                output.Append(']');
                if (item.Fill != null) { output.Append(",\"fill\":"); Quoted(output, item.Fill, limit); output.Append(",\"fillRule\":"); Quoted(output, item.FillRule, limit); }
                if (item.Stroke != null) {
                    output.Append(",\"stroke\":"); Quoted(output, item.Stroke, limit);
                    output.Append(",\"strokeWidth\":").Append(Number(item.StrokeWidth)).Append(",\"lineCap\":"); Quoted(output, item.LineCap!, limit);
                    output.Append(",\"lineJoin\":"); Quoted(output, item.LineJoin!, limit);
                }
            }
            output.Append('}'); CheckJson(output, limit);
        }
        output.Append(']'); CheckJson(output, limit);
    }
    private static string Number(double value) => value == 0 ? "0" : value.ToString("R", CultureInfo.InvariantCulture);
    private static void Quoted(StringBuilder output, string text, int limit) {
        long required = 2;
        foreach (char c in text) required += c < 0x20 ? 6 : c == '"' || c == '\\' ? 2 : 1;
        if (required > limit - output.Length) Fail("DRAWING_JSON_LIMIT", "The portable JSON exceeds the configured limit.", "elements");
        output.Append('"');
        foreach (char c in text) {
            if (c == '"' || c == '\\') output.Append('\\').Append(c);
            else if (c < 0x20) output.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
            else output.Append(c);
        }
        output.Append('"');
    }
    private static void CheckJson(StringBuilder output, int limit) {
        if (output.Length > limit) Fail("DRAWING_JSON_LIMIT", "The portable JSON exceeds the configured limit.", "elements");
    }
}
