#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests {
    /// <summary>Protects strict source geometry validation at the public portable export boundary.</summary>
    public sealed class OfficeDrawingPortableGeometryTests {
        private const string Family = "OfficeIMO Baseline Sans";

        [Theory]
        [InlineData(0, "DRAWING_DIMENSION_INVALID")]
        [InlineData(-1, "DRAWING_DIMENSION_INVALID")]
        [InlineData(double.NaN, "DRAWING_COORDINATE_LIMIT")]
        [InlineData(double.PositiveInfinity, "DRAWING_COORDINATE_LIMIT")]
        [InlineData(double.NegativeInfinity, "DRAWING_COORDINATE_LIMIT")]
        [InlineData(1000001, "DRAWING_COORDINATE_LIMIT")]
        public void MutablePresetDimensionsNeverProducePartialJson(double invalid, string code) {
            foreach (OfficeShapeKind kind in new[] { OfficeShapeKind.Rectangle, OfficeShapeKind.RoundedRectangle, OfficeShapeKind.Ellipse }) {
                foreach (bool width in new[] { true, false }) {
                    OfficeDrawing drawing = Preset(kind);
                    OfficeShape shape = drawing.Shapes[0].Shape;
                    if (width) shape.Width = invalid;
                    else shape.Height = invalid;
                    Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family),
                        out string? json, out IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics), kind + " " + (width ? "width" : "height") + " produced " + json);
                    Assert.Null(json);
                    Assert.Equal(code, Assert.Single(diagnostics).Code);
                    Assert.Equal("elements[0]", diagnostics[0].Location);
                    Assert.Equal(OfficeConversionLossKind.Failure, diagnostics[0].LossKind);
                }
            }
        }

        [Fact]
        public void ValidEllipseRetainsItsFourCubicsAndExactCommandBudget() {
            OfficeDrawing drawing = Preset(OfficeShapeKind.Ellipse);
            Assert.True(OfficeDrawingPortableExporter.TryExportJson(drawing,
                new OfficeDrawingPortableOptions(Family) { MaximumCommands = 6 }, out string? json, out IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics));
            Assert.Empty(diagnostics);
            using JsonDocument document = JsonDocument.Parse(json!);
            JsonElement path = document.RootElement.GetProperty("elements")[0];
            JsonElement commands = path.GetProperty("commands");
            Assert.Equal(6, commands.GetArrayLength());
            Assert.Equal("move", commands[0].GetProperty("op").GetString());
            Assert.Equal(45, commands[0].GetProperty("x").GetDouble());
            Assert.Equal(17, commands[0].GetProperty("y").GetDouble());
            double[] endX = { 25, 5, 25, 45 }, endY = { 27, 17, 7, 17 };
            for (int i = 0; i < 4; i++) {
                Assert.Equal("cubic", commands[i + 1].GetProperty("op").GetString());
                Assert.Equal(endX[i], commands[i + 1].GetProperty("x").GetDouble());
                Assert.Equal(endY[i], commands[i + 1].GetProperty("y").GetDouble());
            }
            Assert.Equal(45, commands[1].GetProperty("x1").GetDouble());
            Assert.Equal(22.522847498307936, commands[1].GetProperty("y1").GetDouble());
            Assert.Equal(36.04569499661587, commands[1].GetProperty("x2").GetDouble());
            Assert.Equal("close", commands[5].GetProperty("op").GetString());
            Assert.Equal("000000", path.GetProperty("fill").GetString());
            Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing,
                new OfficeDrawingPortableOptions(Family) { MaximumCommands = 5 }, out json, out diagnostics));
            Assert.Null(json);
            Assert.Equal("DRAWING_COMMAND_LIMIT", Assert.Single(diagnostics).Code);
        }

        [Theory]
        [InlineData(1200000)]
        [InlineData(-1200000)]
        public void QuadraticSourceControlCannotShrinkPastTheCoordinateLimit(double control) {
            OfficeShape shape = OfficeShape.Path(40, 20, OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.QuadraticBezierTo(control, 10, 20, 20));
            shape.FillColor = OfficeColor.Black;
            OfficeDrawing drawing = new OfficeDrawing(100, 100);
            drawing.AddShape(shape, 5, 7);
            RejectCoordinate(drawing);
        }

        [Fact]
        public void ClipSourceDimensionsCannotShrinkPastTheCoordinateLimit() {
            OfficeDrawing container = new OfficeDrawing(2000000, 100);
            container.AddClippedDrawing(new OfficeDrawing(10, 10), 0, 0, OfficeClipPath.Rectangle(2000000, 20));
            OfficeDrawing drawing = new OfficeDrawing(100, 100);
            drawing.AddEffectDrawing(container, OfficeTransform.Scale(0.1, 0.1));
            RejectCoordinate(drawing);
        }

        [Fact]
        public void ClippedGroupTranslationIsValidatedEvenWithoutChildGeometry() {
            OfficeDrawing drawing = new OfficeDrawing(100, 100);
            drawing.AddClippedDrawing(new OfficeDrawing(10, 10), 1, 0, OfficeClipPath.Rectangle(20, 20), 1000000, 0);
            RejectCoordinate(drawing);
        }

        [Fact]
        public void ValidQuadraticAndScaledClipRetainTheirMappedGeometry() {
            OfficeShape shape = OfficeShape.Path(40, 20, OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.QuadraticBezierTo(30, 10, 20, 20));
            shape.FillColor = OfficeColor.Black;
            OfficeDrawing child = new OfficeDrawing(100, 80);
            child.AddShape(shape, 5, 7);
            OfficeDrawing container = new OfficeDrawing(120, 100);
            container.AddClippedDrawing(child, 7, 9, OfficeClipPath.Rectangle(90, 70), -10, -7);
            OfficeDrawing drawing = new OfficeDrawing(240, 200);
            drawing.AddEffectDrawing(container, OfficeTransform.Scale(2, 2));
            Assert.True(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family),
                out string? json, out IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics));
            Assert.Empty(diagnostics);
            using JsonDocument document = JsonDocument.Parse(json!);
            JsonElement group = document.RootElement.GetProperty("elements")[0];
            Assert.Equal(14, group.GetProperty("clip").GetProperty("x").GetDouble());
            Assert.Equal(18, group.GetProperty("clip").GetProperty("y").GetDouble());
            Assert.Equal(180, group.GetProperty("clip").GetProperty("width").GetDouble());
            Assert.Equal(140, group.GetProperty("clip").GetProperty("height").GetDouble());
            JsonElement commands = group.GetProperty("elements")[0].GetProperty("commands");
            Assert.Equal(4, commands[0].GetProperty("x").GetDouble());
            Assert.Equal(18, commands[0].GetProperty("y").GetDouble());
            Assert.Equal("cubic", commands[1].GetProperty("op").GetString());
            Assert.Equal(44, commands[1].GetProperty("x1").GetDouble());
            Assert.Equal(31.333333333333336, commands[1].GetProperty("y1").GetDouble(), 10);
            Assert.Equal(57.333333333333336, commands[1].GetProperty("x2").GetDouble(), 10);
            Assert.Equal(44.66666666666667, commands[1].GetProperty("y2").GetDouble(), 10);
            Assert.Equal(44, commands[1].GetProperty("x").GetDouble());
            Assert.Equal(58, commands[1].GetProperty("y").GetDouble());
        }

        private static OfficeDrawing Preset(OfficeShapeKind kind) {
            OfficeShape shape = kind == OfficeShapeKind.Ellipse ? OfficeShape.Ellipse(40, 20) :
                kind == OfficeShapeKind.RoundedRectangle ? OfficeShape.RoundedRectangle(40, 20, 5) : OfficeShape.Rectangle(40, 20);
            shape.FillColor = OfficeColor.Black;
            OfficeDrawing drawing = new OfficeDrawing(100, 100);
            return drawing.AddShape(shape, 5, 7);
        }

        private static void RejectCoordinate(OfficeDrawing drawing) {
            Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family),
                out string? json, out IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics), "Invalid source geometry produced " + json);
            Assert.Null(json);
            Assert.Equal("DRAWING_COORDINATE_LIMIT", Assert.Single(diagnostics).Code);
            Assert.StartsWith("elements[0]", diagnostics[0].Location);
        }
    }
}
#endif
