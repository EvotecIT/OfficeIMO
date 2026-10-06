using System.Buffers.Binary;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathNumericFaceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NumericFaceSelectsTheSameAdvanceAndMathConstantsAsTheOnlyMatchingFace(bool scripts) {
        byte[] regular = ManagedTextShapingTestAssets.CreateMathFont();
        byte[] medium = ManagedTextShapingTestAssets.CreateMathFont(table => {
            table[10] = 0; table[11] = 45;
            table[12] = 0; table[13] = 35;
        });
        SetAdvance(medium, 900);
        var selected = new OfficeFontFaceDescriptor(500, 100D);
        OfficeMathRenderOptions Options(bool competingFace) {
            var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Scoped Math", 20D, selected), Padding = 0D };
            options.Fonts.Add("Scoped Math", medium, selected);
            if (competingFace) options.Fonts.Add("Scoped Math", regular, new OfficeFontFaceDescriptor(400, 100D));
            return options;
        }
        OfficeMathExpression expression = scripts
            ? OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2"))
            : OfficeMath.Identifier("x");
        OfficeDrawing expected = OfficeMathRenderer.Render(expression, Options(false));
        OfficeDrawing actual = OfficeMathRenderer.Render(expression, Options(true));
        Assert.Equal(expected.Width, actual.Width, 6);
        Assert.Equal(expected.Height, actual.Height, 6);
        OfficeDrawingText[] expectedText = expected.Elements.OfType<OfficeDrawingText>().ToArray();
        OfficeDrawingText[] actualText = actual.Elements.OfType<OfficeDrawingText>().ToArray();
        Assert.Equal(expectedText.Length, actualText.Length);
        for (int i = 0; i < actualText.Length; i++) {
            Assert.Equal(selected, actualText[i].Font.Face);
            Assert.Equal(expectedText[i].Font.Size, actualText[i].Font.Size, 6);
            Assert.Equal(expectedText[i].TextAdvanceWidth!.Value, actualText[i].TextAdvanceWidth!.Value, 6);
        }
        Assert.Equal(18D, Assert.Single(actualText, text => text.Text == "x").TextAdvanceWidth!.Value, 6);
        if (scripts) Assert.Equal(9D, Assert.Single(actualText, text => text.Text == "2").Font.Size, 6);
    }

    [Fact]
    public void NumericFaceUsesItsDesignedOperatorVariantRatherThanRegularFaceFallback() {
        var selected = new OfficeFontFaceDescriptor(500, 100D);
        var options = new OfficeMathRenderOptions {
            Font = new OfficeFontInfo("Scoped Math", 20D, selected), Padding = 0D
        };
        options.Fonts.Add("Scoped Math", ManagedTextShapingTestAssets.CreateMathConstructionFont(), selected);
        options.Fonts.Add("Scoped Math", ManagedTextShapingTestAssets.CreateMathConstructionFont(math => {
            math[8] = 255; math[9] = 255; // Regular face has constants but no usable variants.
        }), new OfficeFontFaceDescriptor(400, 100D));
        OfficeDrawing drawing = OfficeMathRenderer.Render(OfficeMath.Operator("∑"), options);
        Assert.Equal(44D, Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>()).Shape.Height, 6);
        Assert.Equal(selected, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Font.Face);
    }

    private static void SetAdvance(byte[] font, ushort advance) {
        int count = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
        for (int i = 0; i < count; i++) {
            int record = 12 + i * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "hmtx") continue;
            int offset = (int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(record + 8));
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(offset), advance);
            return;
        }
        throw new InvalidOperationException("The synthetic font must contain horizontal metrics.");
    }
}
