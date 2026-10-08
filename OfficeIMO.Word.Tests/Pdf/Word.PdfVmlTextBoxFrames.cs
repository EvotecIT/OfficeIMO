using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("default", true, true, 0.75D)]
    [InlineData("unfilled", true, false, 0.75D)]
    [InlineData("unstroked", false, true, 0D)]
    [InlineData("explicit", true, false, 2D)]
    [InlineData("transparent", false, false, 0D)]
    [InlineData("child-weight", true, true, 2D)]
    [InlineData("child-disabled", false, true, 0D)]
    [InlineData("rectangle-default", true, true, 0.75D)]
    public void VmlTextBoxRetainsItsFrameWithoutAnExplicitFillColor(string variant, bool stroked, bool filled, double width) {
        using WordDocument document = WordDocument.Create();
        var shape = new V.Shape(new V.TextBox(new TextBoxContent(new Paragraph(
            new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" }),
                new Text("Visible textbox")))))) {
            Id = "VmlFrame", Type = "#_x0000_t202",
            Style = "position:absolute;left:72pt;top:72pt;width:360pt;height:120pt"
        };
        if (!filled) shape.Filled = false;
        if (!stroked && variant != "child-disabled") shape.Stroked = false;
        if (variant == "child-disabled") shape.Append(new V.Stroke { On = false });
        if (variant == "child-weight") shape.Append(new V.Stroke { Weight = "2pt" });
        if (variant == "rectangle-default") shape.Type = "#_x0000_t1";
        if (variant == "explicit") {
            shape.StrokeColor = "#FF0000";
            shape.StrokeWeight = "2pt";
        }
        document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(new Paragraph(new Run(new Picture(shape)))));
        document.AddParagraph("Body");
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = pdf.GetPage(1);
        Assert.Contains("Visible textbox", page.Text);
        var frames = page.Paths.Where(path => (path.IsStroked || path.IsFilled) &&
            path.GetBoundingRectangle() is { } bounds &&
            Math.Abs(bounds.Width - 360D) < 0.01D && Math.Abs(bounds.Height - 120D) < 0.01D).ToArray();
        if (!stroked && !filled) {
            Assert.Empty(frames);
            return;
        }
        var frame = Assert.Single(frames);
        Assert.Equal(stroked, frame.IsStroked);
        Assert.Equal(filled, frame.IsFilled);
        if (stroked) Assert.InRange(Math.Abs(frame.LineWidth - width), 0D, 0.01D);
    }

    [Theory]
    [InlineData("parent-child-weight", 2D)]
    [InlineData("type-transparent", 0D)]
    [InlineData("type-colors", 2D)]
    [InlineData("instance-overrides-type", 3D)]
    [InlineData("child-on", 2D)]
    [InlineData("child-colors", 2D)]
    [InlineData("type-child", 2D)]
    [InlineData("instance-over-type-child", 3D)]
    [InlineData("partial-child", 3D)]
    [InlineData("unitless-parent", 2D)]
    [InlineData("unitless-child", 2D)]
    public void VmlFramesResolveChildSettingsAndShapeTypeInheritance(string variant, double expectedWidth) {
        using WordDocument document = WordDocument.Create();
        var shape = new V.Shape(new V.TextBox(new TextBoxContent(new Paragraph(new Run(new Text("Inherited frame")))))) {
            Id = "InheritedFrame", Type = "#_x0000_t202",
            Style = "position:absolute;left:72pt;top:72pt;width:360pt;height:120pt"
        };
        V.Shapetype? definition = null;
        if (variant.StartsWith("type-") || variant is "instance-overrides-type" or "instance-over-type-child" or "partial-child") {
            definition = new V.Shapetype();
            definition.SetAttributes(new[] {
                new OpenXmlAttribute("id", "", "_x0000_t202"),
                new OpenXmlAttribute("coordsize", "", "21600,21600"),
                new OpenXmlAttribute("path", "", "m,l,21600r21600,l21600,xe")
            });
            if (variant == "type-transparent") {
                definition.SetAttribute(new OpenXmlAttribute("filled", "", "f"));
                definition.SetAttribute(new OpenXmlAttribute("stroked", "", "f"));
            } else if (variant is "type-colors" or "instance-overrides-type") {
                definition.SetAttribute(new OpenXmlAttribute("fillcolor", "", "#FFFF00"));
                definition.SetAttribute(new OpenXmlAttribute("strokecolor", "", "#0000FF"));
                definition.SetAttribute(new OpenXmlAttribute("strokeweight", "", "2pt"));
            } else {
                bool enabled = variant != "instance-over-type-child";
                definition.Append(new V.Fill { On = enabled, Color = "#FFFF00" });
                definition.Append(new V.Stroke { On = enabled, Color = "#0000FF", Weight = "2pt" });
            }
        }
        if (variant is "instance-overrides-type" or "instance-over-type-child" or "child-colors") {
            shape.Filled = true; shape.Stroked = true;
            shape.FillColor = "#00FF00"; shape.StrokeColor = "#FF0000";
            shape.StrokeWeight = variant == "child-colors" ? "1pt" : "3pt";
        }
        if (variant == "child-on") { shape.Filled = false; shape.Stroked = false; }
        if (variant is "child-colors" or "child-on") {
            shape.Append(new V.Fill { On = true, Color = "#FFFF00" });
            shape.Append(new V.Stroke { On = true, Color = "#0000FF", Weight = "2pt" });
        }
        if (variant == "parent-child-weight") {
            shape.StrokeWeight = "1pt"; shape.Append(new V.Stroke { Weight = "2pt" });
        }
        if (variant == "partial-child") {
            shape.Append(new V.Fill { Opacity = "0.5" });
            shape.Append(new V.Stroke { Weight = "3pt" });
        }
        if (variant == "unitless-parent") shape.StrokeWeight = "25400";
        if (variant == "unitless-child") shape.Append(new V.Stroke { Weight = "25400" });
        var picture = new Picture();
        if (definition != null) picture.Append(definition);
        picture.Append(shape);
        document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(new Paragraph(new Run(picture))));
        document.AddParagraph("Body");
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = pdf.GetPage(1);
        Assert.Contains("Inherited frame", page.Text);
        var frames = page.Paths.Where(path => (path.IsStroked || path.IsFilled) &&
            path.GetBoundingRectangle() is { } bounds &&
            Math.Abs(bounds.Width - 360D) < 0.01D && Math.Abs(bounds.Height - 120D) < 0.01D).ToArray();
        if (variant == "type-transparent") { Assert.Empty(frames); return; }
        var frame = Assert.Single(frames);
        Assert.True(frame.IsStroked); Assert.True(frame.IsFilled);
        Assert.InRange(Math.Abs(frame.LineWidth - expectedWidth), 0D, 0.01D);
        bool instanceColor = variant is "instance-overrides-type" or "instance-over-type-child";
        bool inheritedColor = variant is "type-colors" or "child-colors" or "child-on" or "type-child" or "partial-child";
        Assert.Equal(instanceColor ? (1D, 0D, 0D) : inheritedColor ? (0D, 0D, 1D) : (0D, 0D, 0D), frame.StrokeColor.ToRGBValues());
        Assert.Equal(instanceColor ? (0D, 1D, 0D) : inheritedColor ? (1D, 1D, 0D) : (1D, 1D, 1D), frame.FillColor.ToRGBValues());
    }
}
