using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed class OneNotePageSemanticsTests {
    [Fact]
    public void NamedPageSizesWriteCanonicalDimensionsForTheirOrientation() {
        var section = new OneNoteSection { Name = "Page geometry" };
        section.Pages.Add(new OneNotePage {
            Title = "Portrait",
            PageSize = OneNotePageSize.Letter,
            Orientation = OneNotePageOrientation.Portrait,
            Width = 999,
            Height = 888
        });
        section.Pages.Add(new OneNotePage {
            Title = "Landscape",
            PageSize = OneNotePageSize.Letter,
            Orientation = OneNotePageOrientation.Landscape,
            Width = 1,
            Height = 1
        });

        OneNoteSection result = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));

        Assert.Equal(17D, result.Pages[0].Width!.Value, 6);
        Assert.Equal(22D, result.Pages[0].Height!.Value, 6);
        Assert.Equal(22D, result.Pages[1].Width!.Value, 6);
        Assert.Equal(17D, result.Pages[1].Height!.Value, 6);
    }

    [Fact]
    public void CustomPageSizeRequiresBothNativeDimensions() {
        var section = new OneNoteSection { Name = "Custom page geometry" };
        section.Pages.Add(new OneNotePage {
            Title = "Incomplete custom page",
            PageSize = OneNotePageSize.Custom,
            Width = 17
        });

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_PAGE_SIZE_DIMENSIONS", error.Code);
    }

    [Fact]
    public void RoundTripsPageLayoutAndBackgroundPrintoutMetadata() {
        var section = new OneNoteSection { Name = "Page semantics" };
        var page = new OneNotePage {
            Title = "Landscape printout",
            Width = 22,
            Height = 17,
            PageSize = OneNotePageSize.Letter,
            Orientation = OneNotePageOrientation.Landscape,
            RightToLeft = false,
            IsReadOnly = true,
            ResolveChildCollisions = true
        };
        page.Margins.Left = 1.5;
        page.Margins.Right = 1.5;
        page.Margins.Top = 1;
        page.Margins.Bottom = 1;
        page.Margins.OriginX = 1.5;
        page.Margins.OriginY = 1;
        var background = new OneNoteImage {
            FileName = "printout.png",
            MediaType = "image/png",
            AltText = "Printed page",
            OcrText = "Recognized printout text",
            OcrLanguageId = 1033,
            DisplayedPageNumber = 4,
            IsBackground = true,
            SizeSetByUser = true,
            UploadState = 0,
            WidthHalfInches = 22,
            HeightHalfInches = 17,
            Layout = new OneNoteLayout { X = 0, Y = 0, Width = 22, Height = 17 },
            Payload = OneNoteBinaryPayload.FromBytes(OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.White)))
        };
        page.DirectContent.Add(background);
        var outline = new OneNoteOutline {
            Layout = new OneNoteLayout {
                X = 1.5,
                Y = 2,
                Width = 8,
                MinimumWidth = 4.5,
                AlignmentInParent = 0x0009000C,
                AlignmentSelf = 12,
                CollisionPriority = 2,
                Tight = true,
                TightAlignment = true
            }
        };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Overlay" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        OneNoteWriteObject imageNode = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidImageNode);
        Assert.Single(imageNode.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.LanguageId && property.Scalar == 1033UL);
        Assert.DoesNotContain(imageNode.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.RichEditTextLanguageId);

        OneNoteSection result = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage actual = Assert.Single(result.Pages);
        OneNoteImage actualBackground = Assert.IsType<OneNoteImage>(Assert.Single(actual.DirectContent));
        OneNoteLayout actualLayout = Assert.Single(actual.Outlines).Layout!;

        Assert.Equal(OneNotePageSize.Letter, actual.PageSize);
        Assert.Equal(OneNotePageOrientation.Landscape, actual.Orientation);
        Assert.True(actual.IsReadOnly);
        Assert.True(actual.ResolveChildCollisions);
        Assert.Equal(1.5, actual.Margins.Left);
        Assert.Equal(1, actual.Margins.OriginY);
        Assert.True(actualBackground.IsBackground);
        Assert.Equal("Recognized printout text", actualBackground.OcrText);
        Assert.Equal(1033U, actualBackground.OcrLanguageId);
        Assert.Equal(4U, actualBackground.DisplayedPageNumber);
        Assert.True(actualBackground.SizeSetByUser);
        Assert.Equal(4.5, actualLayout.MinimumWidth);
        Assert.Equal(0x0009000CU, actualLayout.AlignmentInParent);
        Assert.Equal(12U, actualLayout.AlignmentSelf);
        Assert.Equal(2U, actualLayout.CollisionPriority);
        Assert.True(actualLayout.TightAlignment);
    }
}
