using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherDrawingBoundaryTests {
    [Theory]
    [InlineData("Simple.pub", 293U, true)]
    [InlineData("Simple.pub", 293U, false)]
    [InlineData("Sample.pub", 299U, true)]
    [InlineData("Sample.pub", 299U, false)]
    public void Text_and_table_bleed_reaches_the_page_clip_and_survives_copies(string file, uint objectId, bool horizontal) {
        byte[] input = PublisherInputContractTests.Mutate("Escher/EscherStm", bytes => {
            Assert.True(MoveAnchor(bytes, 0, bytes.Length, objectId, horizontal));
        }, file);
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PAGE_BLEED_CLIPPED"
            && diagnostic.Location == "Contents/object/" + objectId);
        OfficeDrawing page = document.Pages[file == "Simple.pub" ? 0 : 1].Drawing;
        foreach (OfficeDrawing copy in new[] { page, page.Clone() }) {
            OfficeDrawingRichText[] frames = PublisherNativeTests.Elements(copy).OfType<OfficeDrawingRichText>()
                .Where(frame => frame.SourceElementIds?.Contains("publisher-object-" + objectId) == true).ToArray();
            Assert.NotEmpty(frames);
            Assert.Contains(frames, frame => (horizontal ? frame.X : frame.Y) < 0);
        }
        string svg = document.ToSvg(file == "Simple.pub" ? 0 : 1);
        Assert.Contains("clip-path=", svg);
        Assert.Contains(file == "Simple.pub" ? "0123456789" : "Table on page 2", svg);
        Assert.Equal(document.Pages.Count, PdfReadDocument.Open(document.ToPdfBytes()).Pages.Count);
    }

    // Keep each native frame's dimensions while moving its left/top edge to -12 points.
    // Publisher anchors are signed EMU relative to the physical page centre.
    private static bool MoveAnchor(byte[] bytes, int start, int end, uint objectId, bool horizontal) {
        for (int offset = start; offset < end;) {
            ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
            int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
            if (kind == 0xF004) {
                int anchor = -1;
                bool matches = false;
                for (int child = content; child < boundary;) {
                    ushort type = BitConverter.ToUInt16(bytes, child + 2);
                    int length = checked((int)BitConverter.ToUInt32(bytes, child + 4));
                    if (type == 0xF010) anchor = child + 8;
                    if (type == 0xF011 && length == 10) matches = BitConverter.ToUInt32(bytes, child + 14) == objectId;
                    child += 8 + length;
                }
                if (matches && anchor >= 0) {
                    int first = anchor + (horizontal ? 6 : 12), second = anchor + (horizontal ? 18 : 24);
                    int origin = horizontal ? -3_932_400 : -5_498_400; // A4 centre plus a 12-point bleed.
                    int delta = origin - BitConverter.ToInt32(bytes, first);
                    PublisherInputContractTests.WriteUInt32(bytes, first, unchecked((uint)origin));
                    PublisherInputContractTests.WriteUInt32(bytes, second, unchecked((uint)(BitConverter.ToInt32(bytes, second) + delta)));
                    return true;
                }
            } else if ((initial & 15) == 15 && MoveAnchor(bytes, content, boundary, objectId, horizontal)) return true;
            offset = boundary + (kind is 0xF000 or 0xF002 && boundary < end ? 4 : 0);
        }
        return false;
    }
}
