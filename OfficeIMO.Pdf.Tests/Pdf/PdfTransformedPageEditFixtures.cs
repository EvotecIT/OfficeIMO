using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Builds pages whose existing content or geometry is hostile to appended edits, and measures where edits land on
/// the displayed page.
/// </summary>
public static class PdfTransformedPageEditFixtures {
    internal const double PageWidth = 400D;
    internal const double PageHeight = 300D;

    /// <summary>How the original page content leaves the graphics state.</summary>
    public enum ContentShape {
        Balanced,
        UnclosedSave,
        NestedUnclosedSave,
        BareFlip,
        LeadingExtraRestore,
        TrailingExtraRestoreThenFlip,
        SplitStreams,
        StateLeak
    }

    public enum BoxShape {
        Origin,
        OffsetMediaBox,
        InsetCropBox
    }

    internal sealed class Fixture {
        internal Fixture(ContentShape content, int rotation, BoxShape box) {
            Content = content;
            Rotation = rotation;
            Box = box;
            MediaBox = box == BoxShape.OffsetMediaBox
                ? new[] { 100D, 200D, 100D + PageWidth, 200D + PageHeight }
                : new[] { 0D, 0D, PageWidth, PageHeight };
            CropBox = box == BoxShape.InsetCropBox ? new[] { 20D, 10D, PageWidth - 20D, PageHeight - 10D } : null;
            double[] effective = CropBox ?? MediaBox;
            Left = effective[0];
            Bottom = effective[1];
            Right = effective[2];
            Top = effective[3];
            Bytes = Build(this);
        }

        internal ContentShape Content { get; }
        internal int Rotation { get; }
        internal BoxShape Box { get; }
        internal double[] MediaBox { get; }
        internal double[]? CropBox { get; }
        internal double Left { get; }
        internal double Bottom { get; }
        internal double Right { get; }
        internal double Top { get; }
        internal double Width => Right - Left;
        internal double Height => Top - Bottom;
        internal byte[] Bytes { get; }

        /// <summary>Displayed page width and height after /Rotate.</summary>
        internal (double Width, double Height) DisplaySize => Rotation % 180 == 0 ? (Width, Height) : (Height, Width);

        /// <summary>Maps a default user-space point to top-left displayed-page coordinates.</summary>
        internal (double X, double Y) ToDisplay(double x, double y) => Rotation switch {
            90 => (y - Bottom, x - Left),
            180 => (Right - x, y - Bottom),
            270 => (Top - y, Right - x),
            _ => (x - Left, Top - y)
        };

        /// <summary>Maps a top-left displayed-page point to default user space.</summary>
        internal (double X, double Y) ToUser(double x, double y) => Rotation switch {
            90 => (Left + y, Bottom + x),
            180 => (Right - x, Bottom + y),
            270 => (Right - y, Top - x),
            _ => (Left + x, Top - y)
        };

        public override string ToString() => $"{Content}/R{Rotation}/{Box}";
    }

    internal static IEnumerable<Fixture> All() {
        foreach (ContentShape content in Enum.GetValues(typeof(ContentShape)))
            foreach (int rotation in new[] { 0, 90, 180, 270 })
                foreach (BoxShape box in Enum.GetValues(typeof(BoxShape)))
                    yield return new Fixture(content, rotation, box);
    }

    /// <summary>Blue mark near the user-space lower-left corner and green mark near the upper-right corner.</summary>
    internal static (double X, double Y, double W, double H) BlueMark(Fixture f) => (f.Left + 30D, f.Bottom + 30D, 60D, 40D);
    internal static (double X, double Y, double W, double H) GreenMark(Fixture f) => (f.Right - 80D, f.Top - 70D, 50D, 40D);

    private static byte[] Build(Fixture f) {
        // A y-flip about the effective box keeps the marks at the same user-space position while leaving the CTM mirrored.
        double flip = f.Bottom + f.Top;
        var blue = BlueMark(f);
        var green = GreenMark(f);
        string Rect((double X, double Y, double W, double H) r, bool flipped) =>
            flipped ? Fmt(r.X, flip - r.Y - r.H, r.W, r.H) + " re f" : Fmt(r.X, r.Y, r.W, r.H) + " re f";
        string flipCm = "1 0 0 -1 0 " + Fmt(flip) + " cm";
        string marks(bool flipped) => "0 0 1 rg " + Rect(blue, flipped) + "\n0 1 0 rg " + Rect(green, flipped) + "\n";
        string[] streams = f.Content switch {
            ContentShape.Balanced => new[] { "q\n" + marks(false) + "Q\n" },
            ContentShape.UnclosedSave => new[] { "q\n" + flipCm + "\n" + marks(true) },
            ContentShape.NestedUnclosedSave => new[] { "q\n" + flipCm + "\nq\n" + marks(true) },
            ContentShape.BareFlip => new[] { flipCm + "\n" + marks(true) },
            ContentShape.LeadingExtraRestore => new[] { "Q\n" + flipCm + "\n" + marks(true) },
            ContentShape.TrailingExtraRestoreThenFlip => new[] { "q\n" + marks(false) + "Q\nQ\n" + flipCm + "\n" },
            ContentShape.SplitStreams => new[] { "q\n" + flipCm + "\n", marks(true) },
            ContentShape.StateLeak => new[] { marks(false) + "3 Tr 4 Tz /Low gs 0 0 0 RG 9 w " + Fmt(f.Left, f.Bottom, 5, 5) + " re W n\n" },
            _ => throw new ArgumentOutOfRangeException()
        };

        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>"
        };
        int firstStream = 5;
        string contents = streams.Length == 1
            ? firstStream + " 0 R"
            : "[" + string.Join(" ", streams.Select((_, index) => (firstStream + index) + " 0 R")) + "]";
        string page = "<< /Type /Page /Parent 2 0 R /MediaBox [" + Fmt(f.MediaBox) + "]" +
            (f.CropBox is null ? string.Empty : " /CropBox [" + Fmt(f.CropBox) + "]") +
            (f.Rotation == 0 ? string.Empty : " /Rotate " + f.Rotation.ToString(CultureInfo.InvariantCulture)) +
            " /Resources 4 0 R /Contents " + contents + " >>";
        objects.Add(page);
        objects.Add("<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /ExtGState << /Low << /Type /ExtGState /ca 0.05 /CA 0.05 >> >> >>");
        foreach (string stream in streams) {
            objects.Add("<< /Length " + Encoding.ASCII.GetByteCount(stream).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + stream + "\nendstream");
        }
        return Serialize(objects);
    }

    internal static byte[] Serialize(IReadOnlyList<string> objects) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int i = 0; i < objects.Count; i++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    internal static string Fmt(params double[] values) =>
        string.Join(" ", values.Select(value => value.ToString("0.###", CultureInfo.InvariantCulture)));

    internal readonly struct Box {
        internal Box(int left, int top, int right, int bottom, int count) {
            Left = left; Top = top; Right = right; Bottom = bottom; Count = count;
        }
        internal int Left { get; }
        internal int Top { get; }
        internal int Right { get; }
        internal int Bottom { get; }
        internal int Count { get; }
        internal bool IsEmpty => Count == 0;
        internal double CenterX => (Left + Right + 1) / 2D;
        internal double CenterY => (Top + Bottom + 1) / 2D;
        internal int PixelWidth => Right - Left + 1;
        internal int PixelHeight => Bottom - Top + 1;
        public override string ToString() => IsEmpty ? "empty" : $"[{Left},{Top}..{Right},{Bottom}] n={Count}";
    }

    internal enum Ink { Red, Green, Blue, Dark }

    internal static OfficeRasterImage Render(byte[] pdf) {
        (OfficeRasterImage raster, PdfPageRenderResult result) = PdfDocument.Load(pdf).Render.DisplayRaster(1, new PdfPageDisplayOptions { Scale = 1D }, default);
        if (!result.Succeeded) throw new InvalidOperationException(string.Join("; ", result.Diagnostics));
        return raster;
    }

    internal static Box Find(OfficeRasterImage image, Ink ink, int? left = null, int? top = null, int? right = null, int? bottom = null) {
        byte[] pixels = image.GetPixels();
        int minX = int.MaxValue, minY = int.MaxValue, maxX = -1, maxY = -1, count = 0;
        int x0 = Math.Max(0, left ?? 0), y0 = Math.Max(0, top ?? 0);
        int x1 = Math.Min(image.Width - 1, right ?? image.Width - 1), y1 = Math.Min(image.Height - 1, bottom ?? image.Height - 1);
        for (int y = y0; y <= y1; y++) {
            for (int x = x0; x <= x1; x++) {
                int index = (y * image.Width + x) * 4;
                byte r = pixels[index], g = pixels[index + 1], b = pixels[index + 2], a = pixels[index + 3];
                if (a < 128) continue;
                bool match = ink switch {
                    Ink.Red => r > 180 && g < 90 && b < 90,
                    Ink.Green => g > 180 && r < 90 && b < 90,
                    Ink.Blue => b > 180 && r < 90 && g < 90,
                    _ => r < 90 && g < 90 && b < 90
                };
                if (!match) continue;
                count++;
                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }
        }
        return count == 0 ? default : new Box(minX, minY, maxX, maxY, count);
    }

    internal static byte[] RedImagePngWithGreenTop(int width = 40, int height = 20) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++)
            for (int x = 0; x < width; x++)
                image.SetPixel(x, y, y < height / 4 ? OfficeIMO.OfficeColor.FromNormalizedRgb(0, 1, 0) : OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 0));
        return OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png);
    }

    /// <summary>Visual angle, counter-clockwise on the displayed page, of a text span drawn in user space.</summary>
    internal static double DisplayAngle(PdfTextSpan span, int pageRotation) {
        double angle = (span.RotationDegrees - pageRotation) % 360D;
        if (angle < 0D) angle += 360D;
        return angle;
    }
}
