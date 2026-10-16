using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

internal static class MarkerImageContract {
    private static readonly OfficeColor MarkerColor = OfficeColor.FromRgb(231, 17, 165);

    internal static byte[] CreateFixture() => OfficePngWriter.Encode(new OfficeRasterImage(5, 5, MarkerColor));

    internal static void Verify(HtmlRenderDocument rendered, string svg, byte[] png, byte[] markerPng) {
        HtmlRenderImage[] markers = Visuals(rendered.Pages[0].Scene).OfType<HtmlRenderImage>()
            .Where(image => image.Bytes.SequenceEqual(markerPng)).ToArray();
        if (markers.Length != 2 || markers.Any(image => image.Width != 5D || image.Height != 5D)
            || !OfficePngReader.TryDecode(markerPng, out OfficeRasterImage? source)
            || source!.GetPixel(2, 2) != MarkerColor) {
            throw new InvalidOperationException("The NativeAOT list and footnote markers lost their generated PNG sources.");
        }
        if (!OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out _)
            || CountMarkerIslands(OfficeDrawingRasterRenderer.Render(drawing!, 2D, OfficeColor.White)) != 2) {
            throw new InvalidOperationException("The NativeAOT SVG lost visible list or footnote image-marker paint.");
        }
        if (!OfficePngReader.TryDecode(png, out OfficeRasterImage? raster) || CountMarkerIslands(raster!) != 2) {
            throw new InvalidOperationException("The NativeAOT PNG lost visible list or footnote image-marker paint.");
        }
    }

    private static int CountMarkerIslands(OfficeRasterImage image) {
        var visited = new bool[checked(image.Width * image.Height)];
        var pending = new Queue<int>();
        int islands = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                int start = y * image.Width + x;
                if (visited[start] || image.GetPixel(x, y) != MarkerColor) continue;
                visited[start] = true;
                pending.Enqueue(start);
                int pixels = 0;
                while (pending.Count > 0) {
                    int pixel = pending.Dequeue();
                    pixels++;
                    int px = pixel % image.Width, py = pixel / image.Width;
                    Visit(px - 1, py); Visit(px + 1, py); Visit(px, py - 1); Visit(px, py + 1);
                }
                if (pixels >= 4) islands++;
            }
        }
        return islands;

        void Visit(int x, int y) {
            if (x < 0 || y < 0 || x >= image.Width || y >= image.Height) return;
            int pixel = y * image.Width + x;
            if (visited[pixel] || image.GetPixel(x, y) != MarkerColor) return;
            visited[pixel] = true;
            pending.Enqueue(pixel);
        }
    }

    private static IEnumerable<HtmlRenderVisual> Visuals(IEnumerable<HtmlRenderVisual> source) {
        foreach (HtmlRenderVisual visual in source) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Visuals(children)) yield return child;
        }
    }
}
