using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word.Legacy;

internal sealed partial class WordPerfectReader {
    private void Graphic6(int[] ids) {
        if (ids.Length == 0) throw new InvalidDataException("A WordPerfect box has no template packet.");
        Packet template = Packet6(ids[0], 0x41);
        IEnumerable<int> candidates = ids.Skip(1).Concat(template.Children);
        Packet? content = null;
        foreach (int id in candidates) {
            _budget.Record();
            if (id == 0) continue;
            if (!_packets.TryGetValue(id, out Packet? packet)) throw new InvalidDataException("A WordPerfect box references a missing packet.");
            if (packet.Type == 0x40) { content = packet; break; }
        }
        if (content == null) { Loss("WORDPERFECT_BOX_CONTENT", "Graphics", "A text, equation or unsupported box content was not reconstructed."); return; }
        bool cached = false;
        foreach (int id in content.Children) {
            _budget.Record();
            if (!_packets.TryGetValue(id, out Packet? packet)) throw new InvalidDataException("A WordPerfect graphic references a missing packet.");
            if (packet.Type != 0x6f) continue; // OLE payloads are kept inert at prefix parsing.
            cached = true;
            _budget.Resource(packet.Length);
            byte[] source = new byte[packet.Length]; Buffer.BlockCopy(_data, packet.Offset, source, 0, source.Length);
            OfficeDrawing drawing;
            try { drawing = OfficeWpgGraphicReader.Read(source, _budget.Record, _budget.Item); }
            catch (NotSupportedException) {
                Loss("WORDPERFECT_GRAPHIC_PROFILE", "Graphics", "The cached WPG graphic is outside the basic WPG1 vector profile; it was not substituted by text or an unrelated image.");
                continue;
            }
            catch (ArgumentException exception) { throw new InvalidDataException("The WPG graphic contains invalid geometry.", exception); }
            int width = checked((int)Math.Ceiling(drawing.Width * 96d / 72d));
            int height = checked((int)Math.Ceiling(drawing.Height * 96d / 72d));
            _budget.Image((long)width * height);
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = 96d / 72d, MaximumRasterPixels = _budget.Limits.MaxImagePixels, CancellationToken = _budget.Cancellation
            });
            if (_budget.RemainingResourceBytes == 0) throw new InvalidDataException("WordPerfect exceeds the configured resource-byte limit.");
            using var encoded = new OfficeBoundedMemoryStream(_budget.RemainingResourceBytes);
            OfficePngWriter.EncodeTo(raster, encoded, _budget.Cancellation);
            byte[] png = encoded.ToArray();
            _budget.Resource(png.Length); _budget.Item();
            Paragraph().Runs.Add(new LegacyWordRun(string.Empty) { Image = new LegacyWordImage {
                SourceBytes = source, PngBytes = png, WidthPoints = drawing.Width, HeightPoints = drawing.Height
            } });
            Loss("WORDPERFECT_GRAPHIC_LAYOUT", "Graphics", "Basic WPG1 vectors are rasterized at 96 DPI and placed inline; source box cropping, scaling and floating placement are not reconstructed.");
            return;
        }
        if (!cached && content.Children.Length == 0)
            Inert("WORDPERFECT_GRAPHIC_REFERENCE_INERT", OfficeLegacyInertContentKind.ExternalLinks,
                "A graphic without cached content was not loaded from its source filename.");
    }
}
