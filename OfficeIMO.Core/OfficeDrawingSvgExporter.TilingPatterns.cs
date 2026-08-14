using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendTilingPattern(StringBuilder sb, OfficeDrawingTilingPattern pattern, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget, SvgNearestNeighborRectangleBudget nearestNeighborRectangleBudget) {
        if (pattern.Opacity <= 0D) return;
        string clipId = idPrefix + "officeimo-pattern-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        string tileClipId = idPrefix + "officeimo-pattern-tile-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        string tileId = idPrefix + "officeimo-pattern-tile-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        OfficeImagePlacement area = pattern.Area;
        sb.Append("<defs><clipPath id=\"").Append(clipId).Append("\"><rect x=\"")
            .Append(Format(area.X)).Append("\" y=\"").Append(Format(area.Y))
            .Append("\" width=\"").Append(Format(area.Width)).Append("\" height=\"")
            .Append(Format(area.Height)).Append("\"/></clipPath><clipPath id=\"").Append(tileClipId)
            .Append("\"><rect x=\"0\" y=\"0\" width=\"").Append(Format(pattern.InnerTile.Width))
            .Append("\" height=\"").Append(Format(pattern.InnerTile.Height))
            .Append("\"/></clipPath><g id=\"").Append(tileId).Append("\" clip-path=\"url(#")
            .Append(tileClipId).Append(")\">");
        tilingExpansionBudget.BeginTile(pattern.MaximumTileCount);
        long descendantExpansion;
        try {
            AppendElements(sb, pattern.InnerTile.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget, nearestNeighborRectangleBudget);
            descendantExpansion = tilingExpansionBudget.EndTile();
        } catch {
            tilingExpansionBudget.CancelTile();
            throw;
        }
        sb.Append("</g></defs><g clip-path=\"url(#")
            .Append(clipId).Append(")\"");
        if (pattern.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(pattern.Opacity)).Append('"');
        sb.Append('>');
        foreach (OfficeTransform transform in pattern.GetTileTransforms(pattern.MaximumTileCount)) {
            cancellationToken.ThrowIfCancellationRequested();
            tilingExpansionBudget.Consume(1L + descendantExpansion);
            sb.Append("<use href=\"#").Append(tileId).Append('"')
                .Append(BuildMatrixTransformAttribute(transform, 0D, 0D)).Append("/>");
        }
        sb.Append("</g>");
    }

    private sealed class SvgTilingExpansionBudget {
        private readonly System.Collections.Generic.Stack<long> _tileScopes = new System.Collections.Generic.Stack<long>();
        private long _maximum;
        private long _count;

        internal void BeginTile(int configuredMaximum) {
            _maximum = Math.Max(_maximum, configuredMaximum);
            _tileScopes.Push(0L);
        }

        internal long EndTile() => _tileScopes.Pop();

        internal void CancelTile() {
            if (_tileScopes.Count > 0) _tileScopes.Pop();
        }

        internal void Consume(long renderedInstances) {
            if (renderedInstances <= 0L ||
                renderedInstances > _maximum ||
                _tileScopes.Count == 0 && _count > _maximum - renderedInstances ||
                _tileScopes.Count > 0 && _tileScopes.Peek() > _maximum - renderedInstances) {
                throw new InvalidOperationException("Vector pattern aggregate expansion exceeds the configured tile-count limit.");
            }

            if (_tileScopes.Count == 0) {
                _count += renderedInstances;
                return;
            }

            long current = _tileScopes.Pop();
            _tileScopes.Push(current + renderedInstances);
        }
    }
}
