using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendTilingPattern(StringBuilder sb, OfficeDrawingTilingPattern pattern, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget) {
        if (pattern.Opacity <= 0D) return;
        tilingExpansionBudget.EnterPattern();
        try {
            string clipId = idPrefix + "officeimo-pattern-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
            string tileClipId = idPrefix + "officeimo-pattern-tile-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
            OfficeImagePlacement area = pattern.Area;
            sb.Append("<defs><clipPath id=\"").Append(clipId).Append("\"><rect x=\"")
                .Append(Format(area.X)).Append("\" y=\"").Append(Format(area.Y))
                .Append("\" width=\"").Append(Format(area.Width)).Append("\" height=\"")
                .Append(Format(area.Height)).Append("\"/></clipPath><clipPath id=\"").Append(tileClipId)
                .Append("\"><rect x=\"0\" y=\"0\" width=\"").Append(Format(pattern.InnerTile.Width))
                .Append("\" height=\"").Append(Format(pattern.InnerTile.Height))
                .Append("\"/></clipPath></defs><g clip-path=\"url(#")
                .Append(clipId).Append(")\"");
            if (pattern.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(pattern.Opacity)).Append('"');
            sb.Append('>');
            foreach (OfficeTransform transform in pattern.GetTileTransforms(pattern.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                tilingExpansionBudget.Consume(pattern.MaximumTileCount);
                sb.Append("<g").Append(BuildMatrixTransformAttribute(transform, 0D, 0D))
                    .Append("><g clip-path=\"url(#").Append(tileClipId).Append(")\">");
                AppendElements(sb, pattern.InnerTile.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget);
                sb.Append("</g></g>");
            }
            sb.Append("</g>");
        } finally {
            tilingExpansionBudget.ExitPattern();
        }
    }

    private sealed class SvgTilingExpansionBudget {
        private int _maximum;
        private int _count;
        private int _depth;

        internal void EnterPattern() {
            if (_depth == 0) {
                _maximum = 0;
                _count = 0;
            }
            _depth++;
        }

        internal void ExitPattern() {
            _depth--;
            if (_depth == 0) {
                _maximum = 0;
                _count = 0;
            }
        }

        internal void Consume(int configuredMaximum) {
            _maximum = Math.Max(_maximum, configuredMaximum);
            if (++_count > _maximum) {
                throw new InvalidOperationException("Vector pattern aggregate expansion exceeds the configured tile-count limit.");
            }
        }
    }
}
