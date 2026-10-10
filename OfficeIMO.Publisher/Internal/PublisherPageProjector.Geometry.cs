using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private OfficeArtCustomPathProjection? CustomPath(PublisherEscherShape source, double width, double height) {
        if (!OfficeArtCustomPathProjector.HasPath(source.Properties)) return null;
        if (OfficeArtCustomPathProjector.TryProject(source.Properties, width, height,
            _context.AccountPathWork, _context.Token, out OfficeArtCustomPathProjection? path, out OfficeArtCustomPathFailure failure)) {
            _context.Add("PUB_CUSTOM_PATH_RENDERING_UNQUALIFIED",
                "Literal native path coordinates and commands use the shared nonzero path fill and clipping engines; winding, picture masking and Publisher-rendered appearance remain unqualified.",
                OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(source.Id));
            return path;
        }
        string code = failure switch {
            OfficeArtCustomPathFailure.GuideReference => "PUB_CUSTOM_PATH_GUIDES_UNSUPPORTED",
            OfficeArtCustomPathFailure.Command => "PUB_CUSTOM_PATH_COMMAND_UNSUPPORTED",
            OfficeArtCustomPathFailure.PaintGroups => "PUB_CUSTOM_PATH_PAINT_GROUPS_UNSUPPORTED",
            _ => "PUB_CUSTOM_PATH_INVALID"
        };
        string detail = failure switch {
            OfficeArtCustomPathFailure.GuideReference => "Native custom path coordinates refer to unevaluated geometry guides.",
            OfficeArtCustomPathFailure.Command => "The native custom path uses an unassessed escape or client command.",
            OfficeArtCustomPathFailure.PaintGroups => "The native custom path contains separately painted groups that this profile does not reconstruct.",
            OfficeArtCustomPathFailure.InvalidGeometrySpace => "The native custom path has an invalid geometry coordinate space.",
            _ => "The native custom path has an invalid or incomplete vertex/command table."
        };
        _context.Add(code, detail + " Its preset or bounding rectangle is used instead.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
        return null;
    }
}
