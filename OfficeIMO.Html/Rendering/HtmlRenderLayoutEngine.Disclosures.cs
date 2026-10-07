using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HtmlDisclosureState _disclosures = new HtmlDisclosureState();

    /// <summary>
    /// Closed disclosures expose their first summary element only. This filters
    /// direct text as well as elements, before layout or intrinsic measurement;
    /// authored child display values cannot open the disclosure's content slot.
    /// </summary>
    private bool IsClosedDisclosureChild(INode node) => _disclosures.IsClosedChild(node);

    /// <summary>Applies the same ancestor disclosure state to independently collected bookmark text.</summary>
    private bool IsInsideClosedDisclosure(INode node) => _disclosures.IsInsideClosedContent(node);

}
