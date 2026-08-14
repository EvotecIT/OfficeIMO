using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlCssRunningStringAssignment CaptureRunningElement(
        IElement element,
        string name,
        double containingWidth,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        double orderOffset = 0D) {
        HtmlRenderBoxStyle captureStyle = style.Clone();
        captureStyle.Position = "static";
        captureStyle.ZIndex = "auto";
        HtmlRenderFlowBlock snapshot = LayoutElement(element, containingWidth, captureStyle, parentStyle, depth);
        int snapshotId = ++_nextRunningElementSnapshotId;
        _runningElementSnapshots[snapshotId] = new HtmlCssRunningElementSnapshot(snapshot, element, parentStyle, depth);
        return new HtmlCssRunningStringAssignment(
            HtmlCssRunningElementKeys.ForName(name),
            HtmlCssRunningElementParser.FormatSnapshotId(snapshotId),
            0D,
            orderOffset);
    }
}
