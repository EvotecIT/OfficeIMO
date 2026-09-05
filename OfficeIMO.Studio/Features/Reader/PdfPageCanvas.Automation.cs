using Avalonia.Automation.Peers;
using Avalonia.Automation.Provider;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

internal sealed class PdfPageCanvasAutomationPeer : ControlAutomationPeer {
    private readonly PdfPageCanvas _owner;

    internal PdfPageCanvasAutomationPeer(PdfPageCanvas owner) : base(owner) => _owner = owner;

    protected override AutomationControlType GetAutomationControlTypeCore() => AutomationControlType.Document;

    protected override string GetClassNameCore() => nameof(PdfPageCanvas);

    protected override IReadOnlyList<AutomationPeer> GetChildrenCore() {
        PdfPageScene? scene = _owner.Scene;
        if (scene is null) return Array.Empty<AutomationPeer>();
        IStudioLocalizer localizer = StudioLocalization.Current;
        var children = new List<AutomationPeer>();
        string pageText = string.Concat(scene.Interactions.TextRegions.Select(static region => region.Text));
        if (!string.IsNullOrWhiteSpace(pageText)) {
            children.Add(new PdfPageInteractionAutomationPeer(
                _owner,
                AutomationControlType.Text,
                localizer.FormatOrDefault("PdfPage.Automation.Text", "Page text: {0}", pageText),
                null));
        }
        foreach (PdfPageInteractionRegion region in scene.Interactions.Regions.Where(static item => item.Kind != PdfInteractionKind.Text)) {
            children.Add(new PdfPageInvokableAutomationPeer(
                _owner,
                GetControlType(region),
                GetName(localizer, region),
                region));
        }
        return children;
    }

    internal void RefreshChildren() => InvalidateChildren();

    private static AutomationControlType GetControlType(PdfPageInteractionRegion region) => region.Kind switch {
        PdfInteractionKind.Link => AutomationControlType.Hyperlink,
        PdfInteractionKind.FormWidget => AutomationControlType.Edit,
        PdfInteractionKind.Image => AutomationControlType.Image,
        _ => AutomationControlType.Custom
    };

    private static string GetName(IStudioLocalizer localizer, PdfPageInteractionRegion region) => region.Kind switch {
        PdfInteractionKind.Link => localizer.FormatOrDefault("PdfPage.Automation.Link", "Link: {0}", region.Target ?? localizer.Get("Common.Unavailable")),
        PdfInteractionKind.FormWidget => localizer.FormatOrDefault("PdfPage.Automation.FormField", "Form field: {0}", region.FieldName ?? localizer.Get("Common.Unavailable")),
        PdfInteractionKind.Image => localizer.GetOrDefault("PdfPage.Automation.Image", "Image"),
        PdfInteractionKind.Annotation => localizer.FormatOrDefault("PdfPage.Automation.Annotation", "Annotation: {0}", region.Text ?? region.Subtype ?? localizer.Get("Common.Unavailable")),
        _ => region.Text ?? region.Subtype ?? region.Kind.ToString()
    };
}

internal class PdfPageInteractionAutomationPeer : ControlAutomationPeer {
    private readonly AutomationControlType _controlType;
    private readonly string _name;

    internal PdfPageInteractionAutomationPeer(
        PdfPageCanvas owner,
        AutomationControlType controlType,
        string name,
        PdfPageInteractionRegion? region) : base(owner) {
        _controlType = controlType;
        _name = name;
    }

    protected override AutomationControlType GetAutomationControlTypeCore() => _controlType;

    protected override string GetClassNameCore() => "PdfPageInteraction";

    protected override string GetNameCore() => _name;
}

internal sealed class PdfPageInvokableAutomationPeer : PdfPageInteractionAutomationPeer, IInvokeProvider {
    private readonly PdfPageCanvas _owner;
    private readonly PdfPageInteractionRegion _region;

    internal PdfPageInvokableAutomationPeer(
        PdfPageCanvas owner,
        AutomationControlType controlType,
        string name,
        PdfPageInteractionRegion region) : base(owner, controlType, name, region) {
        _owner = owner;
        _region = region;
    }

    public void Invoke() {
        _owner.SelectRegion(_region, activateLink: true);
    }
}
