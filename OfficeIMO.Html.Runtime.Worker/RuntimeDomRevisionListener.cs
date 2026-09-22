using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeDomRevisionListener(Action changed) : IDomMutationListener {
    public void OnMutation(IDocument document, IMutationRecord record) => changed();
}
