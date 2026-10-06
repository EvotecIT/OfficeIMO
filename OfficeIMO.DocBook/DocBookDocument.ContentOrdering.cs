using System.Xml.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    /// <summary>Inserts component body content before its subdivisions, preserving existing sibling order.</summary>
    internal void AddBodyElement(XElement parent, XElement element) {
        DocBookNodeKind kind = DocBookNames.GetKind(element.Name, Namespace);
        if (IsSupportedComponent(parent) && IsFlowTypedChild(kind)) {
            XNode? last = parent.LastNode;
            while (last != null && last is not XElement) last = last.PreviousNode;
            // Valid components place subdivisions last. Avoid scanning growing, body-only
            // components on every append.
            if (last is not XElement lastElement || DocBookNames.GetKind(lastElement.Name, Namespace) is not (DocBookNodeKind.Section or DocBookNodeKind.Index)) {
                parent.Add(element);
                return;
            }
            // Walk the trailing subdivisions instead of the growing body prefix.
            XElement subdivision = lastElement;
            for (XNode? previous = lastElement.PreviousNode; previous != null; previous = previous.PreviousNode) {
                if (previous is not XElement previousElement) continue;
                if (DocBookNames.GetKind(previousElement.Name, Namespace) is not (DocBookNodeKind.Section or DocBookNodeKind.Index)) break;
                subdivision = previousElement;
            }
            subdivision.AddBeforeSelf(element);
            return;
        }
        parent.Add(element);
    }
}
