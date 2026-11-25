using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Speaker notes read and edited on the document's package while the SDK is not open. Only notes that already have
    /// everything the model would otherwise add on first use (the notes slide, its link back to the slide and the notes
    /// master with its theme) are handled there; the model creates the rest through the SDK.
    /// </summary>
    public partial class PowerPointNotes {
        // The text body of the notes text shape, NoNode when the notes have none; false when the model reads them.
        private bool TryGetPackageNotesBody(out OfficeXmlTree? tree, out int body) {
            tree = null;
            body = OfficeXmlTree.NoNode;
            if (_slide is not { IsPackageBacked: true, RemovedFromPackage: false, PackagePartName: { } slidePart, OwnerDocument: { } document }) return false;
            // Notes the slide doesn't have yet are created, as reading or writing them through the model creates them.
            if ((document.GetPackageNotes(slidePart) ?? document.TryCreatePackageNotes(slidePart)) is not { } notes) return false;
            tree = notes;
            int data = notes.FindChild(notes.Root, OfficeXmlNamespaces.PresentationML, "cSld");
            int shapes = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : notes.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
            if (shapes == OfficeXmlTree.NoNode) return true;
            int shape = FindTextShape(notes, shapes);
            if (shape != OfficeXmlTree.NoNode) body = notes.FindChild(shape, OfficeXmlNamespaces.PresentationML, "txBody");
            return true;
        }

        // The body placeholder, or else the first shape with text, as GetNotesTextShape picks it.
        private static int FindTextShape(OfficeXmlTree notes, int shapes) {
            int withText = OfficeXmlTree.NoNode;
            for (int shape = notes.FindChild(shapes, OfficeXmlNamespaces.PresentationML, "sp"); shape != OfficeXmlTree.NoNode;
                 shape = notes.FindNextSibling(shape, OfficeXmlNamespaces.PresentationML, "sp")) {
                int nonVisual = notes.FindChild(shape, OfficeXmlNamespaces.PresentationML, "nvSpPr");
                int application = nonVisual == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : notes.FindChild(nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
                int placeholder = application == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : notes.FindChild(application, OfficeXmlNamespaces.PresentationML, "ph");
                if (placeholder != OfficeXmlTree.NoNode && notes.GetAttribute(placeholder, OfficeXmlNamespaces.None, "type") == "body") return shape;
                if (withText == OfficeXmlTree.NoNode && notes.FindChild(shape, OfficeXmlNamespaces.PresentationML, "txBody") != OfficeXmlTree.NoNode) withText = shape;
            }
            return withText;
        }

        // Replaces the notes paragraphs as SetParagraphs does, when the notes already have a text shape with a body.
        private bool TrySetPackageText(string? value) {
            if (_slide?.OwnerDocument is not { CanEditPackage: true } document || !TryGetPackageNotesBody(out OfficeXmlTree? notes, out int body)
                || body == OfficeXmlTree.NoNode) {
                return false;
            }
            OfficeXmlTree tree = notes!;
            int template = tree.FindChild(body, OfficeXmlNamespaces.DrawingML, "p");
            int paragraphProperties = OfficeXmlTree.NoNode, endProperties = OfficeXmlTree.NoNode, runProperties = OfficeXmlTree.NoNode;
            if (template != OfficeXmlTree.NoNode) {
                paragraphProperties = tree.FindChild(template, OfficeXmlNamespaces.DrawingML, "pPr");
                endProperties = tree.FindChild(template, OfficeXmlNamespaces.DrawingML, "endParaRPr");
                for (int run = tree.FindChild(template, OfficeXmlNamespaces.DrawingML, "r"); run != OfficeXmlTree.NoNode && runProperties == OfficeXmlTree.NoNode;
                     run = tree.FindNextSibling(run, OfficeXmlNamespaces.DrawingML, "r")) {
                    runProperties = tree.FindChild(run, OfficeXmlNamespaces.DrawingML, "rPr");
                    // Copies of the run's properties would lose prefixes of foreign namespaces the removed run declares.
                    if (runProperties != OfficeXmlTree.NoNode && tree.DeclaresForeignNamespaces(run)) return false;
                }
                if (tree.DeclaresForeignNamespaces(template)) return false;
            }
            for (int paragraph = tree.FindChild(body, OfficeXmlNamespaces.DrawingML, "p"); paragraph != OfficeXmlTree.NoNode;) {
                int next = tree.FindNextSibling(paragraph, OfficeXmlNamespaces.DrawingML, "p");
                tree.Remove(paragraph);
                paragraph = next;
            }
            foreach (string text in PowerPointTextContent.SplitParagraphs(value)) {
                int paragraph = PowerPointTreeText.Append(tree, body, "p");
                if (paragraphProperties != OfficeXmlTree.NoNode) tree.Append(paragraph, tree.CloneElement(paragraphProperties));
                PowerPointTreeText.AppendInlineContent(tree, paragraph, text, runProperties);
                if (endProperties != OfficeXmlTree.NoNode) tree.Append(paragraph, tree.CloneElement(endProperties));
            }
            document.MarkPackageEdited();
            return true;
        }
    }
}
