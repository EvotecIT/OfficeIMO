using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal, strongly-typed StyleDefinition shared by our SmartArt layouts.
    internal static class SmartArtCommonStyle {
        internal static void PopulateStyle(DiagramStylePart part) {
            var style = new Dgm.StyleDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" };
            style.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            style.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            style.Append(new Dgm.StyleDefinitionTitle { Val = "" });
            style.Append(new Dgm.StyleLabelDescription { Val = "" });
            var displayCats = new Dgm.StyleDisplayCategories();
            displayCats.Append(new Dgm.StyleDisplayCategory { Type = "simple", Priority = (UInt32Value)10100U });
            style.Append(displayCats);

            // A minimal label for nodes with theme-based line/fill/effect/font.
            var label = new Dgm.StyleLabel { Name = "node" };
            var dgmStyle = new Dgm.Style();
            var lnRef = new A.LineReference { Index = 2U };
            lnRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var fillRef = new A.FillReference { Index = 1U };
            fillRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var effRef = new A.EffectReference { Index = 0U };
            effRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var fontRef = new A.FontReference { Index = A.FontCollectionIndexValues.Minor };
            fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });
            dgmStyle.Append(lnRef);
            dgmStyle.Append(fillRef);
            dgmStyle.Append(effRef);
            dgmStyle.Append(fontRef);
            label.Append(dgmStyle);

            style.Append(label);
            part.StyleDefinition = style;
        }
    }
}
