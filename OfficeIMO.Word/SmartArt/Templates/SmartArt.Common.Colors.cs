using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal, strongly-typed ColorsDefinition shared by all our SmartArt layouts.
    internal static class SmartArtCommonColors {
        internal static void PopulateColors(DiagramColorsPart part) {
            var colors = new Dgm.ColorsDefinition {
                UniqueId = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"
            };
            colors.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            colors.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            colors.Append(new Dgm.ColorDefinitionTitle { Val = "" });
            colors.Append(new Dgm.ColorTransformDescription { Val = "" });

            var catLst = new Dgm.ColorTransformCategories();
            catLst.Append(new Dgm.ColorTransformCategory { Type = "accent1", Priority = (UInt32Value)11200U });
            colors.Append(catLst);

            // Provide a small set of labels used by shapes/text.
            colors.Append(MakeLabel("node0", accentFill: true, lineAccent: false));
            colors.Append(MakeLabel("lnNode1", accentFill: true, lineAccent: false));
            colors.Append(MakeLabel("alignNode1", accentFill: true, lineAccent: true));

            part.ColorsDefinition = colors;
        }

        private static Dgm.ColorTransformStyleLabel MakeLabel(string name, bool accentFill, bool lineAccent) {
            var lbl = new Dgm.ColorTransformStyleLabel { Name = name };

            var fill = new Dgm.FillColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            if (accentFill) fill.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            else fill.Append(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });

            var line = new Dgm.LineColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            if (lineAccent) line.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            else line.Append(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });

            lbl.Append(fill);
            lbl.Append(line);
            lbl.Append(new Dgm.EffectColorList());
            lbl.Append(new Dgm.TextLineColorList());
            lbl.Append(new Dgm.TextFillColorList());
            lbl.Append(new Dgm.TextEffectColorList());

            return lbl;
        }
    }
}
