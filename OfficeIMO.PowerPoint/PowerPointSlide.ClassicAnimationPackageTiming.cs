using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static bool HasOnlyPackageClassicTiming(OfficeXmlNode timing, IReadOnlyList<PowerPointClassicAnimation> animations) {
            if (timing.IsNone) return animations.Count == 0;
            if (animations.Count == 0 || !HasSafeClassicTimingChildren(timing)) return false;

            OfficeXmlNode[] nodes = TimingNodesNamed(timing, null).ToArray();
            var effects = nodes.Where(node => node.LocalName == "animEffect").ToArray();
            if (effects.Length != animations.Count) return false;
            for (int index = 0; index < effects.Length; index++) {
                OfficeXmlNode target = TimingNodesNamed(effects[index], "spTgt").FirstOrDefault();
                uint id = TryReadClassicShapeId(target, out uint parsed) ? parsed : uint.MaxValue;
                if (id != animations[index].ShapeId) return false;
            }
            return HasOnlyClassicStandardActions(nodes, animations);
        }

        private static bool HasSafeClassicTimingChildren(OfficeXmlNode timing) {
            var pending = new Stack<OfficeXmlNode>();
            pending.Push(timing);
            while (pending.Count > 0) {
                OfficeXmlNode parent = pending.Pop();
                foreach (OfficeXmlNode child in ClassicChildren(parent)) {
                    if (parent.Is(ClassicPresentationNamespace, "attrName") && child.Tree is { } tree
                        && tree.KindOf(child.Node) == OfficeXmlTreeNodeKind.Text) continue;
                    if (!IsClassicTimingChild(parent, child)) return false;
                    pending.Push(child);
                }
            }
            return true;
        }

        private static bool IsClassicTimingChild(OfficeXmlNode parent, OfficeXmlNode child) {
            if (child.NamespaceId == OfficeXmlNamespaces.DrawingML) {
                return parent.Is(ClassicPresentationNamespace, "to") && parent.Parent.Is(ClassicPresentationNamespace, "animClr")
                    && child.LocalName is ("srgbClr" or "schemeClr");
            }
            if (child.NamespaceId != ClassicPresentationNamespace) return false;
            string name = child.LocalName;
            return parent.LocalName switch {
                "timing" => name is "tnLst" or "bldLst",
                "tnLst" => name == "par",
                "par" => name == "cTn",
                "seq" => name is "cTn" or "prevCondLst" or "nextCondLst",
                "cTn" => name is "stCondLst" or "endCondLst" or "childTnLst",
                "stCondLst" or "endCondLst" or "prevCondLst" or "nextCondLst" => name == "cond",
                "cond" => name is "tn" or "tgtEl",
                "tgtEl" => name is "sldTgt" or "spTgt" or "sndTgt",
                "spTgt" => name == "txEl",
                "txEl" => name is "pRg" or "charRg",
                "childTnLst" => name is "par" or "seq" or "animEffect" or "animClr" or "cmd" or "audio" or "set",
                "animEffect" or "cmd" => name == "cBhvr",
                "animClr" or "set" => name is "cBhvr" or "to",
                "audio" => name == "cMediaNode",
                "cBhvr" => name is "cTn" or "tgtEl" or "attrNameLst",
                "cMediaNode" => name is "cTn" or "tgtEl",
                "attrNameLst" => name == "attrName",
                "to" => parent.Parent.Is(ClassicPresentationNamespace, "set") && name == "strVal",
                "bldLst" => name == "bldP",
                _ => false
            };
        }
    }
}
