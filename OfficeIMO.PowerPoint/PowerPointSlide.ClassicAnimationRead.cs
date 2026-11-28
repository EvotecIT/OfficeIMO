using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const int ClassicPresentationNamespace = OfficeXmlNamespaces.PresentationML;

        private IReadOnlyList<PowerPointClassicAnimation> ReadClassicAnimations() {
            var part = PartRef;
            OfficeXmlNode slide = part.Root;
            OfficeXmlNode metadata = ReadClassicAnimationMetadata(slide);
            if (metadata.IsNone) {
                OfficeXmlNode timing = slide.Child(ClassicPresentationNamespace, "timing");
                if (timing.IsNone) return Array.Empty<PowerPointClassicAnimation>();
                // The SDK includes comments and processing instructions when counting some timing children.
                if (slide.Tree is { } tree && tree.HasCommentsOrInstructions()) {
                    timing = part.ForMovingXmlNodes().Root.Child(ClassicPresentationNamespace, "timing");
                }
                return InferClassicAnimations(timing);
            }

            var animations = new List<PowerPointClassicAnimation>();
            foreach (OfficeXmlNode element in metadata.Children()) {
                if (!element.Is(ClassicAnimationNamespace, "animation")) continue;
                IReadOnlyList<OfficeXmlAttribute> attributes = element.Attributes();
                if (!TryReadClassicUInt(attributes, "shapeId", out uint shapeId)
                    || !TryReadClassicByte(attributes, "effect", out byte effect)
                    || !TryReadClassicByte(attributes, "direction", out byte direction)
                    || !TryReadClassicByte(attributes, "build", out byte build)
                    || !TryReadClassicInt(attributes, "delay", out int delay)
                    || !TryReadClassicInt(attributes, "order", out int order)
                    || !TryReadClassicByte(attributes, "after", out byte after)
                    || !TryReadClassicByte(attributes, "textBuild", out byte textBuild)
                    || !TryReadClassicUInt(attributes, "dimColor", out uint dimColor)
                    || !IsClassicAnimationValue(effect, direction, build, delay, after, textBuild)) continue;

                OfficeXmlNode sound = element.Child(ClassicPresentationNamespace, "snd");
                IReadOnlyList<OfficeXmlAttribute>? soundAttributes = sound.IsNone ? null : sound.Attributes();
                animations.Add(new PowerPointClassicAnimation(shapeId,
                    (PowerPointClassicAnimationEffect)effect, direction, (PowerPointClassicAnimationBuildType)build,
                    ReadClassicBool(attributes, "automatic"), delay, order,
                    ReadClassicBool(attributes, "reverse"), ReadClassicBool(attributes, "animateBackground"),
                    (PowerPointClassicAnimationAfterEffect)after, (PowerPointClassicTextBuild)textBuild, dimColor,
                    ReadClassicBool(attributes, "playsSound"), ReadClassicBool(attributes, "stopsSound"),
                    ClassicAttribute(soundAttributes, "embed"), ClassicAttribute(soundAttributes, "name")));
            }
            return animations.OrderBy(static animation => animation.Order).ToArray();
        }

        private static OfficeXmlNode ReadClassicAnimationMetadata(OfficeXmlNode slide) {
            const int p = ClassicPresentationNamespace;
            foreach (OfficeXmlNode extension in slide.Child(p, "extLst").Children(p, "ext")) {
                if (extension.Attribute("uri") == ClassicAnimationExtensionUri) return extension.Child(ClassicAnimationNamespace, "classicAnimations");
            }
            return default;
        }

        private static bool TryReadClassicUInt(IReadOnlyList<OfficeXmlAttribute> attributes, string name, out uint value) =>
            uint.TryParse(ClassicAttribute(attributes, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool TryReadClassicByte(IReadOnlyList<OfficeXmlAttribute> attributes, string name, out byte value) =>
            byte.TryParse(ClassicAttribute(attributes, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool TryReadClassicInt(IReadOnlyList<OfficeXmlAttribute> attributes, string name, out int value) =>
            int.TryParse(ClassicAttribute(attributes, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool ReadClassicBool(IReadOnlyList<OfficeXmlAttribute> attributes, string name) =>
            bool.TryParse(ClassicAttribute(attributes, name), out bool value) && value;

        private static string? ClassicAttribute(IReadOnlyList<OfficeXmlAttribute>? attributes, string name) {
            if (attributes == null) return null;
            foreach (OfficeXmlAttribute attribute in attributes) {
                if (attribute.LocalName == name) return attribute.Value;
            }
            return null;
        }
    }
}
