using System.Linq;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Tests {
    /// <summary>
    ///     AddImage keeps an image's aspect ratio inside the requested box. Tests about rendering an explicit,
    ///     deliberately mismatched frame restore that frame here.
    /// </summary>
    internal static class PowerPointTestPictures {
        internal static PowerPointPicture WithFrame(this PowerPointPicture picture, long left, long top, long width, long height) {
            picture.Left = OfficeLength.Emus(left);
            picture.Top = OfficeLength.Emus(top);
            picture.Width = OfficeLength.Emus(width);
            picture.Height = OfficeLength.Emus(height);
            return picture;
        }
    }

    /// <summary>
    ///     New presentations carry theme text styles, master placeholders and theme colours. Tests about a renderer's
    ///     behaviour for text with no inherited formatting remove those defaults first.
    /// </summary>
    internal static class PowerPointTestThemeDefaults {
        internal static void ClearThemeTextDefaults(this PowerPointDocument presentation) {
            var part = presentation.OpenXmlDocument.PresentationPart!;
            part.Presentation!.DefaultTextStyle = new DocumentFormat.OpenXml.Presentation.DefaultTextStyle();
            foreach (var master in part.SlideMasterParts) {
                master.SlideMaster!.TextStyles = new DocumentFormat.OpenXml.Presentation.TextStyles(
                    new DocumentFormat.OpenXml.Presentation.TitleStyle(), new DocumentFormat.OpenXml.Presentation.BodyStyle(),
                    new DocumentFormat.OpenXml.Presentation.OtherStyle());
                var tree = master.SlideMaster.CommonSlideData!.ShapeTree!;
                foreach (var shape in tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().ToList()) shape.Remove();
            }
        }
    }
}
