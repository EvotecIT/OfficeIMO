using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointTableTextDefaults {
        internal const string Language = "en-US";

        internal static A.TextBody CreateTextBody(string text = "") {
            return new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                CreateParagraph(text));
        }

        internal static A.Paragraph CreateParagraph(string text = "") {
            return new A.Paragraph(
                CreateRun(text),
                new A.EndParagraphRunProperties { Language = Language });
        }

        internal static A.Run CreateRun(string text) {
            return new A.Run(
                new A.RunProperties { Language = Language },
                new A.Text(text ?? string.Empty));
        }
    }
}
