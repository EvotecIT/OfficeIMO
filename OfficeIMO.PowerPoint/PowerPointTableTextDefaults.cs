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
            var paragraph = new A.Paragraph(PowerPointTextContent.CreateInlineContent(
                text, new A.RunProperties { Language = Language }));
            paragraph.Append(new A.EndParagraphRunProperties { Language = Language });
            return paragraph;
        }

        internal static A.Run CreateRun(string text) {
            return new A.Run(
                new A.RunProperties { Language = Language },
                new A.Text(text ?? string.Empty));
        }
    }
}
