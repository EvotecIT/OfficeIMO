using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using WColor = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {
    internal static void Example_RegisterCustomParagraphStyle(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with custom style");
        string filePath = Path.Combine(folderPath, "CustomParagraphStyle.docx");

        var custom = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
        custom.Append(new StyleName { Val = "MyStyle" });
        var runProps = new StyleRunProperties();
        runProps.Append(new RunFonts { Ascii = "Courier New" });
        runProps.Append(new WColor { Val = Color.Red.ToHexColor() });
        runProps.Append(new FontSize { Val = "28" });
        custom.Append(runProps);

        WordParagraphStyle.RegisterCustomStyle("MyStyle", custom);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("Hello world").SetStyleId("MyStyle");
            document.Save(openWord);
        }
    }

    internal static void Example_MultipleCustomParagraphStyles(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with multiple custom styles");
        string filePath = Path.Combine(folderPath, "MultipleCustomParagraphStyles.docx");

        var centeredRed = new Style { Type = StyleValues.Paragraph, StyleId = "CenteredRed" };
        centeredRed.Append(new StyleName { Val = "CenteredRed" });
        centeredRed.Append(new StyleParagraphProperties(new Justification { Val = JustificationValues.Center }));
        var centeredRedRun = new StyleRunProperties();
        centeredRedRun.Append(new WColor { Val = "FF0000" });
        centeredRedRun.Append(new Bold());
        centeredRed.Append(centeredRedRun);
        WordParagraphStyle.RegisterCustomStyle("CenteredRed", centeredRed);

        var greenIndented = new Style { Type = StyleValues.Paragraph, StyleId = "GreenIndented" };
        greenIndented.Append(new StyleName { Val = "GreenIndented" });
        greenIndented.Append(new StyleParagraphProperties(new Indentation { Left = "720" }));
        var greenIndentedRun = new StyleRunProperties();
        greenIndentedRun.Append(new WColor { Val = "00AA00" });
        greenIndentedRun.Append(new Italic());
        greenIndented.Append(greenIndentedRun);
        WordParagraphStyle.RegisterCustomStyle("GreenIndented", greenIndented);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("This paragraph is centered and red").SetStyleId("CenteredRed");
            document.AddParagraph("This paragraph is indented and green").SetStyleId("GreenIndented");
            document.Save(openWord);
        }
    }

    internal static void Example_OverrideBuiltInParagraphStyle(string folderPath, bool openWord) {
        Console.WriteLine("[*] Overriding built-in Normal style");
        string filePath = Path.Combine(folderPath, "OverrideNormalStyle.docx");
        var original = WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal) ?? throw new InvalidOperationException("Normal style definition was not found.");

        var custom = new Style { Type = StyleValues.Paragraph, StyleId = "Normal" };
        var run = new StyleRunProperties();
        run.Append(new WColor { Val = "0000FF" });
        run.Append(new Bold());
        custom.Append(run);
        WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, custom);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("Paragraph with overridden Normal style");
            document.Save(openWord);
        }

        WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, original);
    }
}
