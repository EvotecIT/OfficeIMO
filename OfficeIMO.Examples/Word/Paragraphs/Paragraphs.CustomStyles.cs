using System;
using System.IO;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {
    internal static void Example_RegisterCustomParagraphStyle(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with custom style");
        string filePath = Path.Combine(folderPath, "CustomParagraphStyle.docx");

        var custom = new WordParagraphStyleDefinition("MyStyle") {
            FontName = "Courier New",
            ColorHex = Color.Red.ToRgbHex(),
            FontSizePoints = 14
        };

        WordParagraphStyle.RegisterCustomStyle("MyStyle", custom);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("Hello world").SetStyleId("MyStyle");
            document.Save();
            if (openWord) document.OpenInApplication();
        }
    }

    internal static void Example_MultipleCustomParagraphStyles(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with multiple custom styles");
        string filePath = Path.Combine(folderPath, "MultipleCustomParagraphStyles.docx");

        var centeredRed = new WordParagraphStyleDefinition("CenteredRed") {
            Alignment = WordParagraphAlignment.Center,
            ColorHex = "FF0000",
            Bold = true
        };
        WordParagraphStyle.RegisterCustomStyle("CenteredRed", centeredRed);

        var greenIndented = new WordParagraphStyleDefinition("GreenIndented") {
            LeftIndentTwips = 720,
            ColorHex = "00AA00",
            Italic = true
        };
        WordParagraphStyle.RegisterCustomStyle("GreenIndented", greenIndented);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("This paragraph is centered and red").SetStyleId("CenteredRed");
            document.AddParagraph("This paragraph is indented and green").SetStyleId("GreenIndented");
            document.Save();
            if (openWord) document.OpenInApplication();
        }
    }

    internal static void Example_OverrideBuiltInParagraphStyle(string folderPath, bool openWord) {
        Console.WriteLine("[*] Overriding built-in Normal style");
        string filePath = Path.Combine(folderPath, "OverrideNormalStyle.docx");
        var original = WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal) ?? throw new InvalidOperationException("Normal style definition was not found.");

        var custom = new WordParagraphStyleDefinition("Normal") {
            ColorHex = "0000FF",
            Bold = true
        };
        WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, custom);

        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("Paragraph with overridden Normal style");
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, original);
    }
}
