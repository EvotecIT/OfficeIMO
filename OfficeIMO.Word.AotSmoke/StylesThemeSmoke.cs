using OfficeIMO;
using OfficeIMO.Word;

internal static class StylesThemeSmoke {
    internal static void Run() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.Theme.SetColor(WordThemeColor.Accent1, OfficeColor.Parse("#123456"));
            WordStyle style = document.Styles.AddParagraphStyle("AotStyle", "AOT Style");
            style.ThemeColor = WordThemeColor.Accent1;
            style.LineSpacingAfter = OfficeLength.Points(6);
            style.Borders.BottomStyle = WordBorderStyle.Single;
            document.Styles.AddTableStyle("AotGrid").GetTableRegion(WordTableStyleRegion.FirstRow).Bold = true;
            document.AddParagraph("Styled").SetStyle(style);
            document.Save();
        }
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        WordStyle loaded = reopened.Styles["AotStyle"];
        if (loaded.Color != OfficeColor.Parse("#123456") || loaded.LineSpacingAfter != OfficeLength.Points(6) ||
            loaded.Borders.BottomStyle != WordBorderStyle.Single ||
            reopened.Styles["AotGrid"].GetTableRegion(WordTableStyleRegion.FirstRow).Bold != true) {
            throw new InvalidOperationException("Document styles or theme colours did not survive the round trip.");
        }
        Console.WriteLine("PASS | Document styles, table style regions and theme colours after reload");
    }
}
