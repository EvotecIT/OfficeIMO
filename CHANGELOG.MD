## 0.13.0 - 2024.02.02
### What's Changed
* Add basic method to add textbox to header/footer by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/187
* [WordTextBox] Add additional features by @TopperDEL in https://github.com/EvotecIT/OfficeIMO/pull/188
    * Use Paragraph instead of StdBlock (so that text within a TextBox is no "variable" anymore)
    * Allow changing the insets of the TextBox so that there is no gap between the text of the TextBox and its border
    * Allow the TextBox to be set to "AutoFitText" or not
    * Allow the text to be multiline by respecting the page breaks
    * Added Wrapping functionality including inline
    * Added ability to convert inline to other wrapping types and vice versa (keep in mind that some properties are lost then or go to their defaults)
    * Fixes wrong paragraph being returned when AddImage is used
    * **Breaking Change** Changes WrapTextImage.InFrontText to WrapTextImage.InFrontOfText to match Word naming
* Improve WordTextBox functionality by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/197
* Add support for Net8 by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/198
* Fixes some warnings as reported by Visual Studio by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/199
* Added the ability to clear the default empty paragraph in TableCell. by @tmheath in https://github.com/EvotecIT/OfficeIMO/pull/182

### New Contributors
* @TopperDEL made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/188

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.12.0...v0.13.0

## 0.12.0 - 2023.12.31

### What's Changed
* Add descriptions to AddImage parameters to documentation by @tmheath in https://github.com/EvotecIT/OfficeIMO/pull/175
* Add ability to add TextBox to Word Document (new class WordTextBox) by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/180
* **BREAKING CHANGE** Improve Watermark with colors, rotation and other settings by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/181
* Word table cell paragraph add image fix by @tmheath in https://github.com/EvotecIT/OfficeIMO/pull/176
* Add support for WriteProtection (Always Read Only) and MarkAsFinal settings by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/163

### New Contributors
* @tmheath made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/175

More details. This change adds:
- Ability to modify watermark (colors, text, rotation, width, height)
- Ability to remove watermark
- Ability to find watermarks in document, sections, headers
- Ability to add watermark to document/section which makes watermark show up only on single page
- Ability to add watermark to headers/footers which makes watermark show up on whole section

**Breaking changes**
- This change breaks how watermarks are added. If you add them directly within section/document it will only apply to single page/pages as the SdtBlock gets added to body directly. If you need watermark for the whole section/document you need to add watermark to header/footer for it to apply to given section.

This example shows per section in header:

```cs
public static void Watermark_Sample1(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with Watermark 2");
    string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark 4.docx");

    using (WordDocument document = WordDocument.Create(filePath)) {
        document.AddParagraph("Section 0");
        document.AddHeadersAndFooters();
        document.Sections[0].Header.Default.AddParagraph("Section 0 - In header");
        document.Sections[0].SetMargins(WordMargin.Normal);

        Console.WriteLine(document.Sections[0].Margins.Left.Value);
        Console.WriteLine(document.Sections[0].Margins.Right.Value);

        Console.WriteLine(document.Sections[0].Margins.Type);

        document.Sections[0].Margins.Type = WordMargin.Wide;

        Console.WriteLine(document.Sections[0].Margins.Type);

        Console.WriteLine("----");
        var watermark = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");
        watermark.Color = Color.Red;

        // ColorHex normally returns hex colors, but for watermark it returns string as the underlying value is in string name, not hex
        Console.WriteLine(watermark.ColorHex);

        Console.WriteLine(watermark.Rotation);

        watermark.Rotation = 180;

        Console.WriteLine(watermark.Rotation);

        watermark.Stroked = true;

        Console.WriteLine(watermark.Height);
        Console.WriteLine(watermark.Width);

        // width and height in points (HTML wise)
        watermark.Height = 100.15;
        watermark.Width = 500.18;

        document.AddPageBreak();
        document.AddPageBreak();

        document.AddSection();

        document.AddParagraph("Section 1");

        document.Sections[1].AddHeadersAndFooters();
        document.Sections[1].Header.Default.AddParagraph("Section 1 - In header");
        document.Sections[1].Margins.Type = WordMargin.Narrow;
        Console.WriteLine("----");

        Console.WriteLine("Section 0 - Paragraphs Count: " + document.Sections[0].Header.Default.Paragraphs.Count);
        Console.WriteLine("Section 1 - Paragraphs Count: " + document.Sections[1].Header.Default.Paragraphs.Count);

        Console.WriteLine("----");
        document.Sections[1].AddParagraph("Test");
        document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Draft");

        Console.WriteLine(document.Sections[0].Margins.Left.Value);
        Console.WriteLine(document.Sections[0].Margins.Right.Value);

        Console.WriteLine(document.Sections[1].Margins.Left.Value);
        Console.WriteLine(document.Sections[1].Margins.Right.Value);

        Console.WriteLine(document.Sections[1].Margins.Type);


        document.Settings.SetBackgroundColor(Color.Azure);

        Console.WriteLine("----");

        Console.WriteLine("Watermarks in default header: " + document.Header.Default.Watermarks.Count);

        Console.WriteLine("Watermarks in default footer: " + document.Footer.Default.Watermarks.Count);

        Console.WriteLine("Watermarks in section 0: " + document.Sections[0].Watermarks.Count);
        Console.WriteLine("Watermarks in section 0 (header): " + document.Sections[0].Header.Default.Watermarks.Count);
        Console.WriteLine("Paragraphs in section 0 (header): " + document.Sections[0].Header.Default.Paragraphs.Count);

        Console.WriteLine("Watermarks in section 1: " + document.Sections[1].Watermarks.Count);
        Console.WriteLine("Watermarks in section 1 (header): " + document.Sections[1].Header.Default.Watermarks.Count);
        Console.WriteLine("Paragraphs in section 1 (header): " + document.Sections[1].Header.Default.Paragraphs.Count);

        Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);

        document.Save(false);
    }

    using (WordDocument document = WordDocument.Load(filePath)) {
        //Console.WriteLine("----");
        //Console.WriteLine("Watermarks in default header: " + document.Header.Default.Watermarks.Count);

        //Console.WriteLine("Watermarks in default footer: " + document.Footer.Default.Watermarks.Count);

        //Console.WriteLine("Watermarks in section 0: " + document.Sections[0].Watermarks.Count);
        //Console.WriteLine("Watermarks in section 0 (header): " + document.Sections[0].Header.Default.Watermarks.Count);
        //Console.WriteLine("Paragraphs in section 0 (header): " + document.Sections[0].Header.Default.Paragraphs.Count);

        //Console.WriteLine("Watermarks in section 1: " + document.Sections[1].Watermarks.Count);

        //Console.WriteLine("Paragraphs in section 1 (header): " + document.Sections[1].Header.Default.Paragraphs.Count);

        //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);

        document.Save(openWord);
    }
}
```

This example shows per page:

```
public static void Watermark_Sample3(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with watermark");
    string filePath = System.IO.Path.Combine(folderPath, "Basic Document with watermark and sections.docx");

    using (WordDocument document = WordDocument.Create(filePath)) {

        document.AddParagraph("Section 0");
        document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Confidential");

        document.AddPageBreak();
        document.AddPageBreak();

        var section = document.AddSection();
        section.AddWatermark(WordWatermarkStyle.Text, "Second Mark");

        document.AddParagraph("Section 1");

        document.AddPageBreak();
        document.AddPageBreak();

        var section1 = document.AddSection();

        document.AddParagraph("Section 2");

        document.Sections[2].AddWatermark(WordWatermarkStyle.Text, "New");

        document.AddPageBreak();
        document.AddPageBreak();

        Console.WriteLine("----");
        Console.WriteLine("Watermarks: " + document.Watermarks.Count);
        Console.WriteLine("Watermarks section 0: " + document.Sections[0].Watermarks.Count);
        Console.WriteLine("Watermarks section 1: " + document.Sections[1].Watermarks.Count);
        Console.WriteLine("Watermarks section 2: " + document.Sections[2].Watermarks.Count);

        Console.WriteLine("Paragraphs: " + document.Paragraphs.Count);

        Console.WriteLine("Removing last watermark");

        document.Sections[2].Watermarks[0].Remove();

        Console.WriteLine("Watermarks: " + document.Watermarks.Count);
        Console.WriteLine("Watermarks section 0: " + document.Sections[0].Watermarks.Count);
        Console.WriteLine("Watermarks section 1: " + document.Sections[1].Watermarks.Count);
        Console.WriteLine("Watermarks section 2: " + document.Sections[2].Watermarks.Count);
        Console.WriteLine("Paragraphs: " + document.Paragraphs.Count);

        document.Save(openWord);
    }
}
```

- Adds WordBordersParagraph type and allows setting borders for paragraphs
- Adds ability to add TextBox

```cs
internal static void Example_AddingTextbox2(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with some textbox");

    var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox3.docx");

    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Adding paragraph with some text");

        var textBox = document.AddTextBox("My textbox on the left");

        textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
        // horizontal alignment overwrites the horizontal position offset so only one will work
        textBox.HorizontalAlignment = HorizontalAlignmentValues.Left;
        textBox.VerticalPositionOffsetCentimeters = 3;

        var textBox2 = document.AddTextBox("My textbox on the right");
        textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
        textBox2.WordParagraph.ParagraphAlignment = JustificationValues.Right;
        // horizontal alignment overwrites the horizontal position offset so only one will work
        textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;
        textBox2.VerticalPositionOffsetCentimeters = 3;

        Console.WriteLine(textBox.VerticalPositionOffsetCentimeters);

        Console.WriteLine(document.TextBoxes[0].VerticalPositionOffsetCentimeters);

        Console.WriteLine(document.TextBoxes[1].VerticalPositionOffsetCentimeters);

        document.Save(openWord);
    }
}
```

- Allow setting document to AlwaysOpenReadOnly

```cs
internal static void Example_ProtectAlwaysReadOnly(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with protection 'Always Read Only'");
    string filePath = System.IO.Path.Combine(folderPath, "Basic Document with always read only protection.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Basic paragraph - Page 4");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Color = SixLabors.ImageSharp.Color.Blue;
        paragraph.AddText(" This is continutation in the same line");

        Console.WriteLine("Always read only: " + document.Settings.AlwaysOpenReadOnly);

        document.Settings.AlwaysOpenReadOnly = true;

        Console.WriteLine("Always read only: " + document.Settings.AlwaysOpenReadOnly);

        document.Save(true);
    }
}
```

- Allow setting document to FinalDocument

```cs
internal static void Example_ProtectFinalDocument(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating basic document with protection - Final Document");
    string filePath = System.IO.Path.Combine(folderPath, "Basic Document with setting Word to Final Document.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Basic paragraph - Page 1");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Color = SixLabors.ImageSharp.Color.Blue;
        paragraph.AddText(" This is continutation in the same line");
        paragraph.AddBreak(BreakValues.TextWrapping);

        Console.WriteLine("Final document: " + document.Settings.FinalDocument);

        document.Settings.FinalDocument = true;

        Console.WriteLine("Final document: " + document.Settings.FinalDocument);

        document.Save(openWord);
    }
}
```

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.11.0...v0.12.0

## 0.11.0 - 2023.11.27

### What's Changed
* FindAndReplace allowed for specific paragraphs by @startewho in https://github.com/EvotecIT/OfficeIMO/pull/171

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.10.0...v0.11.0

## 0.10.0 - 2023.10.24

### What's Changed
* Add set styleId and set paragraph alignment method by @startewho in https://github.com/EvotecIT/OfficeIMO/pull/160
* Fixed pos image by @boatwrong in https://github.com/EvotecIT/OfficeIMO/pull/165
* Add Areachat by @startewho in https://github.com/EvotecIT/OfficeIMO/pull/166

### New Contributors
* @boatwrong made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/165

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.9.0...v0.10.0

## 0.9.0 - 2023.08.15

### What's Changed
* Fixes saving properties when using Stream by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/159

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.8.0...v0.9.0

## 0.8.0 - 2023.07.30

### What's Changed
* AddRow functions return the created rows and table cell VerticalAlignment property by @The-Faulty in https://github.com/EvotecIT/OfficeIMO/pull/155
* Improve Lists in few areas by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/110
* Add Hyperlink inside table cell by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/157

### New Contributors
* @The-Faulty made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/155

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.7.0...v0.8.0

## 0.7.0 - 2023.07.16

### What's Changed
* Implements muti paragraph search and replace by @startewho in https://github.com/EvotecIT/OfficeIMO/pull/149 heavily improving the FindAndReplace functionality
* Add support for Footnotes and endnotes by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/154

```cs
internal static void Example_DocumentWithFootNotesEmpty(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with footnotes/end notes");
    string filePath = System.IO.Path.Combine(folderPath, "Document with FootNotes02.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Basic paragraph");
        paragraph.ParagraphAlignment = JustificationValues.Center;

        document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text")
            .AddText(" continuing").AddFootNote("2nd footnote!");

        Console.WriteLine("EndNotes count " + document.EndNotes.Count);
        Console.WriteLine("EndNotes Section count " + document.Sections[0].EndNotes.Count);

        Console.WriteLine("FootNotes count " + document.FootNotes.Count);
        Console.WriteLine("FootNotes Section count " + document.Sections[0].FootNotes.Count);


        var lastFootNoteParagraph = document.AddParagraph("Another paragraph").AddFootNote("more footnotes!")
            .AddText(" more within paragraph").AddFootNote("4th footnote!");

        Console.WriteLine("Is paragraph foot note: " + lastFootNoteParagraph.IsFootNote);

        var footNoteParagraphs = lastFootNoteParagraph.FootNote.Paragraphs;

        Console.WriteLine("Paragraphs within footnote: " + footNoteParagraphs.Count);
        Console.WriteLine("What's the text: " + footNoteParagraphs[1].Text);
        footNoteParagraphs[1].Bold = true;

        document.AddParagraph("Testing endnote - 1").AddEndNote("Test end note 1");

        document.AddParagraph("Test 1");

        document.AddSection();

        document.AddParagraph("Testing endnote - 2").AddEndNote("Test end note 2");

        Console.WriteLine("EndNotes count " + document.EndNotes.Count);
        Console.WriteLine("EndNotes Section count " + document.Sections[0].EndNotes.Count);

        Console.WriteLine("FootNotes count " + document.FootNotes.Count);
        Console.WriteLine("FootNotes Section count " + document.Sections[0].FootNotes.Count);

        document.Save(openWord);
    }
}
```

### New Contributors
* @startewho made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/149

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.6.0...v0.7.0

## 0.6.0 - 2023.07.03

### What's Changed
- Add ability to add charts to WordParagraph by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/144
- Add missing tests for AddTableAfter() and AddTableBefore()  by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/131

This  release fixes
- https://github.com/EvotecIT/OfficeIMO/issues/143
- It allows to add `chart `to existing paragraph (before it was only possible to assign chart to document).
- `WordParagraph `now contains `IsChart` and `Chart` object
- `WordDocument `and `WordSection `contains `ParagraphsCharts `and `Charts `lists

```cs
public static void Example_AddingMultipleCharts(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with charts");
    string filePath = System.IO.Path.Combine(folderPath, "Charts Document.docx");

    using (WordDocument document = WordDocument.Create(filePath)) {
        List<string> categories = new List<string>() {
            "Food", "Housing", "Mix", "Data"
        };

        var paragraphToTest = document.AddParagraph("Test showing adding chart right to existing paragraph");

        // adding charts to document
        document.AddParagraph("This is a bar chart");
        var barChart1 = document.AddBarChart();
        barChart1.AddCategories(categories);
        barChart1.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
        barChart1.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
        barChart1.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        barChart1.BarGrouping = BarGroupingValues.Clustered;
        barChart1.BarDirection = BarDirectionValues.Column;

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        document.AddParagraph("This is a bar chart");
        var barChart2 = document.AddBarChart();
        barChart2.AddCategories(categories);
        barChart2.AddChartBar("USA", 15, Color.Aqua);
        barChart2.RoundedCorners = true;


        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        document.AddParagraph("This is a pie chart");
        var pieChart = document.AddPieChart();
        pieChart.AddCategories(categories);
        pieChart.AddChartPie("Poland", new List<int> { 15, 20, 30 });

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);


        document.AddParagraph("Adding a line chart as required 1");

        var lineChart = document.AddLineChart();
        lineChart.AddChartAxisX(categories);
        lineChart.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        lineChart.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
        lineChart.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        document.AddParagraph("Adding a line chart as required 2");

        var lineChart2 = document.AddLineChart();
        lineChart2.AddChartAxisX(categories);
        lineChart2.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        lineChart2.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
        lineChart2.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        // adding charts to paragraphs directly
        var paragraph = document.AddParagraph("This is a bar chart - but assigned to paragraph 1");
        var barChart3 = paragraph.AddBarChart();
        barChart3.AddCategories(categories);
        barChart3.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
        barChart3.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
        barChart3.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        barChart3.BarGrouping = BarGroupingValues.Clustered;
        barChart3.BarDirection = BarDirectionValues.Column;

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        var paragraph1 = document.AddParagraph("This is a bar chart - but assigned to paragraph 2");
        var barChart5 = paragraph1.AddBarChart();
        barChart5.AddCategories(categories);
        barChart5.AddChartBar("USA", 15, Color.Aqua);
        barChart5.RoundedCorners = true;

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        var paragraph2 = document.AddParagraph("This is a pie chart - but assigned to paragraph");
        var pieChart1 = paragraph2.AddPieChart();
        pieChart1.AddCategories(categories);
        pieChart1.AddChartPie("Poland", new List<int> { 15, 20, 30 });

        var paragraph3 = document.AddParagraph("Adding a line chart as required 1 - but assigned to paragraph");
        var lineChart3 = paragraph3.AddLineChart();
        lineChart3.AddChartAxisX(categories);
        lineChart3.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        lineChart3.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
        lineChart3.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        var paragraph4 = document.AddParagraph("Adding a line chart as required 2 - but assigned to paragraph");
        var lineChart4 = paragraph4.AddLineChart();
        lineChart4.AddChartAxisX(categories);
        lineChart4.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        lineChart4.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
        lineChart4.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        // lets add chart to first paragraph
        var lineChart5 = paragraphToTest.AddLineChart();
        lineChart5.AddChartAxisX(categories);
        lineChart5.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        lineChart5.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
        lineChart5.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        var table = document.AddTable(3, 3);
        table.Rows[0].Cells[0].Paragraphs[0].AddBarChart();
        barChart3.AddCategories(categories);
        barChart3.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
        barChart3.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
        barChart3.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
        barChart3.BarGrouping = BarGroupingValues.Clustered;
        barChart3.BarDirection = BarDirectionValues.Column;

        Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

        Console.WriteLine("Images count: " + document.Sections[0].Images.Count);

        document.Save(openWord);
    }
}
```


**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.5.0...v0.6.0

## 0.5.0 - 2023.06.05
### What's Changed
* Add example to show margins and font family by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/140
* Add additional example by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/141
* Downgrade DocumentFormat.OpenXml by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/135
* Add static method for AddHeadersAndFooters() by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/134
* Add ability to use NewLines ("\r\n"/Environment.NewLine) in AddParagraph, AddText by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/127
* Fix for Incorrectly setting the LineSpacingRule property for paragraph by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/147
* Removed .NET Core 3.1 from Tests

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.9...v0.5.0

## 0.4.9 - 2023.03.19
### What's Changed
* Add very basic find and replace functionality by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/108
* Fixes Color to return Null instead throw an error by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/114
* Add AddParagraph() and AddParagraph(string text) to TableCell by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/115
* Add additional example for Fields by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/122
* Add Tab Stops to Word Paragraphs by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/117
* Add support for Tabs by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/128
* Added additional custom properties handling by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/132

### Possibly breaking change
- Set HighAnsi, EastAsia and ComplexScript at the same time as FontFamily by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/125

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.8...v0.4.9

## 0.4.8 - 2023.02.08

### What's Changed
* Basic support for embedding RTF & HTML & More (AddEmbeddedDocument) by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/104
* Adds CleanupDocument method to merge same formatted runs by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/106

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.7...v0.4.8

## 0.4.7 - 2023.01.14

### What's Changed
* Allows applying custom styling to Hyperlinks by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/100
* Remove exception from wordsection by @crjc in https://github.com/EvotecIT/OfficeIMO/pull/99

### New Contributors
* @crjc made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/99

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.6...v0.4.7

## 0.4.6 - 2023.01.11

### What's Changed
* Added tests for fragmented instruction parsing by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/97
* Add paragraphs property to start on a new page by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/98

## 0.4.5 - 2023.01.09

### What's Changed
* Merge instructions distributed over runs to field string by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/95

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.4...v0.4.5

## 0.4.4 - 2023.01.09

### What's Changed
* Fixes HyperLink AddStyle not being applied by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/93

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.3...v0.4.4

## 0.4.3 - 2023.01.09

### What's Changed
* Improve tests, fix HyperLink order, add method ValidateDocument() by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/87
* Fix badges by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/88
* Ability to insert table before/after paragraph by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/92
  * Paragraph.AddTableAfter()
  * Paragraph.AddTableBefore()

```csharp
internal static void Example_TablesAddedAfterParagraph(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with width and alignment");
    string filePath = System.IO.Path.Combine(folderPath, "Document with Table Alignment.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Lets add table with some alignment ");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Bold = true;
        paragraph.Underline = UnderlineValues.DotDash;

        WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        var paragraph1 = document.AddParagraph("Lets add another table showing text wrapping around, but notice table before and after it anyways, that we just added at the end of the document.");

        WordTable wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        wordTable1.WidthType = TableWidthUnitValues.Pct;
        wordTable1.Width = 3000;

        wordTable1.AllowTextWrap = true;

        var paragraph2 = document.AddParagraph("This paragraph should continue but next to to the table");

        document.AddParagraph();
        document.AddParagraph();

        var paragraph3 = document.AddParagraph("Lets add another table showing AutoFit");

        WordTable wordTable2 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";


        paragraph1.AddParagraphBeforeSelf();
        paragraph1.AddParagraphAfterSelf();

        var table3 = paragraph1.AddTableAfter(4, 4, WordTableStyle.GridTable1LightAccent1);
        table3.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document after paragraph";

        var table4 = paragraph1.AddTableBefore(4, 4, WordTableStyle.GridTable1LightAccent1);
        table4.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document before paragraph";

        document.Save(openWord);
    }
}
```

## 0.4.2 - 2022.11.25

### What's Changed
* The order of paragraph properties and runs does matter by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/83
* This release fixes LISTS. Sorry for that!

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.1...v0.4.2

## 0.4.1 - 2022.11.20

### What's Changed
* Extending available FieldCodes and allow setting parameters by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/75
* Fix check IsToc for Word.Lists by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/76
* Provide multi platform project config by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/79

### New Contributors
* @byteSamurai made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/75

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.0...v0.4.1


## 0.4.0 - 2022.11.13

## What's Changed
* Improve settings for document with PT-BR and HighAnsi chars by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/56
  - FontFamilyHighAnsi for document.Settings which is required for special chars

```csharp
public static void Example_BasicWordWithDefaultFontChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style (PT/BR)");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChangeBR.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        //document.Settings.FontSizeComplexScript = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.FontFamilyHighAnsi = "Calibri Light";
        document.Settings.Language = "pt-Br";

        string title = "INSTRUMENTO PARTICULAR DE CONSTITUIÇÃO DE GARANTIA DE ALIENAÇÃO FIDUCIÁRIA DE IMÓVEL";

        document.AddParagraph(title).SetBold().ParagraphAlignment = JustificationValues.Center;

        document.Save(openWord);
    }
}
```

* Add Compatibility Mode, set default compatibility to the highest version by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/58
  - Adds default settings, including compatibility settings, math properties, and few other things that are automatically added when Microsoft Word creates a document
  - Adds default web settings that are the same to what Microsoft Word does
  - Sample of usage for compatiblity mode

```csharp
public static void Example_BasicWordWithDefaultFontChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style (PT/BR)");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChangeBR.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        //document.Settings.FontSizeComplexScript = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.FontFamilyHighAnsi = "Calibri Light";
        document.Settings.Language = "pt-Br";

        document.Settings.ZoomPreset = PresetZoomValues.BestFit;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2013;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.None;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        string title = "INSTRUMENTO PARTICULAR DE CONSTITUIÇÃO DE GARANTIA DE ALIENAÇÃO FIDUCIÁRIA DE IMÓVEL";

        document.AddParagraph(title).SetBold().ParagraphAlignment = JustificationValues.Center;

        document.Save(openWord);
    }
}
```

* Bump nuggets for tests and excel by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/59
* Fix not saving list when saving to Stream by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/62
* Add an image via stream by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/51
* Add images to headers/footers/tables and other image improvements by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/53

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.BuiltinDocumentProperties.Title = "This is sparta";
    document.BuiltinDocumentProperties.Creator = "Przemek";
    var filePathImage = System.IO.Path.Combine(imagePaths, "Kulek.jpg");

    document.AddHeadersAndFooters();

    var header = document.Header.Default;
    var paragraphHeader = header.AddParagraph("This is header");

    // add image to header, directly to paragraph
    header.AddParagraph().AddImage(filePathImage, 100, 100);

    // add image to footer, directly to paragraph
    document.Footer.Default.AddParagraph().AddImage(filePathImage, 100, 100);

    // add image to header, but to a table
    var table = header.AddTable(2, 2);
    table.Rows[1].Cells[1].Paragraphs[0].Text = "Test123";
    table.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);
    table.Alignment = TableRowAlignmentValues.Right;

    var paragraph = document.AddParagraph("This paragraph starts with some text");
    paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";
    paragraph.Bold = true;

    // add table with an image, but to document
    var table1 = document.AddTable(2, 2);
    table1.Rows[1].Cells[1].Paragraphs[0].Text = "Test - In document";
    table1.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);


    // lets add image to paragraph
    paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
}
```

- Rotation / VerticalFlip / HorizontalFlip / Shape work for images

```csharp
internal static void Example_AddingImagesInline(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with inline images");
    string filePath = System.IO.Path.Combine(folderPath, "DocumentWithInlineImages2.docx");
    string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

    using (WordDocument document = WordDocument.Create(filePath)) {
        var file = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
        var paragraph = document.AddParagraph();
        var pargraphWithImage = paragraph.AddImage(file, 100, 100);

        // Console.WriteLine("Image is inline: " + pargraphWithImage.Image.Rotation);

        pargraphWithImage.Image.VerticalFlip = false;
        pargraphWithImage.Image.HorizontalFlip = false;
        pargraphWithImage.Image.Rotation = 270;
        pargraphWithImage.Image.Shape = ShapeTypeValues.Cloud;


        document.Save(openWord);
    }
}
```

- Added optional Description (AltText) to images
- Added ability to wrap text around image

```cs
internal static void Example_AddingImagesSample4(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with some Images and Samples");
    var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImagesSample4.docx");
    var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

    using var document = WordDocument.Create(filePath);

    var paragraph1 = document.AddParagraph("This paragraph starts with some text");
    paragraph1.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);
    paragraph1.Image.Shape = ShapeTypeValues.Cube;

    var paragraph2 = document.AddParagraph("Image will be placed behind text");
    paragraph2.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.BehindText, "Przemek and Kulek on an image");


    var paragraph3 = document.AddParagraph("Image will be in front of text");
    paragraph3.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.InFrontOfText, "Przemek and Kulek on an image");


    var paragraph5 = document.AddParagraph("Image will be Square");
    paragraph5.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Square, "Przemek and Kulek on an image");


    var paragraph6 = document.AddParagraph("Image will be Through");
    paragraph6.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Through, "Przemek and Kulek on an image");


    var paragraph7 = document.AddParagraph("Image will be Tight");
    paragraph7.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Tight, "Przemek and Kulek on an image");


    var paragraph8 = document.AddParagraph("Image will be Top And Bottom");
    paragraph8.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.TopAndBottom, "Przemek and Kulek on an image");
    paragraph8.Image.Shape = ShapeTypeValues.Can;

    document.Save(openWord);
}
```

* Fixes adding a bookmark on a document by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/68
* Detect read-only before saving by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/67
* Skip creation of run, which isn't needed by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/69
* Add net 7.0, bump dependencies for testing by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/72

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.3.1...v0.4.0

## 0.3.1 - 2022.10.18
- Adds BMP, GIF, PNG, TIFF image formats support [#42](https://github.com/EvotecIT/OfficeIMO/pull/42) by rstm-sf
- Fixes issue with read only documents crashing
- Fixes watermark throwing when Headers aren't added first [#46](https://github.com/EvotecIT/OfficeIMO/pull/46)
- Fixes Adding Watermark to section that has not initialized headers will throw an error [#27](https://github.com/EvotecIT/OfficeIMO/issues/27)
- New fields added to WordSettings allowing setting of default values for the whole document for FontSize, FontSizeComplexScript, FontFamily and Language

```csharp
public static void Example_BasicWordWithDefaultStyleChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChange.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.Language = "pl-PL";

        var paragraph1 = document.AddParagraph("To jest po polsku");

        var paragraph2 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");
        paragraph2.FontSize = 15;
        paragraph2.FontFamily = "Courier New";

        document.Save(openWord);
    }
}

```
- Add `IsLastRun` and `IsFirstRun` properties to WordParagraph

## 0.3.0 - 2022.10.11
- Update DocumentFormat.OpenXml from 2.16.0 to 2.18.0
- Update SixLabors.ImageSharp to 2.1.3

- Adds ability to add nested table into existing table.

```csharp
internal static void Example_NestedTables(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with nested tables");
    string filePath = System.IO.Path.Combine(folderPath, "Document with Nested Tables.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Lets add table ");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Bold = true;
        paragraph.Underline = UnderlineValues.DotDash;

        WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        wordTable.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

        wordTable.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

        document.Save(openWord);
    }
}
```

- Adds `NestedTables` property for WordTable to get all nested tables for given table
- Adds `HasNestedTables` property for WordTable to know if table has nested tables
- Adds `IsNestedTable` property for WordTable to know if table is nested table
- Adds `ParentTable` property for WordTable to find parent table if the table is nested
- Added some summaries to multiple table related methods/properties
- Adds `TablesIncludingNestedTables` property to Sections and Document to make it easy to find all tables within document and manipulate them

- Solves an issue with different word break required https://github.com/EvotecIT/OfficeIMO/issues/37

![image](https://user-images.githubusercontent.com/15063294/194375619-a43d0d20-50b9-445a-a8c7-d778a4c3fac2.png)

```csharp
public static void Example_BasicWordWithBreaks(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with paragraph & breaks");
    string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphsAndBreaks.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph1 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");


        var paragraph2 = document.AddParagraph("Adding paragraph2 with some text and pressing SHIFT+ENTER");
        paragraph2.AddBreak();
        paragraph2.AddText("Continue1");
        paragraph2.AddBreak();
        paragraph2.AddText("Continue2");

        var paragraph3 = document.AddParagraph("Adding paragraph3 with some text and pressing ENTER");


        document.Save(openWord);
    }
}
```

Additionally:
- Renames WordPageBreak to WordBreak to accommodate all Breaks, and not only PageBreak
- *BREAKING CHANGE* Removing WordBreak (or WordPageBreak) no longer by default removes paragraph, but instead requires bool set to true

```csharp
document.Breaks[0].Remove();
document.Breaks[0].Remove(includingParagraph: true);
```

- Add new `IsBreak` property for WordParagraph
- Add `Breaks` property for WordDocument

- Implement Save and Load to/from Stream [#43](https://github.com/EvotecIT/OfficeIMO/pull/43) tnx hisuwh

## - 0.2.1 - 2022.07.31
- Added basic support for Charts [#14](https://github.com/EvotecIT/OfficeIMO/pull/14)

## - 0.2.0 - 2022.07.31

- Added MIT License [#28](https://github.com/EvotecIT/OfficeIMO/pull/28)

- Adds new properties for Tables by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ Alignment
  - ☑️ WidthType
  - ☑️ Width
  - ☑️ ShadingFillColor
  - ☑️ ShadingFillColorHex
  - ☑️ ShadingFillPatern
  - ☑️ Title
  - ☑️ Description
  - ☑️ AllowOverlap
  - ☑️ AllowTextWrap
  - ☑️ ColumnWidth
  - ☑️ RowHeight

- Add table positioning along with properties by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ LeftFromText
  - ☑️ RightFromText
  - ☑️ BottomFromText
  - ☑️ TopFromText
  - ☑️ HorizontalAnchor
  - ☑️ TablePositionY
  - ☑️ TablePositionX
  - ☑️ TablePositionYAlignment
  - ☑️ TablePositionXAlignment
  - ☑️ TableOverlap

- Adds new properties for TableRow by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ FirstCell
  - ☑️ LastCell

- Renames some properties to better name them by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - FirstRow -> ConditionalFormattingFirstRow
  - LastRow -> ConditionalFormattingLastRow
  - FirstColumn -> ConditionalFormattingFirstColumn
  - LastColumn -> ConditionalFormattingLastColumn
  - NoHorizontalBand -> ConditionalFormattingNoHorizontalBand
  - NoVerticalBand -> ConditionalFormattingNoVerticalBand

- Adds new properties for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ FirstRow
  - ☑️ LastRow

- Adds new methods for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ AddComment(author, initials,comment)

- Adds new properties for TableCell by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ TextDirection

## - 0.1.7 - 2022.06.12

### What's Changed
* Fixes **PageOrientation** of page/section if set before page size is applied. In case that happens it always reverted back to `Portrait` mode which is default for newly set **PageSizes**.
* Fixes **PageSize** detection when in `PageOrientationValues.Landscape` mode.

## - 0.1.6 - 2022.06.11

### What's Changed
* Rename `Color` to `ColorHex` property for Paragraphs *BREAKING CHANGE*
* Add `Color` property for Paragraphs as `SixLabors.ImageSharp.Color` *BREAKING CHANGE*

For example:

```csharp
var paragraph = document.AddParagraph("Basic paragraph");
paragraph.ParagraphAlignment = JustificationValues.Center;
paragraph.Color = SixLabors.ImageSharp.Color.Red;
```

```csharp
var paragraph = document.AddParagraph("Basic paragraph");
paragraph.ParagraphAlignment = JustificationValues.Center;
paragraph.ColorHex = "#FFFF00";
```

## - 0.1.5 - 2022.06.04

### What's Changed
* Fixes `TableOfContent.Update()`
* Fixed **SaveAs** functionality by @jordan-hemming in https://github.com/EvotecIT/OfficeIMO/pull/16
* Fixes **Azure Devops** tests to work properly on Linux and MacOs by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/18
* Tables styles are not being applied for loaded documents by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/20
* Add basic support for table cell borders by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/21
  - ☑️ LeftBorder
  - ☑️ RightBorder
  - ☑️ TopBorder
  - ☑️ BottomBorder
  - ☑️ EndBorder
  - ☑️ StartBorder
  - ☑️ InsideHorizontalBorder
  - ☑️ InsideVerticalBorder
  - ☑️ TopLeftToBottomRightBorder
  - ☑️ TopRightToBottomLeftCell
* Add additional tests for table cell borders by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/22

### New Contributors
* @jordan-hemming made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/16

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.1.4...v0.1.5

## - 0.1.4 - 2022.04.03
- First official release
