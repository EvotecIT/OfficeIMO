# Adding Headers and Footers to Sections

Using OfficeIMO to add headers and footers to a Word document is a simple exercise.
Library fully supports adding headers and footers to different sections.
It supports different headers/footers for first page, even page, and odd page.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Sections[0].PageSetup.Orientation = PageOrientationValues.Landscape;
    document.AddParagraph("Test Section0");
    document.PageSetup.DifferentFirstPage = true;
    document.PageSetup.DifferentOddAndEvenPages = true;

    document.Sections[0].PageSetup.FirstPageHeader.AddParagraph().SetText("Test Section 0 - First Header");
    document.Sections[0].PageSetup.Header.AddParagraph().SetText("Test Section 0 - Header");
    document.Sections[0].PageSetup.EvenPageHeader.AddParagraph().SetText("Test Section 0 - Even");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section1 = document.AddSection();
    section1.PageSetup.Orientation = PageOrientationValues.Portrait;
    section1.AddParagraph("Test Section1");
    section1.PageSetup.Header.AddParagraph().SetText("Test Section 1 - Header");
    section1.PageSetup.DifferentFirstPage = true;
    section1.PageSetup.FirstPageHeader.AddParagraph().SetText("Test Section 1 - First Header");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section2 = document.AddSection();
    section2.AddParagraph("Test Section2");
    section2.PageSetup.Orientation = PageOrientationValues.Landscape;
    section2.PageSetup.Header.AddParagraph().SetText("Test Section 2 - Header");

    document.AddParagraph("Test Section2 - Paragraph 1");

    var section3 = document.AddSection();
    section3.AddParagraph("Test Section3");
    section3.PageSetup.Header.AddParagraph().SetText("Test Section 3 - Header");

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].PageSetup.Header.Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].PageSetup.Header.Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].PageSetup.Header.Paragraphs[0].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].PageSetup.Header.Paragraphs[0].Text);
    document.Save(true);
```