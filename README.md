# OfficeIMO - Microsoft Word .NET Library

OfficeIMO is available as NuGet from the gallery and its preferred way of using it.

[![nuget downloads](https://img.shields.io/nuget/dt/officeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](#)
[![nuget downloads](https://wakatime.com/badge/user/f1abc372-39bb-4b06-ad2b-3a24cf161f13/project/3cddaa3c-574a-400b-9870-d0973797eb51.svg)](#)

If you would like to contact me you can do so via Twitter or LinkedIn.

[![twitter](https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social)](https://twitter.com/PrzemyslawKlys)
[![blog](https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg)](https://evotec.xyz/hub)
[![linked](https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn)](https://www.linkedin.com/in/pklys)

## What it's all about

This is a small project (under development) that allows to create Microsoft Word documents (.docx) using .NET.
Underneath it uses [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) but heavily simplifies it.
It was created because working with OpenXML is way too hard for me, and time consuming.
I created it for use within the PowerShell module called [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice),
but thought it may be useful for others to use in the .NET community.

If you want to see what it can do take a look at this [blog post](https://evotec.xyz/officeimo-free-cross-platform-microsoft-word-net-library/).

I used to use the DocX library (which I co-authored/maintained before it was taken over by Xceed) to create Microsoft Word documents,
but it only supports .NET Framework, and their newest community license makes the project unusable.

*As I am not really a developer, and I hardly know what I'm doing if you know how to help out - please do.*

- If you see bad practice, please open an issue/submit PR.
- If you know how to do something in OpenXML that could help this project - please open an issue/submit PR
- If you see something that could work better - please open an issue/submit PR
- If you see something that I made a fool of myself - please open an issue/submit PR
- If you see something that works not the way I think it works - please open an issue/submit PR

I hope you get the drift? If it's bad - open an issue/fix it! I don't know what I'm doing!
The main thing is - it has to work with .NET Framework 4.7.2, .NET Standard 2.0 and so on.

**This project is under development and as such there's a lot of things that can and will change, especially if some people help out.**

| Platform | Status                                                                                                                                                                                                 | Code Coverage                                                                                                                                                                                                                                           | .NET                                                                          |
| -------- |--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| Windows  | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=19"><img src="https://img.shields.io/azure-devops/tests/evotecpl/OfficeIMO/19/master?compact_message&label=Tests%20Windows"></a> | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=19&view=ms.vss-pipelineanalytics-web.new-build-definition-pipeline-analytics-view-cardmetrics"><img src="https://img.shields.io/azure-devops/coverage/evotecpl/OfficeIMO/19"></a> | .NET 4.7.2, NET 4.8, .NET 6.0, .NET 7.0, .NET 8.0, .NET Standard 2.0, .NET Standard 2.1 |
| Linux    | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=22"><img src="https://img.shields.io/azure-devops/tests/evotecpl/OfficeIMO/22/master?compact_message&label=Tests%20Linux"></a>   | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=22&view=ms.vss-pipelineanalytics-web.new-build-definition-pipeline-analytics-view-cardmetrics"><img src="https://img.shields.io/azure-devops/coverage/evotecpl/OfficeIMO/22"></a> | .NET 6.0, .NET 7.0, .NET Standard 2.0, .NET 8.0, .NET Standard 2.1, .NET Core 3.1       |
| MacOs    | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=23"><img src="https://img.shields.io/azure-devops/tests/evotecpl/OfficeIMO/23/master?compact_message&label=Tests%20MacOs"></a>   | <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build?definitionId=23&view=ms.vss-pipelineanalytics-web.new-build-definition-pipeline-analytics-view-cardmetrics"><img src="https://img.shields.io/azure-devops/coverage/evotecpl/OfficeIMO/23"></a> | .NET 6.0, .NET 7.0, .NET Standard 2.0, .NET 8.0, .NET Standard 2.1, .NET Core 3.1       |

## Support This Project

If you find this project helpful, please consider supporting its development.
Your sponsorship will help the maintainers dedicate more time to maintenance and new feature development for everyone.

It takes a lot of time and effort to create and maintain this project.
By becoming a sponsor, you can help ensure that it stays free and accessible to everyone who needs it.

To become a sponsor, you can choose from the following options:

 - [Become a sponsor via GitHub Sponsors :heart:](https://github.com/sponsors/PrzemyslawKlys)
 - [Become a sponsor via PayPal :heart:](https://paypal.me/PrzemyslawKlys)

Your sponsorship is completely optional and not required for using this project.
We want this project to remain open-source and available for anyone to use for free,
regardless of whether they choose to sponsor it or not.

If you work for a company that uses our .NET libraries or PowerShell Modules,
please consider asking your manager or marketing team if your company would be interested in supporting this project.
Your company's support can help us continue to maintain and improve this project for the benefit of everyone.

Thank you for considering supporting this project!

## Please share with the community

Please consider sharing a post about OfficeIMO and the value it provides. It really does help!

[![Share on reddit](https://img.shields.io/badge/share%20on-reddit-red?logo=reddit)](https://reddit.com/submit?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)
[![Share on hacker news](https://img.shields.io/badge/share%20on-hacker%20news-orange?logo=ycombinator)](https://news.ycombinator.com/submitlink?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on twitter](https://img.shields.io/badge/share%20on-twitter-03A9F4?logo=twitter)](https://twitter.com/share?url=https://github.com/EvotecIT/OfficeIMO&t=OfficeIMO)
[![Share on facebook](https://img.shields.io/badge/share%20on-facebook-1976D2?logo=facebook)](https://www.facebook.com/sharer/sharer.php?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on linkedin](https://img.shields.io/badge/share%20on-linkedin-3949AB?logo=linkedin)](https://www.linkedin.com/shareArticle?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)

## Features

Here's a list of features currently supported (and probably a lot I forgot) and those that are planned. It's not a closed list, more of TODO, and I'm sure there's more:

- ☑️ Word basics
  - ☑️ Create
  - ☑️ Load
  - ☑️ Save (auto open on save as an option)
  - ☑️ SaveAs (auto open on save as an option)
- ☑️ Word properties
  - ☑️ Reading basic and custom properties
  - ☑️ Setting basic and custom properties
- ☑️ Sections
  - ☑️ Add Paragraphs
  - ☑️ Add Headers and Footers (Odd/Even/First)
  - ◼️ Remove Headers and Footers (Odd/Even/First)
  - ☑️ Remove Paragraphs
  - ◼️ Remove Sections
- ☑️ Headers and Footers in the document (not including sections)
  - ☑️ Add Default, Odd, Even, First
  - ◼️ Remove Default, Odd, Even, First
- ☑️ Paragraphs/Text and make it bold, underlined, colored and so on
- ☑️ Paragraphs and change alignment
- ☑️ Tables
  - ☑️ [Add and modify table styles (one of 105 built-in styles)](https://evotec.xyz/docs/adding-tables-with-built-in-styles-managing-borders/)
  - ☑️ Add rows and columns
  - ☑️ Add cells
  - ☑️ Add cell properties
  - ☑️ [Add and modify table cell borders](https://evotec.xyz/docs/adding-tables-with-built-in-styles-managing-borders/)
  - ☑️ Remove rows
  - ☑️ Remove cells
  - ☑️ Others
    - ☑️ Merge cells (vertically, horizontally)
    - ◼️ Split cells (vertically)
    - ☑️ Split cells (horizontally)
- ☑️ Images/Pictures (limited support - jpg only / inline type only)
  - ☑️ Add images from file to Word
  - ☑️ Save image from Word to File
  - ◼️ Other image types
  - ◼️ Other location types
- ☑️ Hyperlinks
  - ☑️ Add HyperLink
  - ☑️ Read HyperLink
  - ◼️ Remove HyperLink
  - ☑️ Change HyperLink
- ☑️ PageBreaks
  - ☑️ Add PageBreak
  - ☑️ Read PageBreak
  - ☑️ Remove PageBreak
  - ☑️ Change PageBreak
- ☑️ Bookmarks
  - ☑️ Add Bookmark
  - ☑️ Read Bookmark
  - ☑️ Remove Bookmark
  - ☑️ Change Bookmark
- ◼️ Comments
  - ☑️ Add comments
  - ☑️ Read comments
  - ◼️ Remove comments
  - ◼️ Track comments
- ☑️ Fields
  - ☑️ Add Field
  - ☑️ Read Field
  - ☑️ Remove Field
  - ☑️ Change Field
- ☑️ Footnotes
  - ☑️ Add new footnotes
  - ☑️ Read footnotes
  - ☑️ Remove footnotes
- ☑️ Endnotes
  - ☑️Add new endnotes
  - ☑️Read endnotes
  - ☑️Remove endnotes
- ◼️ Shapes
- ◼️ Charts
  - ☑️ Add charts
- ◼️ Lists
  - ☑️ Add lists
  - ◼️ Remove lists
- ◼️ Table of contents
  - ☑️ Add TOC
- ☑️ Borders
- ☑️ Background
- ◼️ Watermarks
  - ☑️ Add watermark
  - ◼️ Remove watermark
- ☑️ PageBreaks
  - ☑️Add pagebreak
  - ☑️Find pagebreak
  - ☑️Remove pagebreak


## Features (oneliners):

This list of features is for times when you want to quickly fix something rather than playing with full features.
This features are available as part of `WordHelpers` class.

- ☑️ Remove Headers and Footers from a file

## Examples

### Basic Document with few document properties and paragraph

This short example show how to create Word Document with just one paragraph with Text and few document properties.

```csharp
string filePath = @"C:\Support\GitHub\PSWriteOffice\Examples\Documents\BasicDocument.docx";

using (WordDocument document = WordDocument.Create(filePath)) {
    document.Title = "This is my title";
    document.Creator = "Przemysław Kłys";
    document.Keywords = "word, docx, test";

    var paragraph = document.AddParagraph("Basic paragraph");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = SixLabors.ImageSharp.Color.Red;

    document.Save(true);
}
```

### Basic Document with Headers/Footers (first, odd, even)

This short example shows how to add headers and footers to Word Document.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
    document.AddParagraph("Test Section0");
    document.AddHeadersAndFooters();
    document.DifferentFirstPage = true;
    document.DifferentOddAndEvenPages = true;

    document.Sections[0].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
    document.Sections[0].Header.Default.AddParagraph().SetText("Test Section 0 - Header");
    document.Sections[0].Header.Even.AddParagraph().SetText("Test Section 0 - Even");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section1 = document.AddSection();
    section1.PageOrientation = PageOrientationValues.Portrait;
    section1.AddParagraph("Test Section1");
    section1.AddHeadersAndFooters();
    section1.Header.Default.AddParagraph().SetText("Test Section 1 - Header");
    section1.DifferentFirstPage = true;
    section1.Header.First.AddParagraph().SetText("Test Section 1 - First Header");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section2 = document.AddSection();
    section2.AddParagraph("Test Section2");
    section2.PageOrientation = PageOrientationValues.Landscape;
    section2.AddHeadersAndFooters();
    section2.Header.Default.AddParagraph().SetText("Test Section 2 - Header");

    document.AddParagraph("Test Section2 - Paragraph 1");

    var section3 = document.AddSection();
    section3.AddParagraph("Test Section3");
    section3.AddHeadersAndFooters();
    section3.Header.Default.AddParagraph().SetText("Test Section 3 - Header");

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Header.Default.Paragraphs[0].Text);
    document.Save(true);
}
```

### Advanced usage of OfficeIMO

This short example shows multiple features of `OfficeIMO.Word`

```csharp
string filePath = System.IO.Path.Combine(folderPath, "AdvancedDocument.docx");
using (WordDocument document = WordDocument.Create(filePath)) {
    // lets add some properties to the document
    document.BuiltinDocumentProperties.Title = "Cover Page Templates";
    document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
    document.ApplicationProperties.Company = "Evotec Services";

    // we force document to update fields on open, this will be used by TOC
    document.Settings.UpdateFieldsOnOpen = true;

    // lets add one of multiple added Cover Pages
    document.AddCoverPage(CoverPageTemplate.IonDark);

    // lets add Table of Content (1 of 2)
    document.AddTableOfContent(TableOfContentStyle.Template1);

    // lets add page break
    document.AddPageBreak();

    // lets create a list that will be binded to TOC
    var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

    wordListToc.AddItem("How to add a table to document?");

    document.AddParagraph("In the first paragraph I would like to show you how to add a table to the document using one of the 105 built-in styles:");

    // adding a table and modifying content
    var table = document.AddTable(5, 4, WordTableStyle.GridTable5DarkAccent5);
    table.Rows[3].Cells[2].Paragraphs[0].Text = "Adding text to cell";
    table.Rows[3].Cells[2].Paragraphs[0].Color = Color.Blue; ;
    table.Rows[3].Cells[3].Paragraphs[0].Text = "Different cell";

    document.AddParagraph("As you can see adding a table with some style, and adding content to it ").SetBold().SetUnderline(UnderlineValues.Dotted).AddText("is not really complicated").SetColor(Color.OrangeRed);

    wordListToc.AddItem("How to add a list to document?");

    var paragraph = document.AddParagraph("Adding lists is similar to ading a table. Just define a list and add list items to it. ").SetText("Remember that you can add anything between list items! ");
    paragraph.SetColor(Color.Blue).SetText("For example TOC List is just another list, but defining a specific style.");

    var list = document.AddList(WordListStyle.Bulleted);
    list.AddItem("First element of list", 0);
    list.AddItem("Second element of list", 1);

    var paragraphWithHyperlink = document.AddHyperLink("Go to Evotec Blogs", new Uri("https://evotec.xyz"), true, "URL with tooltip");
    // you can also change the hyperlink text, uri later on using properties
    paragraphWithHyperlink.Hyperlink.Uri = new Uri("https://evotec.xyz/hub");
    paragraphWithHyperlink.ParagraphAlignment = JustificationValues.Center;

    list.AddItem("3rd element of list, but added after hyperlink", 0);
    list.AddItem("4th element with hyperlink ").AddHyperLink("included.", new Uri("https://evotec.xyz/hub"), addStyle: true);

    document.AddParagraph();

    var listNumbered = document.AddList(WordListStyle.Heading1ai);
    listNumbered.AddItem("Different list number 1");
    listNumbered.AddItem("Different list number 2", 1);
    listNumbered.AddItem("Different list number 3", 1);
    listNumbered.AddItem("Different list number 4", 1);

    var section = document.AddSection();
    section.PageOrientation = PageOrientationValues.Landscape;
    section.PageSettings.PageSize = WordPageSize.A4;

    wordListToc.AddItem("Adding headers / footers");

    // lets add headers and footers
    document.AddHeadersAndFooters();

    // adding text to default header
    document.Header.Default.AddParagraph("Text added to header - Default");

    var section1 = document.AddSection();
    section1.PageOrientation = PageOrientationValues.Portrait;
    section1.PageSettings.PageSize = WordPageSize.A5;

    wordListToc.AddItem("Adding custom properties to document");

    document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
    document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
    document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

    document.Save(openWord);
}
```
