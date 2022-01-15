# OfficeIMO - Microsoft Word C# Library

<p align="center">
  <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build/results?buildId=latest"><img src="https://img.shields.io/azure-devops/tests/evotecpl/OfficeIMO/19?label=Tests%20Windows"></a>
  <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build/results?buildId=latest"><img src="https://img.shields.io/azure-devops/coverage/evotecpl/OfficeIMO/19"></a>
  <a href="https://www.nuget.org/packages/OfficeIMO"><img alt="Nuget" src="https://img.shields.io/nuget/dt/officeIMO?label=nuget%20downloads"></a>
  <a href="https://www.nuget.org/packages/OfficeIMO"><img alt="Nuget" src="https://img.shields.io/nuget/v/OfficeIMO"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.7.2-red.svg"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg"></a>
</p>

<p align="center">
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/languages/top/evotecit/OfficeIMO.svg"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/languages/code-size/evotecit/OfficeIMO.svg"></a>
  <a href="https://wakatime.com/badge/user/f1abc372-39bb-4b06-ad2b-3a24cf161f13/project/3cddaa3c-574a-400b-9870-d0973797eb51"><img src="https://wakatime.com/badge/user/f1abc372-39bb-4b06-ad2b-3a24cf161f13/project/3cddaa3c-574a-400b-9870-d0973797eb51.svg" alt="wakatime"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

## What it's all about

This is a small project (under development) that allows to create Microsoft Word documents (.docx) using .NET.
It was created because working with OpenXML is way too hard for me, and time consuming.
I originally created it for using within PowerShell module called PSWriteOffice,
but thought it may be useful for others.
I used to use DocX library (which I co-authored, before it was taken over by Xceed) to create Word documents,
but it only supports .NET Framework, and their newest community license makes the project unusuable.

*As I am not really a developer, and I hardly know what I'm doing if you know how to help out - please do.*

- If you see bad practice, please open and issue/submit PR.
- If you know how to do something in OpenXML that could help this project - please open an issue/submit PR
- If you see something that could work better - please open and issue/submit PR
- If you see something that I totally made a fool of myself - please open an issue/submit PR
- If you see something that works not the way I think it works - please open an issue/submit PR

I hope you get the drift? If it's bad - open an issue/fix it! I don't know what I'm doing!
The main thing is - it has to work with .NET Framework 4.7.2, .NET Standard 2.0 and so on.

**This project is under development and as such there's a lot of things that can and will change, especially if some people help out.**

## Features

- [x] Word basics
  - [x] Create
  - [x] Load
  - [x] Save
  - [ ] SaveAs (not working correcly in edge cases)
- [x] Word properties
  - [ ] Reading
  - [ ] Setting
- [x] Sections
  - [x] Add Paragraphs
  - [ ] Add Headers and Footers
  - [ ] Remove Headers and Footers
  - [ ] Remove Paragraphs
  - [ ] Remove Sections
- [ ] Headers and Footers in document (not including sections)
  - [ ] Add Default, Odd, Even, First
- [x] Paragraphs/Text and make it bold, underlined, colored and so on
- [x] Paragraphs and change alignment
- [ ] Tables
  - [ ] Add rows and columns
  - [ ] Add cells
  - [ ] Add cell properties
- [x] Images/Pictures (limited support - jpg only / inline type only)
  - [x] Add images from file to Word
  - [x] Save image from Word to File
  - [ ] Other image types
  - [ ] Other location types
- [ ] Hyperlinks
- [ ] Bookmarks
- [ ] Comments
- [ ] Fields
- [ ] Shapes
- [ ] Charts

## Examples

### Basic Document with few document properties and paragraph

This short example show how to create Word Document with just one paragraph with Text and few document properties.

```csharp
string filePath = "C:\\Support\\GitHub\\PSWriteOffice\\Examples\\Documents\\BasicDocument.docx";

using (WordDocument document = WordDocument.Create(filePath)) {
    document.Title = "This is my title";
    document.Creator = "Przemysław Kłys";
    document.Keywords = "word, docx, test";

    var paragraph = document.InsertParagraph("Basic paragraph");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    document.Save(true);
}
```

### Basic Document with Headers/Footers (first, odd, even)

This short example shows how to add headers and footers to Word Document.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddHeadersAndFooters();
    document.DifferentFirstPage = true;
    document.DifferentOddAndEvenPages = true;

    var paragraphInFooterFirst = document.Footer.First.InsertParagraph();
    paragraphInFooterFirst.Text = "This is a test on first";

    var count = document.Footer.First.Paragraphs.Count;

    var paragraphInFooterOdd = document.Footer.Odd.InsertParagraph();
    paragraphInFooterOdd.Text = "This is a test odd";


    var paragraphHeader = document.Header.Odd.InsertParagraph();
    paragraphHeader.Text = "Header - ODD";

    var paragraphInFooterEven = document.Footer.Even.InsertParagraph();
    paragraphInFooterEven.Text = "This is a test - Even";


    var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    paragraph = document.InsertPageBreak();
!
    paragraph = document.InsertParagraph("Basic paragraph - Page 2");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    paragraph = document.InsertPageBreak();

    paragraph = document.InsertParagraph("Basic paragraph - Page 3");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    paragraph = document.InsertPageBreak();

    paragraph = document.InsertParagraph("Basic paragraph - Page 4");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    paragraph = document.InsertPageBreak();

    paragraph = document.InsertParagraph("Basic paragraph - Page 5");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = System.Drawing.Color.Red.ToHexColor();

    paragraph = document.InsertPageBreak();

    document.Save(false);
}
```

## Learning resources:

I'm using a lot of different resources to make OfficeIMO useful. Following resources may come useful to understand some concepts if you're going to dive into sources.

 - [Packages and general (Open XML SDK)](https://docs.microsoft.com/en-us/office/open-xml/packages-and-general)
 - [Word processing (Open XML SDK)](https://docs.microsoft.com/en-us/office/open-xml/word-processing)
 - [Word 2007 Visual How Tos](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/gg537324(v=office.12)
 - [Office 2010 Visual How Tos](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff467945(v=office.14))
 - [Points, inches and Emus: Measuring units in Office Open XML](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)
 - [English Metric Units and Open XML](http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/)
 - [Open XML: add a picture](https://coders-corner.net/2015/04/11/open-xml-add-a-picture/)
 - [How do you use System.Drawing in .NET Core?](https://www.hanselman.com/blog/how-do-you-use-systemdrawing-in-net-core)
 - [How to add section break next page using openxml?](https://stackoverflow.com/questions/20040613/how-to-add-section-break-next-page-using-openxml)
 - [How to Preserve string with formatting in OpenXML Paragraph, Run, Text?](https://stackoverflow.com/questions/40246590/how-to-preserve-string-with-formatting-in-openxml-paragraph-run-text?rq=1)