# OfficeIMO - Microsoft Word C# Library

<p align="center">
  <a href="https://dev.azure.com/evotecpl/OfficeIMO/_build/results?buildId=latest"><img src="https://dev.azure.com/evotecpl/OfficeIMO/_apis/build/status/EvotecIT.OfficeIMO"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg"></a>
</p>

<p align="center">
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/languages/top/evotecit/OfficeIMO.svg"></a>
  <a href="https://github.com/EvotecIT/OfficeIMO"><img src="https://img.shields.io/github/languages/code-size/evotecit/OfficeIMO.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

## What it's all about

This is a small project that allows to create Microsoft Word documents (.docx) using .NET Standard.
It was created because working with OpenXML is way too hard for me, and time consuming.
I originally created it for using within PowerShell module called PSWriteOffice, 
but thought it may be useful for others.
I used to use DocX library (which I co-authored) to create Word documents, 
but it only supports .NET Framework, and their newest community license makes the project unusuable.

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

## Learning resources: 

As I am not really a developer, and I hardly know what I'm doing I'm using a lot of different resources to make OfficeIMO useful.
Following resources may come useful to understand some concepts if you're going to dive into sources.

 - [Points, inches and Emus: Measuring units in Office Open XML](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)
 - [English Metric Units and Open XML](http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/)
 - [Open XML: add a picture](https://coders-corner.net/2015/04/11/open-xml-add-a-picture/)

## Relevant resources:

While not directly related I required this knowledge to get this up and running

 - [How do you use System.Drawing in .NET Core?](https://www.hanselman.com/blog/how-do-you-use-systemdrawing-in-net-core)