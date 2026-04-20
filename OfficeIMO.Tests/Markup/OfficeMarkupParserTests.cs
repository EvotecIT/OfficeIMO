using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public class OfficeMarkupParserTests {
    [Fact]
    public void Parse_DocumentProfile_MapsCoreMarkdownAndDocumentExtensions() {
        var markup = """
# Title

Intro paragraph.

```officeimo
toc title="Contents" min=1 max=3
```

```officeimo
page-break
```
""";

        var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Document
        });

        Assert.False(result.HasErrors);
        Assert.IsType<OfficeMarkupHeadingBlock>(result.Document.Blocks[0]);
        Assert.IsType<OfficeMarkupParagraphBlock>(result.Document.Blocks[1]);
        Assert.IsType<OfficeMarkupTableOfContentsBlock>(result.Document.Blocks[2]);
        Assert.IsType<OfficeMarkupPageBreakBlock>(result.Document.Blocks[3]);
    }

    [Fact]
    public void Parse_DocumentProfile_PreservesMarkdownAfterSingleLineDirectives() {
        var markup = """
---
profile: document
---

# Architecture Note

::toc min=1 max=3 title="Contents"

::header text="OfficeIMO Markup"

This document starts as normal Markdown.

## Overview

- Markdown remains readable
- Office-specific constructs are explicit

| Area | Status |
| --- | --- |
| Parser | Ready |

::chart type=column title="Quarterly Revenue"
Quarter,Revenue
Q1,120

::page-break

::footer text="Generated from OfficeIMO Markup"
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        Assert.Collection(result.Document.Blocks,
            block => Assert.IsType<OfficeMarkupHeadingBlock>(block),
            block => Assert.IsType<OfficeMarkupTableOfContentsBlock>(block),
            block => Assert.IsType<OfficeMarkupHeaderFooterBlock>(block),
            block => Assert.IsType<OfficeMarkupParagraphBlock>(block),
            block => Assert.IsType<OfficeMarkupHeadingBlock>(block),
            block => Assert.IsType<OfficeMarkupListBlock>(block),
            block => Assert.IsType<OfficeMarkupTableBlock>(block),
            block => Assert.IsType<OfficeMarkupChartBlock>(block),
            block => Assert.IsType<OfficeMarkupPageBreakBlock>(block),
            block => Assert.IsType<OfficeMarkupHeaderFooterBlock>(block));
    }

    [Fact]
    public void Parse_PresentationProfile_MapsSlideBodyThroughMarkdown() {
        var markup = """
```officeimo
slide title="Roadmap" layout=TitleAndContent transition=Fade columns=2

## Now

- Build AST
- Emit code
```
""";

        var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Presentation
        });

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal("Roadmap", slide.Title);
        Assert.Equal("TitleAndContent", slide.Layout);
        Assert.Equal("Fade", slide.Transition);
        Assert.Equal(2, slide.Columns);
        Assert.IsType<OfficeMarkupHeadingBlock>(slide.Blocks[0]);
        Assert.IsType<OfficeMarkupListBlock>(slide.Blocks[1]);
    }

    [Fact]
    public void Parse_PresentationProfile_MapsFrontMatterSlideSeparatorsAndDirectives() {
        var markup = """
---
profile: presentation
title: OfficeIMO Slide DSL Demo
theme: evotec-modern
---

# Executive Summary

@slide {
  layout: title-and-content
  transition: fade duration=0.6
}

- Revenue increased
- Costs stabilized

::notes
Use this as the opening business summary.

---

# Architecture Overview

@slide {
  layout: blank
}

::mermaid
graph LR
  A[Markup] --> B[AST]
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        Assert.Equal(OfficeMarkupProfile.Presentation, result.Document.Profile);
        Assert.Equal("OfficeIMO Slide DSL Demo", result.Document.Metadata["title"]);
        Assert.Equal(2, result.Document.Blocks.Count);

        var firstSlide = Assert.IsType<OfficeMarkupSlideBlock>(result.Document.Blocks[0]);
        Assert.Equal("Executive Summary", firstSlide.Title);
        Assert.Equal("title-and-content", firstSlide.Layout);
        Assert.Equal("fade duration=0.6", firstSlide.Transition);
        Assert.Equal("Use this as the opening business summary.", firstSlide.Notes);
        Assert.IsType<OfficeMarkupListBlock>(Assert.Single(firstSlide.Blocks));

        var secondSlide = Assert.IsType<OfficeMarkupSlideBlock>(result.Document.Blocks[1]);
        Assert.Equal("Architecture Overview", secondSlide.Title);
        Assert.IsType<OfficeMarkupDiagramBlock>(Assert.Single(secondSlide.Blocks));
    }

    [Fact]
    public void Parse_PresentationProfile_PreservesDirectiveLikeLinesInsideFencedCodeBlocks() {
        var markup = """
---
profile: presentation
---

# Demo

@slide {
  layout: blank
}

```bat
@echo off
::chart type=column
echo done
```
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var code = Assert.IsType<OfficeMarkupCodeBlock>(Assert.Single(slide.Blocks));
        Assert.Equal("bat", code.Language);
        Assert.Contains("@echo off", code.Content, StringComparison.Ordinal);
        Assert.Contains("::chart type=column", code.Content, StringComparison.Ordinal);
        Assert.DoesNotContain(slide.Blocks, block => block is OfficeMarkupChartBlock or OfficeMarkupExtensionBlock);
    }

    [Fact]
    public void Parse_PresentationProfile_DoesNotSplitSlidesOnFencedSeparatorContent() {
        var markup = """
---
profile: presentation
---

# Demo

@slide {
  layout: blank
}

```yaml
---
name: sample
---
```
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var code = Assert.IsType<OfficeMarkupCodeBlock>(Assert.Single(slide.Blocks));
        Assert.Equal("yaml", code.Language);
        Assert.Contains("name: sample", code.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_PresentationProfile_DoesNotConsumeFencedSlideExamplesAsSlideMetadata() {
        var markup = """
---
profile: presentation
---

# Demo

@slide {
  layout: blank
}

```bat
@slide {
  layout: section
}
```
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal("blank", slide.Layout);
        var code = Assert.IsType<OfficeMarkupCodeBlock>(Assert.Single(slide.Blocks));
        Assert.Contains("@slide {", code.Content, StringComparison.Ordinal);
        Assert.Contains("layout: section", code.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_WorkbookProfile_MapsAtSheetAndColonRangeFormula() {
        var markup = """
---
profile: workbook
title: Revenue Model
---

@sheet {
  name: Summary
}

::range address=A1
Quarter,Revenue
Q1,120

::formula cell=C2
=B2*2
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        Assert.Equal(OfficeMarkupProfile.Workbook, result.Document.Profile);
        Assert.IsType<OfficeMarkupSheetBlock>(result.Document.Blocks[0]);
        var range = Assert.IsType<OfficeMarkupRangeBlock>(result.Document.Blocks[1]);
        Assert.Equal("A1", range.Address);
        Assert.Equal(2, range.Values.Count);
        var formula = Assert.IsType<OfficeMarkupFormulaBlock>(result.Document.Blocks[2]);
        Assert.Equal("C2", formula.Cell);
        Assert.Equal("=B2*2", formula.Expression);
    }

    [Fact]
    public void Parse_PresentationProfile_MapsColonChartData() {
        var markup = """
---
profile: presentation
---

# Data

@slide {
  layout: blank
}

::chart type=column title="Quarterly Revenue"
Quarter,Revenue,Costs
Q1,120,80
Q2,180,95
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var chart = Assert.IsType<OfficeMarkupChartBlock>(Assert.Single(slide.Blocks));
        Assert.Equal("column", chart.ChartType);
        Assert.Equal("Quarterly Revenue", chart.Title);
        Assert.Equal(3, chart.Data.Count);
        Assert.Equal(new[] { "Quarter", "Revenue", "Costs" }, chart.Data[0]);
    }

    [Fact]
    public void Parse_PresentationProfile_MapsLayoutDirectivesAsSemanticNodes() {
        var markup = """
---
profile: presentation
---

# Layout

@slide {
  layout: blank
}

::textbox x=6% y=8% w=70% h=10% style=hero-title
Pipeline: From Text to PPTX

::columns gap=4%

::column width=48%
## Flow

::card title="Key Outcomes"
- Parser typed layout nodes
- Exporter keeps real PPTX output
""";

        var result = OfficeMarkupParser.Parse(markup);

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var textBox = Assert.IsType<OfficeMarkupTextBoxBlock>(slide.Blocks[0]);
        Assert.Equal("6%", textBox.Placement?.X);
        Assert.Equal("70%", textBox.Placement?.Width);
        var columns = Assert.IsType<OfficeMarkupColumnsBlock>(slide.Blocks[1]);
        Assert.Equal("4%", columns.Gap);
        Assert.IsType<OfficeMarkupColumnBlock>(slide.Blocks[2]);
        Assert.IsType<OfficeMarkupCardBlock>(slide.Blocks[3]);
    }

    [Fact]
    public void StyleResolver_AppliesBuiltInStylesAndAttributeOverrides() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
theme: evotec-modern
---

# Styled

@slide {
  layout: blank
}

::textbox style=hero-title color=#FFFFFF font-size=34
Pipeline
""");

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var textBox = Assert.IsType<OfficeMarkupTextBoxBlock>(Assert.Single(slide.Blocks));
        var style = OfficeMarkupStyleResolver.Create(result.Document).Resolve(textBox);

        Assert.NotNull(style);
        Assert.Equal("hero-title", style!.Name);
        Assert.Equal(34, style.FontSize);
        Assert.True(style.Bold);
        Assert.Equal("#FFFFFF", style.TextColor);
    }

    [Fact]
    public void Validate_RejectsDocumentExtensionInPresentationProfile() {
        var result = OfficeMarkupParser.Parse("""
```officeimo
page-break
```
""", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Presentation
        });

        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Severity == OfficeMarkupDiagnosticSeverity.Error
            && diagnostic.Message.Contains("PageBreak", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_RejectsPresentationLayoutNodeInWorkbookProfile() {
        var result = OfficeMarkupParser.Parse("""
::textbox
Workbook should not accept slide textboxes.
""", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Workbook
        });

        Assert.True(result.HasErrors);
        var diagnostic = Assert.Single(result.Diagnostics, diagnostic =>
            diagnostic.Severity == OfficeMarkupDiagnosticSeverity.Error
            && diagnostic.Message.Contains("TextBox", StringComparison.Ordinal));
        Assert.NotNull(diagnostic.Node);
        Assert.Contains("::textbox", diagnostic.Node!.SourceText, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_WorkbookProfile_MapsSheetRangeTableChartAndFormula() {
        var markup = """
```officeimo
sheet name="Revenue"
```

```officeimo
range address=A1:B3
Product,Revenue
A,120
```

```officeimo
formula cell=B4 value="=SUM(B2:B3)"
```

```officeimo
table name="RevenueTable" range=A1:B3 header=true
```

```officeimo
chart type=column title="Revenue" source=A1:B3
```
""";

        var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Workbook
        });

        Assert.False(result.HasErrors);
        Assert.IsType<OfficeMarkupSheetBlock>(result.Document.Blocks[0]);
        var range = Assert.IsType<OfficeMarkupRangeBlock>(result.Document.Blocks[1]);
        Assert.Equal("A1:B3", range.Address);
        Assert.Equal(2, range.Values.Count);
        Assert.IsType<OfficeMarkupFormulaBlock>(result.Document.Blocks[2]);
        Assert.IsType<OfficeMarkupNamedTableBlock>(result.Document.Blocks[3]);
        Assert.IsType<OfficeMarkupChartBlock>(result.Document.Blocks[4]);
    }

    [Fact]
    public void Parse_MermaidFence_BecomesCommonDiagramNode() {
        var result = OfficeMarkupParser.Parse("""
```mermaid
flowchart LR
  A --> B
```
""", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Presentation
        });

        Assert.False(result.HasErrors);
        var diagram = Assert.IsType<OfficeMarkupDiagramBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal("mermaid", diagram.Language);
        Assert.Contains("A --> B", diagram.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_PresentationProfile_MapsMermaidPlacement() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Diagram

@slide {
  layout: blank
}

::mermaid x=8% y=18% w=60% h=44% fit=stretch
flowchart LR
  A --> B
""");

        Assert.False(result.HasErrors);
        var slide = Assert.IsType<OfficeMarkupSlideBlock>(Assert.Single(result.Document.Blocks));
        var diagram = Assert.IsType<OfficeMarkupDiagramBlock>(Assert.Single(slide.Blocks));
        Assert.Equal("8%", diagram.Placement?.X);
        Assert.Equal("60%", diagram.Placement?.Width);
        Assert.Equal("stretch", diagram.Attributes["fit"]);
    }


    [Fact]
    public void CSharpEmitter_UsesProfileSpecificEntryPoint() {
        var result = OfficeMarkupParser.Parse("""
```officeimo
slide title="Roadmap"
```
""", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Presentation
        });

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains("PowerPointPresentation.Create", code, StringComparison.Ordinal);
        Assert.Contains("presentation.AddSlide", code, StringComparison.Ordinal);
        Assert.Contains("Roadmap", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_UsesWorkbookCommands() {
        var result = OfficeMarkupParser.Parse("""
```officeimo
sheet name="Revenue"
```
""", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Workbook
        });

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("New-OfficeExcelWorkbook", code, StringComparison.Ordinal);
        Assert.Contains("Add-OfficeExcelWorksheet", code, StringComparison.Ordinal);
        Assert.Contains("Revenue", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_PresentationWrapsTopLevelMarkdownAsImplicitSlides() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Quarterly Review

- Revenue grew 18%
- Churn improved
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains("PowerPointSlide slide1 = presentation.AddSlide();", code, StringComparison.Ordinal);
        Assert.Contains("slide1.AddTextBox(@\"Quarterly Review\");", code, StringComparison.Ordinal);
        Assert.Contains("Revenue grew 18%", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Presentation-level Heading", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Presentation-level List", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_PresentationWrapsTopLevelMarkdownAsImplicitSlides() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Quarterly Review

- Revenue grew 18%
- Churn improved
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("$slide1 = Add-OfficePowerPointSlide -Presentation $presentation", code, StringComparison.Ordinal);
        Assert.Contains("Add-OfficePowerPointText -Slide $slide1 -Text 'Quarterly Review'", code, StringComparison.Ordinal);
        Assert.Contains("Revenue grew 18%", code, StringComparison.Ordinal);
        Assert.DoesNotContain("# Heading:", code, StringComparison.Ordinal);
        Assert.DoesNotContain("# List:", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_PresentationPreservesChartsNotesAndTransitionDetails() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Data

@slide {
  layout: blank
  transition: fade duration=0.6 speed=fast advance-on-click=false advance-after=5
}

::chart type=column title="Revenue" x=8% y=20% w=80% h=55% category-title=Quarter value-title=Amount value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
Quarter,Revenue,Costs
Q1,120,80
Q2,180,95

::notes
Explain the revenue trend.
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains("SlideTransition.Fade", code, StringComparison.Ordinal);
        Assert.DoesNotContain("SlideTransition.FadeDuration", code, StringComparison.Ordinal);
        Assert.Contains("slide1.TransitionSpeed = SlideTransitionSpeed.Fast;", code, StringComparison.Ordinal);
        Assert.Contains("slide1.TransitionDurationSeconds = 0.6;", code, StringComparison.Ordinal);
        Assert.Contains("slide1.TransitionAdvanceOnClick = false;", code, StringComparison.Ordinal);
        Assert.Contains("slide1.TransitionAdvanceAfterSeconds = 5;", code, StringComparison.Ordinal);
        Assert.Contains("Transition details", code, StringComparison.Ordinal);
        Assert.Contains("Transition effect: @\"fade\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition native enum: @\"Fade\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition duration: @\"0.6\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition speed: @\"fast\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition advance-on-click: @\"false\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition advance-after: @\"5\"", code, StringComparison.Ordinal);
        Assert.Contains("new PowerPointChartData", code, StringComparison.Ordinal);
        Assert.Contains("new PowerPointChartSeries(@\"Revenue\"", code, StringComparison.Ordinal);
        Assert.Contains("var chart1 = slide1.AddChart(chartData1)", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetTitle(@\"Revenue\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetCategoryAxisTitle(@\"Quarter\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetValueAxisTitle(@\"Amount\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetValueAxisNumberFormat(@\"#,##0\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetLegend(C.LegendPositionValues.Right)", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetDataLabelPosition(C.DataLabelPositionValues.OutsideEnd)", code, StringComparison.Ordinal);
        Assert.Contains("slide1.Notes.Text", code, StringComparison.Ordinal);
        Assert.Contains("Placement: x=@\"8%\"", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_PresentationMapsDirectionalTransitionVariants() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Push

@slide {
  layout: blank
  transition: push direction=left duration=0.5
}

Body

---

# Warp

@slide {
  layout: blank
  transition: warp direction=out duration=0.7
}

Body
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains("slide1.Transition = SlideTransition.PushLeft;", code, StringComparison.Ordinal);
        Assert.Contains("slide2.Transition = SlideTransition.WarpOut;", code, StringComparison.Ordinal);
        Assert.DoesNotContain("SlideTransition.Push;", code, StringComparison.Ordinal);
        Assert.Contains("Transition details: @\"push direction=left duration=0.5\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition details: @\"warp direction=out duration=0.7\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition native enum: @\"PushLeft\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition direction: @\"left\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition duration: @\"0.5\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition native enum: @\"WarpOut\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition direction: @\"out\"", code, StringComparison.Ordinal);
        Assert.Contains("Transition duration: @\"0.7\"", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_WorkbookUsesRangeAddressFormulaAndNativeChartCall() {
        var result = OfficeMarkupParser.Parse("""
---
profile: workbook
---

@sheet {
  name: Revenue
}

::range address=B3
Product,2024
A,120

::formula cell=D4
=C4-B4

::chart type=column title="Revenue" source=B3:C4 cell=F2 width=480 height=320 category-title=Product value-title=Revenue value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains(@"GetOrAddSheet(@""Sheet1"").CellValue(3, 2, @""Product"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Sheet1"").CellValue(4, 3, @""120"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Sheet1"").CellFormula(4, 4, @""=C4-B4"")", code, StringComparison.Ordinal);
        Assert.Contains(@"var chart1 = GetOrAddSheet(@""Sheet1"").AddChartFromRange(@""B3:C4"", row: 2, column: 6", code, StringComparison.Ordinal);
        Assert.Contains("ExcelChartType.ColumnClustered", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetCategoryAxisTitle(@\"Product\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetValueAxisTitle(@\"Revenue\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetValueAxisNumberFormat(@\"#,##0\")", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetLegend(C.LegendPositionValues.Right)", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetDataLabels(showValue: true", code, StringComparison.Ordinal);
        Assert.Contains("C.DataLabelPositionValues.OutsideEnd", code, StringComparison.Ordinal);
        Assert.Contains("chart1.SetValueAxisGridlines(showMajor: true", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_WorkbookCreatesDefaultSheetForUnqualifiedTargetsBeforeAnySheetDirective() {
        var result = OfficeMarkupParser.Parse("""
---
profile: workbook
---

::range address=A1
Metric,Value
Revenue,120

::formula cell=C2
=B2*2
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains(@"GetOrAddSheet(@""Sheet1"").CellValue(1, 1, @""Metric"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Sheet1"").CellFormula(2, 3, @""=B2*2"")", code, StringComparison.Ordinal);
        Assert.DoesNotContain("sheet!", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CSharpEmitter_WorkbookPreservesSheetQualifiedTargets() {
        var result = OfficeMarkupParser.Parse("""
---
profile: workbook
---

::range address=Data!A1
Metric,Value,Double
Revenue,120,

::formula cell=Data!C2
=B2*2

::format target=Data!B2:C2 numberFormat="#,##0" fill=#D9EAD3 color=#112233 bold=true italic=true underline=true align=center valign=middle wrap=true border=thin border-color=#445566

::table name="DataTable" range=Data!A1:C2 header=true

::chart type=column title="Qualified Chart" source=Data!DataTable cell=Dashboard!B2 width=480 height=320
""");

        var code = new OfficeMarkupCSharpEmitter().Emit(result.Document);

        Assert.Contains("ExcelSheet GetOrAddSheet(string name)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellValue(1, 1, @""Metric"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellFormula(2, 3, @""=B2*2"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").FormatCell(2, 2, @""#,##0"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellBackground(2, 2, @""#D9EAD3"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellFontColor(2, 2, @""#112233"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellBold(2, 2, true)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellItalic(2, 2, true)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellUnderline(2, 2, true)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellVerticalAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").WrapCells(2, 2, 2)", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").CellBorder(2, 2, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, @""#445566"")", code, StringComparison.Ordinal);
        Assert.Contains(@"GetOrAddSheet(@""Data"").AddTable(@""A1:C2""", code, StringComparison.Ordinal);
        Assert.Contains(@"var chartSourceSheet1 = GetOrAddSheet(@""Data"")", code, StringComparison.Ordinal);
        Assert.Contains(@"chartSourceSheet1.GetTableRange(@""DataTable"")", code, StringComparison.Ordinal);
        Assert.Contains(@"var chart1 = GetOrAddSheet(@""Dashboard"").AddChart(chartDataRange1, row: 2, column: 2", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_WorkbookPreservesSheetQualifiedTargets() {
        var result = OfficeMarkupParser.Parse("""
---
profile: workbook
---

::range address=Data!A1
Metric,Value,Double
Revenue,120,

::formula cell=Data!C2
=B2*2

::format target=Data!B2:C2 numberFormat="#,##0" fill=#D9EAD3 color=#112233 bold=true italic=true underline=true align=center valign=middle wrap=true border=thin border-color=#445566

::table name="DataTable" range=Data!A1:C2 header=true

::chart type=column title="Qualified Chart" source=Data!DataTable cell=Dashboard!B2 width=480 height=320
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("function Get-OrAddOfficeExcelWorksheet", code, StringComparison.Ordinal);
        Assert.Contains("Set-OfficeExcelCell -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data') -Row 1 -Column 1 -Value 'Metric'", code, StringComparison.Ordinal);
        Assert.Contains("Set-OfficeExcelFormula -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data') -Cell 'C2' -Formula '=B2*2'", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').FormatCell(2, 2, '#,##0')", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellBackground(2, 2, '#D9EAD3')", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellFontColor(2, 2, '#112233')", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellBold(2, 2, $true)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellItalic(2, 2, $true)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellUnderline(2, 2, $true)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellAlign(2, 2, [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::Center)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellVerticalAlign(2, 2, [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::Center)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').WrapCells(2, 2, 2)", code, StringComparison.Ordinal);
        Assert.Contains("(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data').CellBorder(2, 2, [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::Thin, '#445566')", code, StringComparison.Ordinal);
        Assert.Contains("Add-OfficeExcelTable -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Data') -Range 'A1:C2' -Name 'DataTable'", code, StringComparison.Ordinal);
        Assert.Contains("Add-OfficeExcelChart -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Dashboard') -Type 'column' -Source 'Data!DataTable' -Row 2 -Column 2", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_WorkbookCreatesDefaultSheetForUnqualifiedTargetsBeforeAnySheetDirective() {
        var result = OfficeMarkupParser.Parse("""
---
profile: workbook
---

::range address=A1
Metric,Value
Revenue,120

::formula cell=C2
=B2*2
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("Set-OfficeExcelCell -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Sheet1') -Row 1 -Column 1 -Value 'Metric'", code, StringComparison.Ordinal);
        Assert.Contains("Set-OfficeExcelFormula -Worksheet (Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name 'Sheet1') -Cell 'C2' -Formula '=B2*2'", code, StringComparison.Ordinal);
        Assert.DoesNotContain("-Worksheet $sheet", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_CommentsMultilineColumnBodiesSafely() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Layout

@slide {
  layout: blank
}

::column width=48%
## Notes
- One
- Two
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("# ## Notes", code, StringComparison.Ordinal);
        Assert.Contains("# - One", code, StringComparison.Ordinal);
        Assert.DoesNotContain("\n- One", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_EmitsConcreteInlineChartData() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Data

@slide {
  layout: blank
}

::chart type=column title="Revenue"
Quarter,Revenue,Costs
Q1,120,80
Q2,180,95
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("$chartData1 = @{", code, StringComparison.Ordinal);
        Assert.Contains("Categories = @('Q1', 'Q2')", code, StringComparison.Ordinal);
        Assert.Contains("Name = 'Revenue'", code, StringComparison.Ordinal);
        Assert.Contains("Values = @(120, 180)", code, StringComparison.Ordinal);
        Assert.Contains("-Data $chartData1", code, StringComparison.Ordinal);
        Assert.DoesNotContain("-Data $chartData\r", code, StringComparison.Ordinal);
        Assert.DoesNotContain("-Data $chartData\n", code, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellEmitter_PresentationMapsDirectionalTransitionVariants() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Push

@slide {
  layout: blank
  transition: push direction=up duration=0.5 speed=slow advance-on-click=true
}

Body

---

# Ferris

@slide {
  layout: blank
  transition: ferris direction=right duration=0.8
}

Body
""");

        var code = new OfficeMarkupPowerShellEmitter().Emit(result.Document);

        Assert.Contains("$slide1.Transition = [OfficeIMO.PowerPoint.SlideTransition]::PushUp", code, StringComparison.Ordinal);
        Assert.Contains("$slide2.Transition = [OfficeIMO.PowerPoint.SlideTransition]::FerrisRight", code, StringComparison.Ordinal);
        Assert.DoesNotContain("[OfficeIMO.PowerPoint.SlideTransition]::Push ", code, StringComparison.Ordinal);
        Assert.Contains("$slide1.TransitionSpeed = [OfficeIMO.PowerPoint.SlideTransitionSpeed]::Slow", code, StringComparison.Ordinal);
        Assert.Contains("$slide1.TransitionDurationSeconds = 0.5", code, StringComparison.Ordinal);
        Assert.Contains("$slide1.TransitionAdvanceOnClick = $true", code, StringComparison.Ordinal);
        Assert.Contains("# Transition details: push direction=up duration=0.5", code, StringComparison.Ordinal);
        Assert.Contains("# Transition details: ferris direction=right duration=0.8", code, StringComparison.Ordinal);
        Assert.Contains("# Transition native enum: PushUp", code, StringComparison.Ordinal);
        Assert.Contains("# Transition direction: up", code, StringComparison.Ordinal);
        Assert.Contains("# Transition duration: 0.5", code, StringComparison.Ordinal);
        Assert.Contains("# Transition speed: slow", code, StringComparison.Ordinal);
        Assert.Contains("# Transition advance-on-click: true", code, StringComparison.Ordinal);
        Assert.Contains("# Transition native enum: FerrisRight", code, StringComparison.Ordinal);
        Assert.Contains("# Transition direction: right", code, StringComparison.Ordinal);
        Assert.Contains("# Transition duration: 0.8", code, StringComparison.Ordinal);
    }

    [Fact]
    public void TransitionResolver_ParsesDirectionalAndTimingAttributes() {
        var resolved = OfficeMarkupTransitionResolver.Parse("push direction=left duration=0.5");

        Assert.Equal("push", resolved.Effect);
        Assert.Equal("PushLeft", resolved.ResolvedIdentifier);
        Assert.True(resolved.HasArguments);
        Assert.Equal("left", resolved.Attributes["direction"]);
        Assert.Equal("0.5", resolved.Attributes["duration"]);
    }

    [Fact]
    public void PowerPointExporter_CreatesOpenablePresentation() {
        var markup = """
---
profile: presentation
title: Export Demo
---

# Quarterly Review

@slide {
  layout: title-and-content
  transition: fade
}

- Revenue grew 18%
- Churn improved

::notes
Open with the top-line result.

---

# Architecture Overview

@slide {
  layout: blank
}

::textbox x=6% y=8% w=70% h=10% style=hero-title
Pipeline: From Text to PPTX

::mermaid x=6% y=24% w=60% h=42%
flowchart LR
  Markup --> AST
  AST --> PPTX

---

# Data

@slide {
  layout: blank
}

::chart type=column title="Quarterly Revenue" x=10% y=22% w=80% h=58% category-title=Quarter value-title=Revenue value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
Quarter,Revenue
Q1,120
Q2,180
Q3,260
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            Assert.True(File.Exists(path));
            using (var presentation = PowerPointPresentation.Open(path)) {
                Assert.Equal(3, presentation.Slides.Count);
                Assert.Equal(SlideTransition.Fade, presentation.Slides[0].Transition);
                Assert.Equal("Open with the top-line result.", presentation.Slides[0].Notes.Text);
                Assert.Contains(presentation.Slides[0].Shapes.OfType<PowerPointTextBox>(), box => box.Text.Contains("Quarterly Review", StringComparison.Ordinal));
                Assert.Contains(presentation.Slides[0].Shapes.OfType<PowerPointAutoShape>(), shape =>
                    shape.Name.Contains("Summary Card", StringComparison.Ordinal));
                Assert.Contains(presentation.Slides[0].Shapes.OfType<PowerPointAutoShape>(), shape =>
                    shape.Name.Contains("Canvas Rail", StringComparison.Ordinal));
                Assert.Contains(presentation.Slides[1].Shapes.OfType<PowerPointTextBox>(), box =>
                    box.Text.Contains("Pipeline: From Text to PPTX", StringComparison.Ordinal)
                    && box.FontSize == 32
                    && box.Bold
                    && box.TextAutoFit == PowerPointTextAutoFit.Normal);
                Assert.Contains(presentation.Slides[1].Shapes.OfType<PowerPointAutoShape>(), shape =>
                    shape.Name.Contains("Canvas Wash", StringComparison.Ordinal));
                Assert.Contains(presentation.Slides[1].Shapes.OfType<PowerPointAutoShape>(), shape =>
                    shape.Name.Contains("Diagram Panel", StringComparison.Ordinal));
                Assert.DoesNotContain(presentation.Slides[1].Shapes.OfType<PowerPointTextBox>(), box =>
                    box.Text.Contains("Architecture Overview", StringComparison.Ordinal));
                Assert.Contains(presentation.Slides[1].Shapes.OfType<PowerPointTextBox>(), box =>
                    box.Text.Contains("Mermaid diagram", StringComparison.OrdinalIgnoreCase)
                    && box.Text.Contains("Mermaid renderer", StringComparison.Ordinal));
                Assert.DoesNotContain(presentation.Slides[1].Shapes.OfType<PowerPointTextBox>(), box =>
                    box.Text.Contains("Markup --> AST", StringComparison.Ordinal)
                    || box.Text.Contains("AST --> PPTX", StringComparison.Ordinal));
                Assert.NotEmpty(presentation.Slides[2].Charts);
                Assert.Contains(presentation.Slides[2].Shapes.OfType<PowerPointAutoShape>(), shape =>
                    shape.Name.Contains("Chart Panel", StringComparison.Ordinal));
            }

            using var package = PresentationDocument.Open(path, false);
            var validationErrors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package);
            Assert.Empty(validationErrors);
            var chartXml = package.PresentationPart!.SlideParts.SelectMany(part => part.ChartParts).First().ChartSpace!.OuterXml;
            Assert.Contains("Quarter", chartXml, StringComparison.Ordinal);
            Assert.Contains("Revenue", chartXml, StringComparison.Ordinal);
            Assert.Contains("#,##0", chartXml, StringComparison.Ordinal);
            Assert.Contains("r\"", chartXml, StringComparison.Ordinal);
            Assert.Contains("outEnd", chartXml, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_WrapsPresentationMarkdownWithoutExplicitSlides() {
        var markup = """
---
profile: presentation
---

# Quarterly Review

- Revenue grew 18%
- Churn improved
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            Assert.True(File.Exists(path));
            using var presentation = PowerPointPresentation.Open(path);
            Assert.Single(presentation.Slides);
            Assert.Contains(presentation.Slides[0].Shapes.OfType<PowerPointTextBox>(), box => box.Text.Contains("Quarterly Review", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_ComposesSemanticTwoColumnSlidesWithDesignerPanels() {
        var markup = """
---
profile: presentation
theme: evotec-modern
---

# Architecture Overview

@slide {
  layout: two-columns
}

Stay in semantic Markdown until the slide actually needs more control.

::columns gap=4%

::column width=48%
## Flow
- Parser builds the semantic AST
- Emitters target C# and PowerShell
- Exporters produce Office files

::column width=48%
## Why it helps
- Authoring stays readable
- Layout is still intentional
- Generated code remains the escape hatch
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using var presentation = PowerPointPresentation.Open(path);
            var slide = Assert.Single(presentation.Slides);

            Assert.Contains(slide.Shapes.OfType<PowerPointTextBox>(), box =>
                box.Text.Contains("Architecture Overview", StringComparison.Ordinal));
            Assert.Contains(slide.Shapes.OfType<PowerPointTextBox>(), box =>
                box.Text.Contains("Flow", StringComparison.Ordinal));
            Assert.Contains(slide.Shapes.OfType<PowerPointTextBox>(), box =>
                box.Text.Contains("Why it helps", StringComparison.Ordinal));
            Assert.Equal(2, slide.Shapes.OfType<PowerPointAutoShape>()
                .Count(shape => shape.Name.Contains("Semantic Column Panel", StringComparison.Ordinal)));
            Assert.DoesNotContain(slide.Shapes.OfType<PowerPointAutoShape>(), shape =>
                shape.Name.Contains("Canvas Rail", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ParserAndEmitters_PreservePresentationSectionMetadata() {
        var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Intro

@slide {
  section: Introduction
}

Hello

---

# Deep Dive

@slide {
  section: Architecture
}

World
""");

        var slides = result.Document.Blocks.OfType<OfficeMarkupSlideBlock>().ToList();
        Assert.Equal(new[] { "Introduction", "Architecture" }, slides.Select(slide => slide.Section).ToArray());

        var csharp = new OfficeMarkupCSharpEmitter().Emit(result.Document);
        Assert.Contains("presentation.AddSection(@\"Introduction\", startSlideIndex: 0);", csharp, StringComparison.Ordinal);
        Assert.Contains("presentation.AddSection(@\"Architecture\", startSlideIndex: 1);", csharp, StringComparison.Ordinal);

        var powershell = new OfficeMarkupPowerShellEmitter().Emit(result.Document);
        Assert.Contains("$null = $presentation.AddSection('Introduction', 0)", powershell, StringComparison.Ordinal);
        Assert.Contains("$null = $presentation.AddSection('Architecture', 1)", powershell, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointExporter_CreatesPresentationSectionsFromSlideMetadata() {
        var markup = """
---
profile: presentation
---

# Intro

@slide {
  section: Introduction
}

Hello

---

# Summary

@slide {
  section: Introduction
}

Still intro

---

# Architecture

@slide {
  section: Deep Dive
}

Details
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using var presentation = PowerPointPresentation.Open(path);
            var sections = presentation.GetSections().ToArray();

            Assert.Equal(new[] { "Introduction", "Deep Dive" }, sections.Select(section => section.Name).ToArray());
            Assert.Equal(new[] { 0, 1 }, sections[0].SlideIndices.ToArray());
            Assert.Equal(new[] { 2 }, sections[1].SlideIndices.ToArray());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_MapsDirectionalTransitionDetailsToNativeTransitions() {
        var markup = """
---
profile: presentation
---

# Push

@slide {
  layout: blank
  transition: push direction=left duration=0.5 speed=fast advance-on-click=false advance-after=4
}

Body

---

# Warp

@slide {
  layout: blank
  transition: warp direction=out duration=0.7
}

Body

---

# Ferris

@slide {
  layout: blank
  transition: ferris direction=right
}

Body
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using var presentation = PowerPointPresentation.Open(path);
            Assert.Equal(3, presentation.Slides.Count);
            Assert.Equal(SlideTransition.PushLeft, presentation.Slides[0].Transition);
            Assert.Equal(SlideTransitionSpeed.Fast, presentation.Slides[0].TransitionSpeed);
            Assert.Equal(0.5, presentation.Slides[0].TransitionDurationSeconds);
            Assert.False(presentation.Slides[0].TransitionAdvanceOnClick);
            Assert.Equal(4.0, presentation.Slides[0].TransitionAdvanceAfterSeconds);
            Assert.Equal(SlideTransition.WarpOut, presentation.Slides[1].Transition);
            Assert.Equal(0.7, presentation.Slides[1].TransitionDurationSeconds);
            Assert.Equal(SlideTransition.FerrisRight, presentation.Slides[2].Transition);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_AppliesRelativeBackgroundImageOverlayAndSkipsFallbackCanvas() {
        var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var imageSource = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
        var localImage = Path.Combine(tempDirectory, "BackgroundImage.png");
        File.Copy(imageSource, localImage, overwrite: true);

        var markup = """
---
profile: presentation
---

# Hero

@slide {
  layout: blank
  background: image("BackgroundImage.png") overlay=rgba(0,0,0,0.35)
}

::textbox x=8% y=12% w=50% h=12% style=hero-title
Background image slide
""";
        var path = Path.Combine(tempDirectory, "background-slide.pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                BaseDirectory = tempDirectory,
                RenderMermaidDiagrams = false
            });

            using (var presentation = PowerPointPresentation.Open(path)) {
                var slide = Assert.Single(presentation.Slides);
                Assert.Contains(slide.Shapes.OfType<PowerPointAutoShape>(), shape => shape.Name.Contains("Background Overlay", StringComparison.Ordinal));
                Assert.DoesNotContain(slide.Shapes.OfType<PowerPointAutoShape>(), shape => shape.Name.Contains("Canvas Rail", StringComparison.Ordinal));
            }

            using (var package = PresentationDocument.Open(path, false)) {
                var slidePart = Assert.Single(package.PresentationPart!.SlideParts);
                var blipFill = slidePart.Slide.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();
                Assert.NotNull(blipFill);
            }
        } finally {
            if (Directory.Exists(tempDirectory)) {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_ResolvesRelativeSlideImagesFromBaseDirectory() {
        var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var imageSource = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
        var localImage = Path.Combine(tempDirectory, "EvotecLogo.png");
        File.Copy(imageSource, localImage, overwrite: true);

        var markup = """
---
profile: presentation
---

# Visual

@slide {
  layout: blank
}

![Logo](EvotecLogo.png)
""";
        var path = Path.Combine(tempDirectory, "relative-image-slide.pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                BaseDirectory = tempDirectory,
                RenderMermaidDiagrams = false
            });

            using var package = PresentationDocument.Open(path, false);
            var slidePart = Assert.Single(package.PresentationPart!.SlideParts);
            Assert.NotEmpty(slidePart.ImageParts);
            Assert.DoesNotContain("Image: EvotecLogo.png", slidePart.Slide.OuterXml, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(tempDirectory)) {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_PreservesJpegAspectRatioForContainedImages() {
        var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var imageSource = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");
        var localImage = Path.Combine(tempDirectory, "Kulek.jpg");
        File.Copy(imageSource, localImage, overwrite: true);

        var markup = """
---
profile: presentation
---

# Visual

@slide {
  layout: blank
}

::image src=Kulek.jpg x=10% y=20% w=30% h=20% fit=contain
""";
        var path = Path.Combine(tempDirectory, "contained-jpeg-slide.pptx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                BaseDirectory = tempDirectory,
                RenderMermaidDiagrams = false
            });

            using var presentation = PowerPointPresentation.Open(path);
            var slide = Assert.Single(presentation.Slides);
            var picture = Assert.Single(slide.Pictures);
            Assert.True(picture.WidthInches < 3.0, "Contained JPEG should not stretch to the full placement width.");
            Assert.InRange(picture.HeightInches, 1.12, 1.13);
        } finally {
            if (Directory.Exists(tempDirectory)) {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_DoesNotThrowForUrlImageSources() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        var markup = """
---
profile: presentation
---

# Demo

@slide {
  layout: blank
}

::image src=https://example.com/logo.png
""";

        try {
            var result = OfficeMarkupParser.Parse(markup);

            Assert.False(result.HasErrors);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                IncludeUnsupportedBlocksAsText = true,
                RenderMermaidDiagrams = false
            });

            using var presentation = PresentationDocument.Open(path, false);
            var slidePart = Assert.Single(presentation.PresentationPart!.SlideParts);
            Assert.Contains("Image: https://example.com/logo.png", slidePart.Slide.OuterXml, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_HonorsCustomSlideSizeOptionsForPercentPlacement() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        var markup = """
---
profile: presentation
---

# Custom Size

@slide {
  layout: blank
}

::textbox x=80% y=10% w=15% h=10%
Scaled textbox
""";

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                SlideWidthInches = 13.333,
                SlideHeightInches = 7.5,
                RenderMermaidDiagrams = false
            });

            using var presentation = PowerPointPresentation.Open(path);
            Assert.Equal(13.333, presentation.SlideSize.WidthInches, 3);
            Assert.Equal(7.5, presentation.SlideSize.HeightInches, 3);

            var slide = Assert.Single(presentation.Slides);
            var textBox = Assert.Single(slide.Shapes.OfType<PowerPointTextBox>(), box => box.Text.Contains("Scaled textbox", StringComparison.Ordinal));
            Assert.Equal(13.333 * 0.80, textBox.LeftInches, 3);
            Assert.Equal(7.5 * 0.10, textBox.TopInches, 3);
            Assert.Equal(13.333 * 0.15, textBox.WidthInches, 3);
            Assert.Equal(7.5 * 0.10, textBox.HeightInches, 3);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_AppliesNativeGradientBackgroundFromSemanticColors() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        var markup = """
---
profile: presentation
theme: evotec-modern
---

# Hero

@slide {
  layout: blank
  background: gradient(accent1, #FFFFFF)
}

::textbox x=8% y=12% w=50% h=12% style=hero-title
Gradient background slide
""";

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using (var package = PresentationDocument.Open(path, false)) {
                var slidePart = Assert.Single(package.PresentationPart!.SlideParts);
                var properties = slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;
                Assert.Null(properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>());

                var gradient = Assert.IsType<DocumentFormat.OpenXml.Drawing.GradientFill>(properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.GradientFill>());
                var stops = gradient.GetFirstChild<DocumentFormat.OpenXml.Drawing.GradientStopList>()!
                    .Elements<DocumentFormat.OpenXml.Drawing.GradientStop>()
                    .ToArray();
                Assert.Equal(2, stops.Length);
                Assert.Equal("2563EB", stops[0].GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?.Val?.Value);
                Assert.Equal("FFFFFF", stops[1].GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?.Val?.Value);
                Assert.Equal(8100000, gradient.GetFirstChild<DocumentFormat.OpenXml.Drawing.LinearGradientFill>()?.Angle?.Value);
            }

            using (var presentation = PowerPointPresentation.Open(path)) {
                var slide = Assert.Single(presentation.Slides);
                Assert.DoesNotContain(slide.Shapes.OfType<PowerPointAutoShape>(), shape => shape.Name.Contains("Canvas Rail", StringComparison.Ordinal));
            }
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_AppliesNativeGradientBackgroundAngleFromDirective() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        var markup = """
---
profile: presentation
theme: evotec-modern
---

# Hero

@slide {
  layout: blank
  background: gradient(primary, accent1) angle=45
}

::textbox x=8% y=12% w=50% h=12% style=hero-title
Gradient angle slide
""";

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using (var package = PresentationDocument.Open(path, false)) {
                var slidePart = Assert.Single(package.PresentationPart!.SlideParts);
                var properties = slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;
                var gradient = Assert.IsType<DocumentFormat.OpenXml.Drawing.GradientFill>(properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.GradientFill>());
                Assert.Equal(2700000, gradient.GetFirstChild<DocumentFormat.OpenXml.Drawing.LinearGradientFill>()?.Angle?.Value);
            }

            using (var presentation = PowerPointPresentation.Open(path)) {
                var slide = Assert.Single(presentation.Slides);
                Assert.DoesNotContain(slide.Shapes.OfType<PowerPointAutoShape>(), shape => shape.Name.Contains("Canvas Rail", StringComparison.Ordinal));
            }
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_MapsDirectionalTransitionsToNativePowerPointTransitions() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        var markup = """
---
profile: presentation
---

# Fade

@slide {
  transition: fade duration=0.6
}

---

# Push left

@slide {
  transition: push direction=left duration=0.5
}

---

# Warp out

@slide {
  transition: warp direction=out duration=0.5
}

---

# Ferris right

@slide {
  transition: ferris direction=right
}

---

# Blinds vertical

@slide {
  transition: blinds direction=vertical
}

---

# Comb horizontal

@slide {
  transition: comb direction=horizontal
}
""";

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                RenderMermaidDiagrams = false
            });

            using (var presentation = PowerPointPresentation.Open(path)) {
                var slides = presentation.Slides.ToArray();
                Assert.Equal(6, slides.Length);
                Assert.Equal(SlideTransition.Fade, slides[0].Transition);
                Assert.Equal(SlideTransition.PushLeft, slides[1].Transition);
                Assert.Equal(SlideTransition.WarpOut, slides[2].Transition);
                Assert.Equal(SlideTransition.FerrisRight, slides[3].Transition);
                Assert.Equal(SlideTransition.BlindsVertical, slides[4].Transition);
                Assert.Equal(SlideTransition.CombHorizontal, slides[5].Transition);
            }
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_RendersMermaidAsImageWhenRendererAvailable() {
        // The fake renderer is a Windows cmd shim because the real renderer is optional
        // in developer and CI environments.
        if (Path.DirectorySeparatorChar != '\\') {
            return;
        }

        var directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markup.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "diagram.pptx");
        Directory.CreateDirectory(directory);

        try {
            var rendererPath = CreateFakeMermaidRenderer(directory);
            var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Architecture

@slide {
  layout: blank
}

::mermaid x=8% y=20% w=70% h=40% fit=contain
flowchart LR
  Markup --> AST
  AST --> PPTX
""");

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                MermaidRendererPath = rendererPath,
                TemporaryDirectory = directory
            });

            using var package = PresentationDocument.Open(path, false);
            var slideParts = package.PresentationPart!.SlideParts.ToList();
            Assert.Contains(slideParts.SelectMany(part => part.ImageParts), part =>
                string.Equals(part.ContentType, "image/png", StringComparison.OrdinalIgnoreCase));
            var slideXml = string.Join(Environment.NewLine, slideParts.Select(part => part.Slide.OuterXml));
            Assert.Contains("OfficeIMO Markup Diagram Panel", slideXml, StringComparison.Ordinal);
            Assert.DoesNotContain("Markup --> AST", slideXml, StringComparison.Ordinal);
            Assert.DoesNotContain("Mermaid renderer", slideXml, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PowerPointExporter_RendersMermaidWhenRendererWritesLargeRedirectedOutput() {
        if (Path.DirectorySeparatorChar != '\\') {
            return;
        }

        var directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markup.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "diagram-noisy.pptx");
        Directory.CreateDirectory(directory);

        try {
            var rendererPath = CreateNoisyFakeMermaidRenderer(directory);
            var result = OfficeMarkupParser.Parse("""
---
profile: presentation
---

# Architecture

@slide {
  layout: blank
}

::mermaid x=8% y=20% w=70% h=40% fit=contain
flowchart LR
  Markdown --> AST
  AST --> Office
""");

            new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                OutputPath = path,
                MermaidRendererPath = rendererPath,
                TemporaryDirectory = directory
            });

            using var package = PresentationDocument.Open(path, false);
            var slideParts = package.PresentationPart!.SlideParts.ToList();
            Assert.Contains(slideParts.SelectMany(part => part.ImageParts), part =>
                string.Equals(part.ContentType, "image/png", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExcelExporter_CreatesOpenableWorkbookWithRangeTableFormulaAndChart() {
        var markup = """
---
profile: workbook
title: Revenue Workbook
---

@sheet {
  name: Revenue
}

# Revenue Workbook

::range address=A1
Product,2024,2025
A,100,120
B,80,92
C,60,77

::table name="RevenueTable" range=A1:C4 header=true

::formula cell=D2
=C2-B2

::format target=D2 numberFormat=0.00 bold=true

::chart type=column title="Revenue" source=A1:C4 cell=F2 width=480 height=320 category-title=Product value-title=Revenue value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            Assert.True(File.Exists(path));
            using (var spreadsheet = SpreadsheetDocument.Open(path, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.OfType<Sheet>().First(item => item.Name?.Value == "Revenue");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                Assert.Equal("Product", GetCellValue(spreadsheet, worksheetPart, "A1"));
                Assert.Equal("120", GetCellValue(spreadsheet, worksheetPart, "C2"));
                Assert.Equal("C2-B2", GetCell(worksheetPart, "D2")!.CellFormula!.Text);
                Assert.True(worksheetPart.TableDefinitionParts.Any());
                Assert.NotNull(worksheetPart.DrawingsPart);
                Assert.True(worksheetPart.DrawingsPart!.ChartParts.Any());
                var sheetView = worksheetPart.Worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().First();
                var pane = sheetView.GetFirstChild<Pane>()!;
                Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
                Assert.Equal(1D, pane.VerticalSplit!.Value);
                Assert.False(sheetView.ShowGridLines!.Value);
                var chartXml = worksheetPart.DrawingsPart!.ChartParts.First().ChartSpace.OuterXml;
                Assert.Contains("2563EB", chartXml, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Product", chartXml, StringComparison.Ordinal);
                Assert.Contains("Revenue", chartXml, StringComparison.Ordinal);
                Assert.Contains("#,##0", chartXml, StringComparison.Ordinal);
                Assert.Contains("outEnd", chartXml, StringComparison.Ordinal);
            }

            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            Assert.Empty(document.ValidateOpenXml());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExcelExporter_PlacesSheetQualifiedTableChartOnRequestedSheet() {
        var markup = """
---
profile: workbook
title: Dashboard Workbook
---

@sheet {
  name: Revenue
}

::range address=A1
Product,2024,2025
A,100,120
B,80,92
C,60,77

::table name="RevenueTable" range=A1:C4 header=true

@sheet {
  name: Dashboard
}

::chart type=column title="Revenue" source=Revenue!RevenueTable cell=B2 width=480 height=320

::formula cell=H1
=1+1
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            using var spreadsheet = SpreadsheetDocument.Open(path, false);
            var workbookPart = spreadsheet.WorkbookPart!;
            var sheets = workbookPart.Workbook.Sheets!.OfType<Sheet>().ToList();
            var revenueSheet = sheets.First(item => item.Name?.Value == "Revenue");
            var dashboardSheet = sheets.First(item => item.Name?.Value == "Dashboard");
            var revenuePart = (WorksheetPart)workbookPart.GetPartById(revenueSheet.Id!);
            var dashboardPart = (WorksheetPart)workbookPart.GetPartById(dashboardSheet.Id!);

            Assert.True(revenuePart.TableDefinitionParts.Any());
            Assert.Null(revenuePart.DrawingsPart);
            Assert.NotNull(dashboardPart.DrawingsPart);
            Assert.True(dashboardPart.DrawingsPart!.ChartParts.Any());
            Assert.Contains("2563EB", dashboardPart.DrawingsPart!.ChartParts.First().ChartSpace.OuterXml, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("1+1", GetCell(dashboardPart, "H1")!.CellFormula!.Text);
            Assert.False(dashboardPart.Worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().First().ShowGridLines!.Value);

            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            Assert.Empty(document.ValidateOpenXml());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExcelExporter_AcceptsSheetQualifiedWorkbookTargets() {
        var markup = """
---
profile: workbook
title: Qualified Workbook
---

::range address=Data!A1
Metric,Value,Double
Revenue,120,
Cost,80,

::formula cell=Data!C2
=B2*2

::formula cell=Data!C3
=B3*2

::format target=Data!B2:C3 numberFormat="#,##0" bold=true

::table name="DataTable" range=Data!A1:C3 header=true

::chart type=column title="Qualified Chart" source=Data!DataTable cell=Dashboard!B2 width=480 height=320
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            using var spreadsheet = SpreadsheetDocument.Open(path, false);
            var workbookPart = spreadsheet.WorkbookPart!;
            var sheets = workbookPart.Workbook.Sheets!.OfType<Sheet>().ToList();
            var dataSheet = sheets.First(item => item.Name?.Value == "Data");
            var dashboardSheet = sheets.First(item => item.Name?.Value == "Dashboard");
            var dataPart = (WorksheetPart)workbookPart.GetPartById(dataSheet.Id!);
            var dashboardPart = (WorksheetPart)workbookPart.GetPartById(dashboardSheet.Id!);

            Assert.Equal("Metric", GetCellValue(spreadsheet, dataPart, "A1"));
            Assert.Equal("120", GetCellValue(spreadsheet, dataPart, "B2"));
            Assert.Equal("B2*2", GetCell(dataPart, "C2")!.CellFormula!.Text);
            Assert.True(dataPart.TableDefinitionParts.Any());
            Assert.Null(dataPart.DrawingsPart);
            Assert.NotNull(dashboardPart.DrawingsPart);
            Assert.True(dashboardPart.DrawingsPart!.ChartParts.Any());
            Assert.False(dashboardPart.Worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().First().ShowGridLines!.Value);

            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            Assert.Empty(document.ValidateOpenXml());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExcelExporter_AppliesWorkbookFormattingColorFillAndBoldTogether() {
        var markup = """
---
profile: workbook
title: Formatting Workbook
---

::range address=Data!A1
Metric,Value
Revenue,120

::format target=Data!B2 numberFormat="#,##0" fill=#D9EAD3 color=#112233 bold=true italic=true underline=true align=center valign=middle border=thin border-color=#445566
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            var snapshot = document.CreateInspectionSnapshot();
            var dataSheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Data");
            var valueCell = Assert.Single(dataSheet.Cells, cell => cell.Row == 2 && cell.Column == 2);

            Assert.NotNull(valueCell.Style);
            Assert.Equal("#,##0", valueCell.Style!.NumberFormatCode);
            Assert.True(valueCell.Style.Bold);
            Assert.True(valueCell.Style.Italic);
            Assert.True(valueCell.Style.Underline);
            Assert.Equal("FF112233", valueCell.Style.FontColorArgb);
            Assert.Equal("FFD9EAD3", valueCell.Style.FillColorArgb);
            Assert.Equal("center", valueCell.Style.HorizontalAlignment);
            Assert.Equal("center", valueCell.Style.VerticalAlignment);
            Assert.NotNull(valueCell.Style.Border);
            Assert.Equal("thin", valueCell.Style.Border!.Left!.Style);
            Assert.Equal("FF445566", valueCell.Style.Border.Left.ColorArgb);
            Assert.Equal("thin", valueCell.Style.Border.Right!.Style);
            Assert.Equal("thin", valueCell.Style.Border.Top!.Style);
            Assert.Equal("thin", valueCell.Style.Border.Bottom!.Style);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExcelExporter_DoesNotForceBoldWhenFormattingBlockOmitsBoldAttribute() {
        var markup = """
---
profile: workbook
title: Formatting Workbook
---

::range address=Data!A1
Metric,Value
Revenue,120

::format target=Data!B2 fill=#D9EAD3 color=#112233 align=right
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path
            });

            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            var snapshot = document.CreateInspectionSnapshot();
            var dataSheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Data");
            var valueCell = Assert.Single(dataSheet.Cells, cell => cell.Row == 2 && cell.Column == 2);

            Assert.NotNull(valueCell.Style);
            Assert.False(valueCell.Style!.Bold);
            Assert.Equal("FF112233", valueCell.Style.FontColorArgb);
            Assert.Equal("FFD9EAD3", valueCell.Style.FillColorArgb);
            Assert.Equal("right", valueCell.Style.HorizontalAlignment);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExcelExportOptions_EnableWorkbookRepairHardeningByDefault() {
        var options = new OfficeMarkupExcelExportOptions();

        Assert.True(options.SafePreflight);
        Assert.True(options.ValidateOpenXml);
        Assert.True(options.SafeRepairDefinedNames);
    }

    [Fact]
    public void ExcelExporter_AcceptsExplicitWorkbookRepairHardeningOptions() {
        var markup = """
---
profile: workbook
title: Hardened Workbook
---

::range address=Data!A1
Quarter,Revenue,Costs
Q1,120,85
Q2,180,94
Q3,260,132
Q4,320,150

::table name="RevenueTable" range=Data!A1:C5 header=true

::chart type=column title="Quarterly Revenue" source=Data!RevenueTable cell=Dashboard!B2 width=480 height=320 category-title=Quarter value-title=Amount value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                OutputPath = path,
                SafePreflight = true,
                ValidateOpenXml = true,
                SafeRepairDefinedNames = true
            });

            Assert.True(File.Exists(path));
            using var document = OfficeIMO.Excel.ExcelDocument.Load(path, readOnly: true);
            Assert.Empty(document.ValidateOpenXml());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void WordExporter_CreatesOpenableDocumentWithDocumentBlocksAndChart() {
        var markup = """
---
profile: document
title: Architecture Note
---

# Architecture Note

::toc min=1 max=3 title="Contents"

::header text="OfficeIMO Markup"

This document starts as normal Markdown.

## Overview

- Markdown remains readable
- Office-specific constructs are explicit

| Area | Status |
| --- | --- |
| Parser | Ready |
| Exporters | Growing |

::chart type=column title="Quarterly Revenue" width=480 height=320
Quarter,Revenue
Q1,120
Q2,180

::page-break

::footer text="Generated from OfficeIMO Markup"
""";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try {
            var result = OfficeMarkupParser.Parse(markup);

            new OfficeMarkupWordExporter().Export(result.Document, new OfficeMarkupWordExportOptions {
                OutputPath = path
            });

            Assert.True(File.Exists(path));
            using (var document = OfficeIMO.Word.WordDocument.Load(path, readOnly: true)) {
                Assert.NotNull(document.TableOfContent);
                Assert.True(document.PageBreaks.Count >= 1);
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("Architecture Note", StringComparison.Ordinal));
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("This document starts", StringComparison.Ordinal));
                Assert.NotEmpty(document.Tables);
                Assert.NotEmpty(document.Charts);
                Assert.Empty(document.ValidateDocument());
            }

            using (var package = WordprocessingDocument.Open(path, false)) {
                Assert.True(package.MainDocumentPart!.ChartParts.Any());
            }
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void WordExporter_IncrementsOrderedListNumbersInsideSections() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try {
            var documentModel = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
            var section = new OfficeMarkupSectionBlock("Appendix");
            var list = new OfficeMarkupListBlock(ordered: true, start: 1);
            list.Items.Add(new OfficeMarkupListItem("First"));
            list.Items.Add(new OfficeMarkupListItem("Second"));
            list.Items.Add(new OfficeMarkupListItem("Third"));
            section.Blocks.Add(list);
            documentModel.Blocks.Add(section);

            new OfficeMarkupWordExporter().Export(documentModel, new OfficeMarkupWordExportOptions {
                OutputPath = path
            });

            using var document = OfficeIMO.Word.WordDocument.Load(path, readOnly: true);
            Assert.Contains(document.Paragraphs, paragraph => string.Equals(paragraph.Text, "1. First", StringComparison.Ordinal));
            Assert.Contains(document.Paragraphs, paragraph => string.Equals(paragraph.Text, "2. Second", StringComparison.Ordinal));
            Assert.Contains(document.Paragraphs, paragraph => string.Equals(paragraph.Text, "3. Third", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    private static Cell? GetCell(WorksheetPart worksheetPart, string cellReference) {
        return worksheetPart.Worksheet!.Descendants<Cell>().FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
        var cell = GetCell(worksheetPart, cellReference);
        if (cell == null) {
            return string.Empty;
        }

        var text = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(text, out var sharedStringIndex)) {
            return document.WorkbookPart!.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
        }

        return text;
    }

    private static string CreateFakeMermaidRenderer(string directory) {
        var rendererPath = Path.Combine(directory, "fake-mmdc.cmd");
        File.WriteAllText(rendererPath, """
@echo off
set "out="
:args
if "%~1"=="" goto run
if /I "%~1"=="-o" goto output
shift
goto args
:output
shift
set "out=%~1"
shift
goto args
:run
if "%out%"=="" exit /b 2
powershell -NoProfile -ExecutionPolicy Bypass -Command "[IO.File]::WriteAllBytes($env:out,[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='))"
exit /b %ERRORLEVEL%
""");
        return rendererPath;
    }

    private static string CreateNoisyFakeMermaidRenderer(string directory) {
        var rendererPath = Path.Combine(directory, "fake-mmdc-noisy.cmd");
        File.WriteAllText(rendererPath, """
@echo off
set "out="
:args
if "%~1"=="" goto run
if /I "%~1"=="-o" goto output
shift
goto args
:output
shift
set "out=%~1"
shift
goto args
:run
if "%out%"=="" exit /b 2
for /L %%I in (1,1,8000) do <nul set /p ="x"
echo noisy stderr 1>&2
powershell -NoProfile -ExecutionPolicy Bypass -Command "[IO.File]::WriteAllBytes($env:out,[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='))"
exit /b %ERRORLEVEL%
""");
        return rendererPath;
    }
}
