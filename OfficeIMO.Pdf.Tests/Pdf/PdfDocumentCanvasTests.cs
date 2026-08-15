using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentCanvasTests {
    [Fact]
    public void CanvasFormDataExportOmitsNoExportFields() {
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas
                .TextField("enabled", "included", 20D, 20D, 120D, 20D)
                .TextField("disabled", "excluded", 20D, 50D, 120D, 20D, style: new PdfFormFieldStyle { IsNoExport = true }))
            .ToBytes();

        PdfFormDataSet data = PdfDocument.Open(bytes).Forms.ExportData();

        Assert.Contains(data.Fields, field => field.Name == "enabled" && field.Values.SequenceEqual(new[] { "included" }));
        Assert.DoesNotContain(data.Fields, field => field.Name == "disabled");
    }

    [Fact]
    public void CanvasFormFields_CreatePositionedInspectableAcroFormWidgets() {
        var textStyle = new PdfFormFieldStyle {
            IsRequired = true,
            AlternateName = "Contact name"
        };
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas
                .TextField("ContactName", "Ada", 20D, 30D, 160D, 24D, style: textStyle)
                .CheckBox("ContactAccept", true, 20D, 70D, 16D, 16D)
                .ChoiceField("ContactCountry", new[] { "Poland", "Germany" }, new[] { "Poland" }, 20D, 105D, 160D, 24D)
                .RadioButton("ContactMethod", "Email", false, 20D, 145D, 16D, 16D)
                .RadioButton("ContactMethod", "Phone", true, 80D, 145D, 16D, 16D))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal(4, info.FormFields.Count);
        PdfFormField text = Assert.Single(info.FormFields, field => field.Name == "ContactName");
        Assert.Equal("Ada", text.Value);
        Assert.True(text.IsRequired);
        Assert.Equal("Contact name", text.AlternateName);
        Assert.InRange(Assert.Single(text.Widgets).Width, 159.9D, 160.1D);

        PdfFormField checkBox = Assert.Single(info.FormFields, field => field.Name == "ContactAccept");
        Assert.True(checkBox.IsCheckBox);
        Assert.Equal("Yes", checkBox.Value);

        PdfFormField choice = Assert.Single(info.FormFields, field => field.Name == "ContactCountry");
        Assert.Equal(new[] { "Poland", "Germany" }, choice.Options.Select(option => option.DisplayText).ToArray());
        Assert.Equal("Poland", choice.Value);

        PdfFormField radio = Assert.Single(info.FormFields, field => field.Name == "ContactMethod");
        Assert.True(radio.IsRadioButton);
        Assert.Equal("Phone", radio.Value);
        Assert.Equal(2, radio.Widgets.Count);
        Assert.True(radio.Widgets[1].X1 > radio.Widgets[0].X1);
    }

    [Fact]
    public void CanvasChoiceField_ValueOnlyOverloadRejectsDuplicateExportValues() {
        var options = new[] {
            new PdfFormFieldOption("x", "First"),
            new PdfFormFieldOption("x", "Second")
        };

        ArgumentException exception = Assert.Throws<ArgumentException>(() => new PdfPageCanvas().ChoiceField(
            "Choice",
            options,
            new[] { "x" },
            20D,
            20D,
            120D,
            24D));

        Assert.Contains("export values must be unique", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasChoiceField_StringMultiSelectNormalizesDuplicateValuesBeforeSerialization() {
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas.ChoiceField(
                "Choice",
                new[] { "A", "B" },
                new[] { "A", "A" },
                20D,
                20D,
                120D,
                36D,
                isComboBox: false,
                allowsMultipleSelection: true))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal(new[] { 0 }, field.SelectedIndices);
        Assert.Equal(new[] { "A" }, field.SelectedOptions.Select(option => option.ExportValue).ToArray());
    }

    [Fact]
    public void CanvasChoiceField_RejectsMultiSelectComboBoxesWhenAdded() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentException>(() => canvas.ChoiceField(
            "Strings",
            new[] { "A", "B" },
            new[] { "A" },
            20D,
            20D,
            120D,
            36D,
            isComboBox: true,
            allowsMultipleSelection: true));
        Assert.Throws<ArgumentException>(() => canvas.ChoiceField(
            "Options",
            new[] { new PdfFormFieldOption("A", "First"), new PdfFormFieldOption("B", "Second") },
            new[] { "A" },
            20D,
            20D,
            120D,
            36D,
            isComboBox: true,
            allowsMultipleSelection: true));
    }

    [Fact]
    public void CanvasEditableComboFieldsAllowCustomScalarValues() {
        var style = new PdfFormFieldStyle { IsEditableChoice = true };
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas
                .ChoiceField("Strings", new[] { "A", "B" }, new[] { "Custom" }, 20D, 20D, 120D, 24D, style: style)
                .ChoiceField("Options", new[] {
                    new PdfFormFieldOption("A", "First"),
                    new PdfFormFieldOption("B", "Second")
                }, new[] { "Other" }, 20D, 60D, 120D, 24D, style: style))
            .ToBytes();

        PdfFormField[] fields = PdfInspector.Inspect(bytes).FormFields.ToArray();
        Assert.Equal(new[] { "Custom", "Other" }, fields.Select(field => field.Value));
        Assert.All(fields, field => Assert.True(field.IsEditableChoice));
    }

    [Fact]
    public void CanvasFormFields_RejectPeriodsInPartialFieldNames() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentException>(() => canvas.TextField("user.email", "Ada", 20D, 20D, 120D, 24D));
        Assert.Throws<ArgumentException>(() => canvas.CheckBox("user.accepted", true, 20D, 20D, 14D, 14D));
        Assert.Throws<ArgumentException>(() => canvas.ChoiceField("user.country", new[] { "Poland" }, new[] { "Poland" }, 20D, 20D, 120D, 24D));
        Assert.Throws<ArgumentException>(() => canvas.RadioButton("user.method", "Email", true, 20D, 20D, 14D, 14D));
    }

    [Fact]
    public void CanvasCheckBox_RejectsReservedAndNonAsciiAppearanceStatesWhenAdded() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentException>(() => canvas.CheckBox("Check", false, 20D, 20D, 14D, 14D, "Off"));
        Assert.Throws<ArgumentException>(() => canvas.CheckBoxWithExportValue("Check", false, 20D, 20D, 14D, 14D, "Off", "off-export"));
        Assert.Throws<ArgumentException>(() => canvas.CheckBox("Check", false, 20D, 20D, 14D, 14D, "Y\u2713"));
        Assert.Throws<ArgumentException>(() => canvas.CheckBoxWithExportValue("Check", false, 20D, 20D, 14D, 14D, "Y\u2713", "accepted"));
    }

    [Fact]
    public void CanvasRadioButtons_CanStartWithNoSelectedWidget() {
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas
                .RadioButton("Preference", "One", false, 20D, 20D, 14D, 14D)
                .RadioButton("Preference", "Two", false, 50D, 20D, 14D, 14D))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.True(field.IsRadioButton);
        Assert.Equal("Off", field.Value);
        Assert.All(field.Widgets, widget => Assert.Equal("Off", widget.AppearanceState));
    }

    [Fact]
    public void CanvasRadioButtons_CanShareOneFieldAcrossPages() {
        byte[] bytes = PdfDocument.Create()
            .Page(page => page.Canvas(canvas => canvas.RadioButton("AcrossPages", "First", false, 20D, 20D, 14D, 14D)))
            .Page(page => page.Canvas(canvas => canvas.RadioButton("AcrossPages", "Second", true, 20D, 20D, 14D, 14D)))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("AcrossPages", field.Name);
        Assert.Equal("Second", field.Value);
        Assert.Equal(new[] { 1, 2 }, field.PageNumbers);
        Assert.Equal(2, field.Widgets.Count);
    }

    [Fact]
    public void CanvasRadioButtons_NormalizeMultipleCrossPageSelectionsToTheLastOption() {
        byte[] bytes = PdfDocument.Create()
            .Page(page => page.Canvas(canvas => canvas.RadioButton("AcrossPages", "First", true, 20D, 20D, 14D, 14D)))
            .Page(page => page.Canvas(canvas => canvas.RadioButton("AcrossPages", "Second", true, 20D, 20D, 14D, 14D)))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);

        Assert.Equal("Second", field.Value);
        Assert.Equal(new[] { "Off", "Second" }, field.Widgets.Select(widget => widget.AppearanceState).ToArray());
    }

    [Fact]
    public void CanvasRadioButtons_CanUseDifferentWidgetDimensionsWithinOneField() {
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas
                .RadioButton("VariableSize", "Small", false, 20D, 20D, 12D, 12D)
                .RadioButton("VariableSize", "Large", true, 60D, 20D, 24D, 18D))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("Large", field.Value);
        Assert.Equal(12D, field.Widgets[0].Width, 3);
        Assert.Equal(24D, field.Widgets[1].Width, 3);
        Assert.Equal(18D, field.Widgets[1].Height, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CanvasRadioButtons_RejectMixedFieldLevelStateWithinOneGroup(bool disabledFirst) {
        var enabled = new PdfFormFieldStyle();
        var disabled = new PdfFormFieldStyle { IsReadOnly = true, IsNoExport = true };
        PdfFormFieldStyle first = disabledFirst ? disabled : enabled;
        PdfFormFieldStyle second = disabledFirst ? enabled : disabled;
        PdfDocument document = PdfDocument.Create()
            .Canvas(canvas => canvas
                .RadioButton("MixedState", "First", false, 20D, 20D, 14D, 14D, first)
                .RadioButton("MixedState", "Second", true, 50D, 20D, 14D, 14D, second));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());

        Assert.Contains("consistent read-only and no-export", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasRadioButtons_RejectMixedFieldLevelStateAcrossPages() {
        PdfDocument document = PdfDocument.Create()
            .Page(page => page.Canvas(canvas => canvas.RadioButton("MixedPages", "First", false, 20D, 20D, 14D, 14D)))
            .Page(page => page.Canvas(canvas => canvas.RadioButton(
                "MixedPages",
                "Second",
                true,
                20D,
                20D,
                14D,
                14D,
                new PdfFormFieldStyle { IsReadOnly = true })));

        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void CanvasRadioButtons_CanSeparateAppearanceStatesFromUnicodeExportValues() {
        byte[] bytes = PdfDocument.Create()
            .Page(page => page.Canvas(canvas => canvas.RadioButtonWithExportValue("Preference", "Option1", "caf\u00E9", false, 20D, 20D, 14D, 14D)))
            .Page(page => page.Canvas(canvas => canvas.RadioButtonWithExportValue("Preference", "Option2", "th\u00E9", true, 50D, 20D, 14D, 14D)))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);

        Assert.Equal("Option2", field.Value);
        Assert.Equal(new[] { "caf\u00E9", "th\u00E9" }, field.Options.Select(option => option.ExportValue).ToArray());
        Assert.Equal("th\u00E9", Assert.Single(PdfDocument.Open(bytes).Forms.ExportData().Fields).Values[0]);
    }

    [Fact]
    public void CanvasActualText_PreservesLogicalExtractionForReversePositionedFragments() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.ActualText("ABC", logical => logical
                .Text("A", 50D, 10D, 10D, 20D)
                .Text("B", 35D, 10D, 10D, 20D)
                .Text("C", 20D, 10D, 10D, 20D)))
            .ToBytes();

        Assert.Contains("ABC", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("/ActualText", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasActualText_ReplacesTextInsideEffectGroupsWithoutDuplication() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.ActualText("AB", logical => logical
                .Effect(OfficeIMO.Drawing.OfficeTransform.Identity, .5D, effect => effect.Text("A", 20D, 10D, 10D, 20D))
                .Text("B", 35D, 10D, 10D, 20D)))
            .ToBytes();

        Assert.Equal("AB", string.Concat(PdfReadDocument.Open(bytes).ExtractText().Where(character => !char.IsWhiteSpace(character))));
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "Span");
    }

    [Fact]
    public void CanvasActualText_RejectsInvalidArgumentsAndEmptyBuilders() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentNullException>(() => canvas.ActualText(null!, _ => { }));
        Assert.Throws<ArgumentException>(() => canvas.ActualText(string.Empty, _ => { }));
        Assert.Throws<ArgumentNullException>(() => canvas.ActualText("Text", null!));
        Assert.Throws<ArgumentException>(() => canvas.ActualText("Text", _ => { }));
    }

    [Fact]
    public void CanvasStructure_GroupsFragmentedHeadingAndParagraphTextUnderSection() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Section, section => section
                    .Structure(PdfCanvasStructureRole.Heading1, heading => heading
                        .Text(new[] { PdfTextRun.Normal("Heading") }, PdfCanvasTextStructureRole.Span, 10D, 10D, 120D, 20D))
                    .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph
                        .Text(new[] { PdfTextRun.Normal("Paragraph") }, PdfCanvasTextStructureRole.Span, 10D, 40D, 120D, 20D))))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, element => element.StructureType == "Sect");
        PdfStructureElementInfo heading = Assert.Single(tagged.StructureElements, element => element.StructureType == "H1");
        PdfStructureElementInfo paragraph = Assert.Single(tagged.StructureElements, element => element.StructureType == "P");
        Assert.Contains(heading.ObjectNumber, section.ChildElementObjectNumbers);
        Assert.Contains(paragraph.ObjectNumber, section.ChildElementObjectNumbers);
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "Span"));
    }

    [Fact]
    public void CanvasStructure_KeepsInteractiveFormWidgetsUnderTheActiveSemanticParent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Section, section => section
                .TextField("ContactName", "Ada", 10D, 10D, 100D, 20D)))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, element => element.StructureType == "Sect");
        PdfStructureElementInfo form = Assert.Single(tagged.StructureElements, element => element.StructureType == "Form");

        Assert.Contains(form.ObjectNumber, section.ChildElementObjectNumbers);
        Assert.Equal(1, form.ObjectReferenceCount);
    }

    [Fact]
    public void CanvasStructure_BuildsNestedListAndTableHierarchyWithCellAttributes() {
        var headerOptions = new PdfCanvasStructureOptions {
            HeaderScope = PdfCanvasTableHeaderScope.Column,
            ColumnSpan = 2
        };
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.List, list => list
                    .Structure(PdfCanvasStructureRole.ListItem, item => item
                        .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("1.", 10D, 10D, 20D, 20D))
                        .Structure(PdfCanvasStructureRole.ListBody, body => body.Text("First item", 35D, 10D, 100D, 20D))))
                .Structure(PdfCanvasStructureRole.Table, table => table
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Header", 10D, 40D, 100D, 20D), headerOptions))))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        PdfStructureElementInfo list = Assert.Single(tagged.StructureElements, element => element.StructureType == "L");
        PdfStructureElementInfo listItem = Assert.Single(tagged.StructureElements, element => element.StructureType == "LI");
        PdfStructureElementInfo label = Assert.Single(tagged.StructureElements, element => element.StructureType == "Lbl");
        PdfStructureElementInfo body = Assert.Single(tagged.StructureElements, element => element.StructureType == "LBody");
        Assert.Contains(listItem.ObjectNumber, list.ChildElementObjectNumbers);
        Assert.Contains(label.ObjectNumber, listItem.ChildElementObjectNumbers);
        Assert.Contains(body.ObjectNumber, listItem.ChildElementObjectNumbers);

        PdfStructureElementInfo table = Assert.Single(tagged.StructureElements, element => element.StructureType == "Table");
        PdfStructureElementInfo row = Assert.Single(tagged.StructureElements, element => element.StructureType == "TR");
        PdfStructureElementInfo header = Assert.Single(tagged.StructureElements, element => element.StructureType == "TH");
        Assert.Contains(row.ObjectNumber, table.ChildElementObjectNumbers);
        Assert.Contains(header.ObjectNumber, row.ChildElementObjectNumbers);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Scope /Column", raw, StringComparison.Ordinal);
        Assert.Contains("/ColSpan 2", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasStructure_RejectsInvalidRolesOptionsAndEmptyBuilders() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentOutOfRangeException>(() => canvas.Structure((PdfCanvasStructureRole)99, _ => { }));
        Assert.Throws<ArgumentNullException>(() => canvas.Structure(PdfCanvasStructureRole.List, null!));
        Assert.Throws<ArgumentException>(() => canvas.Structure(PdfCanvasStructureRole.List, _ => { }));
        Assert.Throws<ArgumentException>(() => canvas.Structure(
            PdfCanvasStructureRole.List,
            nested => nested.Text("Item", 0D, 0D, 20D, 20D),
            new PdfCanvasStructureOptions { ColumnSpan = 2 }));
        Assert.Throws<ArgumentException>(() => canvas.Structure(
            PdfCanvasStructureRole.TableCell,
            nested => nested.Text("Cell", 0D, 0D, 20D, 20D),
            new PdfCanvasStructureOptions { HeaderScope = PdfCanvasTableHeaderScope.Row }));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfCanvasStructureOptions { ColumnSpan = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfCanvasStructureOptions { HeaderScope = (PdfCanvasTableHeaderScope)99 });
    }

    [Fact]
    public void CanvasArtifact_ExcludesDecorativeContentFromTheStructureTree() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas
                .Artifact(artifact => artifact
                    .Text("Decoration", 10D, 10D, 100D, 20D)
                    .SearchableText("Searchable decoration", 10D, 30D)
                    .Table(new[] { new[] { "Decorative header" }, new[] { "Decorative cell" } }, 10D, 40D, 100D, 40D)
                    .Figure("Decorative figure", figure => figure.Text("Figure decoration", 10D, 85D, 100D, 20D)))
                .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Meaningful", 10D, 40D, 100D, 20D)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        Assert.Contains("/Artifact BMC", raw, StringComparison.Ordinal);
        Assert.Equal(2, tagged.StructureElements.Count(element => element.StructureType == "P"));
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType is "Span" or "TH" or "TD" or "Figure");
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.DoesNotContain("Decoration", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("Searchable decoration", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("Decorative header", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("Decorative cell", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("Figure decoration", extracted, StringComparison.Ordinal);

        var canvas = new PdfPageCanvas();
        Assert.Throws<ArgumentNullException>(() => canvas.Artifact(null!));
        Assert.Throws<ArgumentException>(() => canvas.Artifact(_ => { }));
    }

    [Fact]
    public void CanvasArtifact_SuppressesEveryInteractiveDescendant() {
        const string uri = "https://evotec.xyz/decorative";
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Artifact(artifact => artifact
                .Text(new[] { PdfTextRun.Link("Decorative link", uri) }, 10D, 10D, 100D, 20D)
                .TextAnnotation("Decorative note", 10D, 35D)
                .FreeTextAnnotation("Decorative free text", 35D, 35D, 100D, 20D)
                .HighlightAnnotation("Decorative highlight", 10D, 60D, 100D, 20D)
                .Outline("Decorative outline", 1, 10D)
                .Structure(PdfCanvasStructureRole.Paragraph, nested => nested
                    .TextField("DecorativeField", "Value", 10D, 85D, 100D, 20D))))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.FormFields);
        Assert.Empty(info.GetAnnotationsBySubtype("Text"));
        Assert.Empty(info.GetAnnotationsBySubtype("FreeText"));
        Assert.Empty(info.GetAnnotationsBySubtype("Highlight"));
        Assert.Empty(info.Outlines);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "Form");
        Assert.Contains("/Artifact BMC", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasOutline_SupportsPerEntryExpansionState() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { OutlineExpansionLevel = 0 })
            .Canvas(canvas => canvas
                .Outline("Open parent", 1, 10D, PdfOutlineState.Open)
                .Outline("Open child", 2, 20D)
                .Outline("Closed parent", 1, 30D, PdfOutlineState.Closed)
                .Outline("Closed child", 2, 40D))
            .ToBytes();

        IReadOnlyList<PdfOutlineItem> outlines = PdfReadDocument.Open(bytes).Outlines;
        Assert.True(outlines[0].IsExpanded);
        Assert.False(outlines[1].IsExpanded);
        Assert.Equal("Open child", Assert.Single(outlines[0].Children).Title);
        Assert.Equal("Closed child", Assert.Single(outlines[1].Children).Title);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageCanvas().Outline("Invalid", 1, 1D, (PdfOutlineState)99));
    }

    [Fact]
    public void CanvasFigure_GroupsMixedCanvasContentUnderOneTaggedFigure() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Figure("Composite diagram", figure => figure
                .Shape(OfficeShape.Rectangle(20D, 10D), 12D, 12D)
                .Text("Diagram label", 12D, 28D, 100D, 20D)
                .Image(CreateMinimalRgbPng(), 120D, 12D, 20D, 20D, alternativeText: "Nested image alt")))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        PdfStructureElementInfo figure = Assert.Single(tagged.StructureElements, element => element.StructureType == "Figure");
        Assert.Equal("Composite diagram", figure.AlternateText);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "P");
        Assert.Equal(1, CountOccurrences(Encoding.ASCII.GetString(bytes), "/Figure <<"));
    }

    [Fact]
    public void CanvasFigure_RejectsMissingAlternativeTextOrBuilder() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentException>(() => canvas.Figure(" ", _ => { }));
        Assert.Throws<ArgumentNullException>(() => canvas.Figure("Figure", null!));
        Assert.Throws<ArgumentException>(() => canvas.Figure("Figure", _ => { }));
    }

    [Fact]
    public void CanvasDrawing_UsesOneOuterFigureForNestedImageAccessibility() {
        var drawing = new OfficeDrawing(20D, 20D)
            .AddImage(
                CreateMinimalRgbPng(),
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 20D, 20D)),
                "Nested image alternative text");

        foreach (double size in new[] { 20D, 40D }) {
            byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
                .TaggedPdfCatalogMarkers()
                .Canvas(canvas => canvas.Drawing(
                    drawing,
                    10D,
                    10D,
                    size,
                    size,
                    new PdfDrawingStyle { AlternativeText = "Outer drawing alternative text" }))
                .ToBytes();

            PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
            PdfStructureElementInfo figure = Assert.Single(tagged.StructureElements, element => element.StructureType == "Figure");
            Assert.Equal("Outer drawing alternative text", figure.AlternateText);
            Assert.Equal(1, CountOccurrences(Encoding.ASCII.GetString(bytes), "/Figure <<"));

            byte[] decorative = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
                .Canvas(canvas => canvas.Drawing(
                    drawing,
                    10D,
                    10D,
                    size,
                    size,
                    new PdfDrawingStyle { Decorative = true }))
                .ToBytes();
            Assert.DoesNotContain("/Figure << /Alt", Encoding.ASCII.GetString(decorative), StringComparison.Ordinal);
        }
    }

    [Fact]
    public void CanvasStructure_RetainedImageKeepsFigureUnderDeclaredParent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Section, section => section
                .Image(CreateMinimalRgbPng(), 10D, 10D, 20D, 20D, alternativeText: "Nested image alternative text")))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, element => element.StructureType == "Sect");
        PdfStructureElementInfo figure = Assert.Single(tagged.StructureElements, element => element.StructureType == "Figure");
        Assert.Equal("Nested image alternative text", figure.AlternateText);
        Assert.Contains(figure.ObjectNumber, section.ChildElementObjectNumbers);
    }

    [Fact]
    public void CanvasClip_DropsStructureForAChildImageOutsideTheClip() {
        OfficeDrawing drawing = new OfficeDrawing(20D, 20D)
            .AddImage(
                CreateMinimalRgbPng(),
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 20D, 20D)),
                "Clipped image alternative text");

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 120D,
                PageHeight = 120D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Clip(60D, 10D, 20D, 20D, clipped => clipped
                .Image(CreateMinimalRgbPng(), 10D, 40D, 20D, 20D, alternativeText: "Direct clipped image alternative text")
                .Drawing(drawing, 10D, 10D, 20D, 20D)))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "Figure");
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("/Figure << /Alt", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Image", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasText_RendersAtFixedTopLeftCoordinatesWithoutMovingFlowContent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 72,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Text("CanvasTitle", 24, 20, 120, 24, fontSize: 12, color: PdfColor.FromRgb(20, 90, 160)))
            .Paragraph(paragraph => paragraph.Text("FlowAfterCanvas"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double canvasY = FindWordStartY(page, "CanvasTitle");
        double flowY = FindWordStartY(page, "FlowAfterCanvas");

        Assert.InRange(FindWordStartX(page, "CanvasTitle"), 23D, 26D);
        Assert.True(canvasY > flowY, "Canvas text should render above the flow paragraph when placed near the page top.");
        Assert.InRange(flowY, 77D, 91D);
    }

    [Fact]
    public void CanvasText_EmitsTypedHeadingStructureWhenTagged() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Text(
                new[] { PdfTextRun.Normal("Canvas semantic heading") },
                PdfCanvasTextStructureRole.Heading2,
                24,
                20,
                180,
                24,
                fontSize: 12))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        Assert.Contains("Document", tagged.StructureTypes);
        Assert.Contains("H2", tagged.StructureTypes);
        Assert.True(tagged.MarkedContentReferenceCount >= 1);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageCanvas().Text(
            new[] { PdfTextRun.Normal("Invalid") },
            (PdfCanvasTextStructureRole)99,
            0,
            0,
            10,
            10));
    }

    [Fact]
    public void CanvasShape_RendersRectangleAtFixedTopLeftCoordinates() {
        var shape = OfficeShape.Rectangle(60, 20);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1.25D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Shape(shape, 30, 40))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("30 100 60 20 re", content, StringComparison.Ordinal);
        Assert.Contains("1.25 w", content, StringComparison.Ordinal);
        Assert.Contains(" B", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasShape_WithRotation_RendersUsingSharedShapeTransform() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Shape(shape, 30, 40, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0 -1 -1 0 60 130 cm", content, StringComparison.Ordinal);
        Assert.Contains("0 0 40 20 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void Shape_AllowsHorizontalAndVerticalLineBounds() {
        OfficeShape horizontalLine = OfficeShape.Line(0, 0, 80, 0);
        horizontalLine.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        horizontalLine.StrokeWidth = 2D;

        OfficeShape verticalLine = OfficeShape.Line(0, 0, 0, 50);
        verticalLine.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        verticalLine.StrokeWidth = 2D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Shape(horizontalLine)
            .Shape(verticalLine)
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void CanvasDrawing_RendersSharedVectorSceneInsideFixedFrame() {
        var drawing = new OfficeDrawing(50, 20);
        var shape = OfficeShape.Rectangle(20, 10);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;
        drawing.AddShape(shape, 5, 5);
        drawing.AddText(
            "SceneText",
            8,
            4,
            36,
            10,
            new OfficeFontInfo("Aptos", 6D, OfficeFontStyle.Bold),
            PdfColor.FromRgb(31, 78, 121).ToOfficeColor(),
            OfficeTextAlignment.Center);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 140,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Drawing(drawing, 20, 30, 100, 40))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Group << /S /Transparency /I true /K false >>", content, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        string text = string.Join("", letters.Select(letter => letter.Value));
        Assert.Contains("SceneText", text, StringComparison.Ordinal);
        Assert.All(letters, letter => Assert.InRange(letter.StartBaseLine.X, 20D, 120D));
    }

    [Fact]
    public void CanvasDrawing_DownscalesScenesLargerThanThePageWithoutClippingTheFrame() {
        var drawing = new OfficeDrawing(400D, 160D);
        OfficeShape shape = OfficeShape.Rectangle(400D, 160D);
        shape.FillColor = OfficeColor.Red;
        shape.StrokeWidth = 0D;
        drawing.AddShape(shape, 0D, 0D);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240D,
                PageHeight = 140D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Drawing(drawing, 20D, 30D, 100D, 40D))
            .ToBytes();

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(110, 50));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(125, 50));
    }

    [Fact]
    public void CanvasClip_KeepsImagesWhoseTransformedDrawingBoundsIntersectTheClip() {
        var drawing = new OfficeDrawing(20D, 20D)
            .AddImage(
                CreateMinimalRgbPng(),
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 20D, 20D)));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 120D,
                PageHeight = 120D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(60D, 10D, 20D, 80D, clipped =>
                clipped.Drawing(drawing, 10D, 10D, 80D, 80D)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("4 0 0 4", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasClip_KeepsEffectImagesWhenOnlyTheRotatedFootprintIntersects() {
        var source = new OfficeDrawing(60D, 50D)
            .AddImage(
                CreateMinimalRgbPng(),
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(10D, 20D, 40D, 10D), rotationDegrees: 45D));
        var drawing = new OfficeDrawing(100D, 50D)
            .AddEffectDrawing(source, OfficeTransform.Translate(40D, 0D));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 120D,
                PageHeight = 120D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(0D, 18D, 120D, 2D, clipped =>
                clipped.Drawing(drawing, 10D, 10D, 100D, 50D)))
            .ToBytes();

        Assert.Contains("/Im1 Do", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Drawing_AllowsHorizontalAndVerticalLineBounds() {
        var drawing = new OfficeDrawing(100, 70)
            .AddShape(OfficeShape.Line(0, 0, 80, 0), 10, 10)
            .AddShape(OfficeShape.Line(0, 0, 0, 50), 20, 10);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Drawing(drawing)
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void FlowDrawing_RendersSharedVectorSceneText() {
        var drawing = new OfficeDrawing(120, 36)
            .AddText(
                "FlowSceneText",
                8,
                8,
                104,
                16,
                new OfficeFontInfo("Aptos", 10D, OfficeFontStyle.Bold),
                PdfColor.FromRgb(31, 78, 121).ToOfficeColor(),
                OfficeTextAlignment.Center);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("FlowSceneText", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_RendersImageAtFixedTopLeftCoordinatesWithAltText() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, alternativeText: "Canvas logo"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("60 0 0 30 30 90 cm", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <43616E766173206C6F676F> >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_WithSourceCrop_ClipsAndOffsetsImageInsideDeclaredFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 40, 50, 60, 30, new PdfImageStyle {
                SourceCrop = new PdfImageSourceCrop(left: 0.5D, top: 0D, right: 0D, bottom: 0D)
            }, linkUri: "https://evotec.xyz/cropped"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("120 0 0 30 -20 80 cm", content, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0.5 1 re", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
        AssertClose(40D, link.X1);
        AssertClose(80D, link.Y1);
        AssertClose(100D, link.X2);
        AssertClose(110D, link.Y2);
    }

    [Fact]
    public void CanvasImage_WithRotation_RendersImageAroundDeclaredFrameCenter() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0 60 -30 0 75 75 cm", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_RendersBeforeFollowingShapeInCanvasOrder() {
        var shape = OfficeShape.Rectangle(70, 35);
        shape.FillColor = PdfColor.FromRgb(255, 255, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .Image(CreateMinimalRgbPng(), 30, 40, 60, 30)
                .Shape(shape, 25, 35))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int imageDraw = content.IndexOf("/Im1 Do", StringComparison.Ordinal);
        int shapeDraw = content.IndexOf("25 90 70 35 re", StringComparison.Ordinal);

        Assert.True(imageDraw >= 0, "Expected the canvas image draw command to be present.");
        Assert.True(shapeDraw >= 0, "Expected the following canvas shape draw command to be present.");
        Assert.True(imageDraw < shapeDraw, "Canvas images should be painted in declared order instead of being appended after later canvas items.");
    }

    [Fact]
    public void CanvasTextBox_RendersStyledBoxAndClippedTextAtFixedCoordinates() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Premium text box", 30, 40, 140, 50, new PdfCanvasTextBoxStyle {
                Background = PdfColor.FromRgb(245, 250, 255),
                BorderColor = PdfColor.FromRgb(24, 96, 160),
                BorderWidth = 1.5D,
                PaddingX = 8D,
                PaddingY = 6D,
                FontSize = 10D,
                TextColor = PdfColor.FromRgb(20, 40, 70),
                Align = PdfAlign.Center
            }))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("30 90 140 50 re", content, StringComparison.Ordinal);
        Assert.Contains("1.5 w", content, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.InRange(FindWordStartX(page, "Premium"), 61D, 91D);
        Assert.InRange(FindWordStartY(page, "Premium"), 120D, 135D);
    }

    [Fact]
    public void CanvasDiagnosticOverloads_PreservePreviousClrSignatures() {
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.TextBox), new[] {
            typeof(string),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfCanvasTextBoxStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.TextBox), new[] {
            typeof(IEnumerable<PdfTextRun>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfCanvasTextBoxStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.Table), new[] {
            typeof(IEnumerable<string[]>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfTableStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.Table), new[] {
            typeof(IEnumerable<PdfTableCell[]>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfTableStyle),
            typeof(double)
        }));
    }

    [Fact]
    public void CanvasTextBox_ReportsClippedContentDuringRender() {
        PdfLayoutDiagnostic? diagnostic = null;
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 180,
                PageHeight = 120,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox(
                "One two three four five six seven eight nine ten eleven twelve",
                20,
                20,
                80,
                18,
                new PdfCanvasTextBoxStyle {
                    Background = null,
                    BorderColor = null,
                    FontSize = 12D,
                    PaddingX = 0D,
                    PaddingY = 0D
                },
                rotationAngle: 0D,
                diagnosticHandler: item => diagnostic = item));

        Assert.Null(diagnostic);

        document.ToBytes();

        Assert.NotNull(diagnostic);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, diagnostic!.Kind);
        Assert.Equal("PdfCanvasTextBox", diagnostic.Source);
        Assert.True(diagnostic.HasBounds);
        Assert.Equal(20D, diagnostic.X);
        Assert.Equal(20D, diagnostic.Y);
        Assert.Equal(80D, diagnostic.Width);
        Assert.Equal(18D, diagnostic.Height);
    }

    [Fact]
    public void CanvasTextBox_WithAsymmetricPadding_UsesIndividualEdges() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Asymmetric", 30, 40, 140, 50, new PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                PaddingLeft = 20D,
                PaddingRight = 4D,
                PaddingTop = 6D,
                PaddingBottom = 2D,
                FontSize = 10D
            }))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("50 92 116 42 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_UsesConfiguredVerticalAlignmentInsideFixedFrame() {
        static PdfCanvasTextBoxStyle Style(PdfVerticalAlign verticalAlign) =>
            new PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                PaddingX = 0D,
                PaddingY = 0D,
                FontSize = 10D,
                LineHeight = 12D,
                VerticalAlign = verticalAlign
            };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 200,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .TextBox("TopAlign", 20, 30, 90, 90, Style(PdfVerticalAlign.Top))
                .TextBox("MiddleAlign", 130, 30, 90, 90, Style(PdfVerticalAlign.Middle))
                .TextBox("BottomAlign", 240, 30, 90, 90, Style(PdfVerticalAlign.Bottom)))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double topY = FindWordStartY(page, "TopAlign");
        double middleY = FindWordStartY(page, "MiddleAlign");
        double bottomY = FindWordStartY(page, "BottomAlign");

        Assert.True(topY > middleY + 30D, $"Expected middle-aligned text to render lower than top-aligned text. Top: {topY:0.##}, middle: {middleY:0.##}.");
        Assert.True(middleY > bottomY + 30D, $"Expected bottom-aligned text to render lower than middle-aligned text. Middle: {middleY:0.##}, bottom: {bottomY:0.##}.");
    }

    [Fact]
    public void CanvasTextBox_RejectsInvalidVerticalAlignment() {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new PdfCanvasTextBoxStyle {
                VerticalAlign = (PdfVerticalAlign)99
            });

        Assert.Contains("Canvas text box vertical alignment must be Top, Middle, or Bottom.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_RendersBackgroundBeforeTextAndFollowingShape() {
        var shape = OfficeShape.Rectangle(25, 20);
        shape.FillColor = PdfColor.FromRgb(255, 255, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(30, 64, 175).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .TextBox("Layered", 30, 40, 120, 42, new PdfCanvasTextBoxStyle {
                    Background = PdfColor.FromRgb(250, 250, 250),
                    BorderColor = PdfColor.FromRgb(75, 85, 99),
                    BorderWidth = 1D,
                    FontSize = 10D
                })
                .Shape(shape, 40, 48))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int textBoxDraw = content.IndexOf("30 98 120 42 re", StringComparison.Ordinal);
        int textStart = content.IndexOf("BT", textBoxDraw, StringComparison.Ordinal);
        int followingShapeDraw = content.IndexOf("40 112 25 20 re", StringComparison.Ordinal);

        Assert.True(textBoxDraw >= 0, "Expected the text box background rectangle to be present.");
        Assert.True(textStart > textBoxDraw, "Expected text box text to render after its own background.");
        Assert.True(followingShapeDraw > textStart, "Expected later canvas items to render after the complete text box.");
    }

    [Fact]
    public void CanvasTable_RendersFixedFrameStyledCellsAndText() {
        var style = new PdfTableStyle {
            HeaderRowCount = 1,
            RowStripeFill = null,
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 70D, 50D },
            RowMinHeights = new System.Collections.Generic.List<double?> { 24D, 36D },
            CellFills = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor> {
                [(1, 1)] = PdfColor.FromRgb(230, 245, 255)
            },
            CellPaddings = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding> {
                [(1, 1)] = new PdfCellPadding { Left = 8D, Right = 8D, Top = 4D, Bottom = 4D }
            },
            CellAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(1, 1)] = PdfColumnAlign.Center
            },
            CellVerticalAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(1, 1)] = PdfCellVerticalAlign.Middle
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "Name", "Score" },
                new[] { "OfficeIMO", "99" }
            }, 30, 30, 120, 60, style))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 90 120 60 re", raw, StringComparison.Ordinal);
        Assert.Contains("100 150 m", raw, StringComparison.Ordinal);
        Assert.Contains("100 90 l", raw, StringComparison.Ordinal);
        Assert.Contains("30 126 m", raw, StringComparison.Ordinal);
        Assert.Contains("150 126 l", raw, StringComparison.Ordinal);
        Assert.Contains("100 90 50 36 re", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Name", text, StringComparison.Ordinal);
        Assert.Contains("OfficeIMO", text, StringComparison.Ordinal);
        Assert.Contains("99", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_ReportsClippedCellContentDuringRender() {
        var diagnostics = new List<PdfLayoutDiagnostic>();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 140,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(
                new[] {
                    new[] {
                        PdfTableCell.TextCell("One two three four five six seven eight nine ten eleven twelve")
                    }
                },
                20,
                20,
                80,
                22,
                new PdfTableStyle {
                    RowMinHeights = new List<double?> { 22D },
                    ColumnWidthPoints = new List<double?> { 80D },
                    CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
                        [(0, 0)] = new PdfCellPadding { Left = 2D, Right = 2D, Top = 2D, Bottom = 2D }
                    }
                },
                rotationAngle: 0D,
                diagnosticHandler: diagnostics.Add));

        Assert.Empty(diagnostics);

        document.ToBytes();

        PdfLayoutDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, diagnostic.Kind);
        Assert.Equal("PdfCanvasTableCell", diagnostic.Source);
        Assert.True(diagnostic.HasBounds);
        Assert.Equal(20D, diagnostic.X);
        Assert.Equal(20D, diagnostic.Y);
        Assert.Equal(80D, diagnostic.Width);
        Assert.Equal(22D, diagnostic.Height);
    }

    [Fact]
    public void CanvasTable_WithRotation_RendersInsideRotatedFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "Name", "Score" },
                new[] { "OfficeIMO", "99" }
            }, 30, 30, 120, 60, rotationAngle: 90D))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int transform = raw.IndexOf("0 1 -1 0 210 30 cm", StringComparison.Ordinal);
        int tableRect = raw.IndexOf("30 90 120 60 re", StringComparison.Ordinal);

        Assert.True(transform >= 0, "Expected a rotation matrix around the declared table frame center.");
        Assert.True(tableRect > transform, "Expected table geometry to render inside the rotated frame.");
    }

    [Fact]
    public void CanvasTable_RendersRichCellImagesAndFormControls() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    "Assets",
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 12, 12) },
                    checkBoxes: new[] { new PdfTableCellCheckBox("Canvas.Approved", isChecked: true, size: 10) },
                    formFields: new[] { PdfTableCellFormField.TextField("Canvas.Owner", "Ada", width: 44, height: 12, fontSize: 8) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Approved" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Owner" && field.IsTextField && field.Value == "Ada");
    }

    [Fact]
    public void CanvasTable_WithRotation_RotatesRichCellImagesAndFormControls() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    "Assets",
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 12, 12) },
                    checkBoxes: new[] { new PdfTableCellCheckBox("Canvas.Rotated", isChecked: true, size: 10) },
                    formFields: new[] { PdfTableCellFormField.TextField("Canvas.RotatedOwner", "Ada", width: 44, height: 12, fontSize: 8) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            }, rotationAngle: 90D))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int tableTransform = raw.IndexOf("0 1 -1 0", StringComparison.Ordinal);
        int imageDraw = raw.IndexOf("/Im1 Do", StringComparison.Ordinal);
        Assert.True(tableTransform >= 0, "Expected a rotation matrix around the declared table frame center.");
        Assert.True(imageDraw > tableTransform, "Expected the cell image to render inside the rotated table frame.");
        Assert.Contains("12 0 0 12", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("0 12 -12 0", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Rotated" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.RotatedOwner" && field.IsTextField && field.Value == "Ada");
    }

    [Fact]
    public void CanvasClip_ClipsInlineTableImages() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    string.Empty,
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 40, 40) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(50, 24, 34, 86, clipped => clipped.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            })))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int clip = raw.IndexOf("50 50 34 86 re W", StringComparison.Ordinal);
        int imageDraw = raw.IndexOf("/Im1 Do", StringComparison.Ordinal);
        Assert.True(clip >= 0, "Expected the canvas clip path in the page content stream.");
        Assert.True(imageDraw > clip, "Expected the table-cell image to render inside the canvas clip state.");
        Assert.DoesNotContain("50 90 20 40 re W", raw, StringComparison.Ordinal);

        Assert.Empty(PdfInspector.Inspect(bytes).FormFields);
    }

    [Fact]
    public void CanvasClip_RequiresWidgetsToBeFullyContainedByRectangularClips() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.Clip(
            20D,
            20D,
            50D,
            30D,
            clipped => clipped.TextField("Partial", "Ada", 50D, 10D, 40D, 20D))).ToBytes());

        byte[] contained = PdfDocument.Create().Canvas(canvas => canvas.Clip(
            20D,
            20D,
            80D,
            40D,
            clipped => clipped.TextField("Contained", "Ada", 30D, 30D, 40D, 20D))).ToBytes();
        Assert.Single(PdfInspector.Inspect(contained).FormFields);

        OfficeClipPath triangle = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(80D, 0D),
            OfficePathCommand.LineTo(40D, 40D),
            OfficePathCommand.Close());
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.Clip(
            20D,
            20D,
            triangle,
            clipped => clipped.CheckBox("NonRectangular", true, 50D, 25D, 12D, 12D))).ToBytes());
    }

    [Fact]
    public void CanvasClip_EmptyPathSuppressesFormFields() {
        byte[] bytes = PdfDocument.Create().Canvas(canvas => canvas.Clip(
            20D,
            20D,
            OfficeClipPath.Empty(),
            clipped => clipped
                .TextField("HiddenText", "Ada", 20D, 20D, 80D, 20D)
                .CheckBox("HiddenCheck", true, 20D, 50D, 12D, 12D)))
            .ToBytes();

        Assert.Empty(PdfInspector.Inspect(bytes).FormFields);
    }

    [Fact]
    public void CanvasClip_ClipsVisualAnnotationsInsideFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(20, 20, 100, 80, clipped => clipped
                .TextAnnotation("Clipped text annotation", 10, 10, 40, 30)
                .TextAnnotation("Outside text annotation", 140, 20, 20, 20)
                .FreeTextAnnotation("Clipped free text annotation", 30, 50, 160, 50)
                .HighlightAnnotation("Clipped highlight annotation", 110, 90, 40, 20)))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfAnnotation text = Assert.Single(info.GetAnnotationsBySubtype("Text"));
        PdfAnnotation freeText = Assert.Single(info.GetAnnotationsBySubtype("FreeText"));
        PdfAnnotation highlight = Assert.Single(info.GetAnnotationsBySubtype("Highlight"));

        Assert.Equal("Clipped text annotation", text.Contents);
        AssertClose(20D, text.X1);
        AssertClose(120D, text.Y1);
        AssertClose(50D, text.X2);
        AssertClose(140D, text.Y2);
        Assert.Equal("Clipped free text annotation", freeText.Contents);
        AssertClose(30D, freeText.X1);
        AssertClose(60D, freeText.Y1);
        AssertClose(120D, freeText.X2);
        AssertClose(110D, freeText.Y2);
        Assert.Equal("Clipped highlight annotation", highlight.Contents);
        AssertClose(110D, highlight.X1);
        AssertClose(60D, highlight.Y1);
        AssertClose(120D, highlight.X2);
        AssertClose(70D, highlight.Y2);
    }

    [Fact]
    public void CanvasClip_SuppressesVisualAnnotationsOutsideNonRectangularRegion() {
        OfficeClipPath triangle = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.LineTo(0D, 80D),
            OfficePathCommand.Close());
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160
            })
            .Canvas(canvas => canvas.Clip(20D, 20D, triangle, clipped => clipped
                .TextAnnotation("Visible text", 25D, 25D, 10D, 10D)
                .TextAnnotation("Hidden text", 100D, 80D, 10D, 10D)
                .FreeTextAnnotation("Visible free text", 40D, 25D, 15D, 10D)
                .FreeTextAnnotation("Hidden free text", 85D, 75D, 15D, 10D)
                .HighlightAnnotation("Visible highlight", 25D, 45D, 15D, 8D)
                .HighlightAnnotation("Hidden highlight", 90D, 65D, 15D, 8D)
                .Image(CreateMinimalRgbPng(), 50D, 40D, 40D, 30D, linkUri: "https://example.com/partial")))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal("Visible text", Assert.Single(info.GetAnnotationsBySubtype("Text")).Contents);
        Assert.Equal("Visible free text", Assert.Single(info.GetAnnotationsBySubtype("FreeText")).Contents);
        Assert.Equal("Visible highlight", Assert.Single(info.GetAnnotationsBySubtype("Highlight")).Contents);
        PdfLinkAnnotation link = Assert.Single(info.GetLinkAnnotationsByUri("https://example.com/partial"));
        Assert.True(link.Width < 40D || link.Height < 30D);
        Assert.True(link.Width > 0D && link.Height > 0D);
        const double clipBottomY = 60D;
        foreach (double pageX in new[] { link.X1, link.X2 }) {
            foreach (double pageY in new[] { link.Y1, link.Y2 }) {
                double localX = pageX - 20D;
                double localY = 80D - (pageY - clipBottomY);
                Assert.True(localX / 100D + localY / 80D <= 1.001D);
            }
        }
    }

    [Fact]
    public void CanvasClip_PreservesInlineImageClipPathInsideFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(20, 20, 100, 80, clipped => clipped.Image(CreateMinimalRgbPng(), 30, 30, 40, 40, new PdfImageStyle {
                ClipPath = OfficeClipPath.Rectangle(20, 20)
            })))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 60 100 80 re W", raw, StringComparison.Ordinal);
        Assert.Contains("30 110 20 20 re W", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("30 90 40 40 re W", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasClip_AcceptsRoundedSharedClipPath() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(20, 20, OfficeClipPath.RoundedRectangle(100, 80, 10), clipped =>
                clipped.Image(CreateMinimalRgbPng(), 20, 20, 100, 80)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.Contains(" W n\n", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("20 60 100 80 re W", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasClip_AcceptsFreeformSharedClipPath() {
        OfficeClipPath triangle = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.LineTo(100, 0),
            OfficePathCommand.LineTo(50, 80),
            OfficePathCommand.Close());
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(20, 20, triangle, clipped =>
                clipped.Image(CreateMinimalRgbPng(), 20, 20, 100, 80)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 140 m 120 140 l 70 60 l h W* n", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("20 60 100 80 re W", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_SkipsVerticalGridDividersInsideMergedCells() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { PdfTableCell.Span("Merged", 2) },
                new[] { PdfTableCell.TextCell("Left"), PdfTableCell.TextCell("Right") }
            }, 30, 30, 120, 60))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 90 120 60 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("90 150 m", raw, StringComparison.Ordinal);
        Assert.Contains("90 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("90 90 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_RejectsUnboundedLogicalGridBeforeRendering() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { PdfTableCell.Span("Oversized", 262145) }
            }, 30, 30, 120, 60));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains("exceeding the supported limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_RowSpanSkipsContinuationRowStripeFill() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { PdfTableCell.Merge("Span", rowSpan: 2), PdfTableCell.TextCell("Top") },
                new[] { PdfTableCell.TextCell("Bottom") }
            }, 30, 30, 120, 60, new PdfTableStyle {
                HeaderRowCount = 0,
                RowStripeFill = PdfColor.FromRgb(220, 235, 250),
                ColumnWidthPoints = new System.Collections.Generic.List<double?> { 60D, 60D },
                RowMinHeights = new System.Collections.Generic.List<double?> { 30D, 30D }
            }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain("30 90 120 30 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("30 90 60 30 re", raw, StringComparison.Ordinal);
        Assert.Contains("90 90 60 30 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_WithRotation_RendersBoxAndTextInsideRotatedGroup() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Rotated box", 30, 40, 120, 42, new PdfCanvasTextBoxStyle {
                Background = PdfColor.FromRgb(250, 250, 250),
                BorderColor = PdfColor.FromRgb(75, 85, 99),
                BorderWidth = 1D,
                FontSize = 10D
            }, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int transform = content.IndexOf("0 1 -1 0 209 29 cm", StringComparison.Ordinal);
        int rectangle = content.IndexOf("30 98 120 42 re", StringComparison.Ordinal);
        int textStart = content.IndexOf("BT", rectangle, StringComparison.Ordinal);

        Assert.True(transform >= 0, "Expected a rotation matrix around the declared text box frame center.");
        Assert.True(rectangle > transform, "Expected the text box geometry to render inside the rotated group.");
        Assert.True(textStart > rectangle, "Expected text to render after the rotated text box background.");
    }

    [Fact]
    public void CanvasTextBox_InvalidPadding_ThrowsClearDiagnostic() {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 100,
                    PageHeight = 100
                })
                .Canvas(canvas => canvas.TextBox("Bad", 0, 0, 20, 10, new PdfCanvasTextBoxStyle {
                    PaddingY = 5D
                })));

        Assert.Contains("Canvas text box padding must leave a positive text area.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasRotation_NonFiniteAngle_ThrowsClearDiagnostic() {
        var shape = OfficeShape.Rectangle(10, 10);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.TextBox("Bad", 0, 0, 10, 10, rotationAngle: double.NegativeInfinity)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Shape(shape, 0, 0, rotationAngle: double.NaN)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 0, 0, 10, 10, rotationAngle: double.PositiveInfinity)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Table(new[] { new[] { "Bad" } }, 0, 0, 10, 10, rotationAngle: double.NaN)));
    }

    [Fact]
    public void CanvasTextBox_WithRotationAndLinkedRun_RotatesLinkAnnotationBounds() {
        PdfOptions options = CreateCanvasOptions();
        const string uri = "https://evotec.xyz/canvas-textbox";
        var style = new PdfCanvasTextBoxStyle {
            FontSize = 10D
        };

        byte[] flatBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.TextBox(new[] {
                PdfTextRun.Link("Linked", uri)
            }, 30, 40, 120, 42, style))
            .ToBytes();
        byte[] rotatedBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.TextBox(new[] {
                PdfTextRun.Link("Linked", uri)
            }, 30, 40, 120, 42, style, rotationAngle: 90D))
            .ToBytes();

        PdfLinkAnnotation flatLink = Assert.Single(PdfInspector.Inspect(flatBytes).LinkAnnotations);
        PdfLinkAnnotation rotatedLink = Assert.Single(PdfInspector.Inspect(rotatedBytes).LinkAnnotations);
        var expected = RotateRectangle(flatLink, 30, 98, 120, 42, 90D);

        Assert.Equal(uri, rotatedLink.Uri);
        AssertClose(expected.X1, rotatedLink.X1);
        AssertClose(expected.Y1, rotatedLink.Y1);
        AssertClose(expected.X2, rotatedLink.X2);
        AssertClose(expected.Y2, rotatedLink.Y2);
    }

    [Fact]
    public void CanvasImage_WithRotationAndLink_RotatesLinkAnnotationBounds() {
        PdfOptions options = CreateCanvasOptions();
        const string uri = "https://evotec.xyz/canvas-image";

        byte[] flatBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, linkUri: uri))
            .ToBytes();
        byte[] rotatedBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, linkUri: uri, rotationAngle: 90D))
            .ToBytes();

        PdfLinkAnnotation flatLink = Assert.Single(PdfInspector.Inspect(flatBytes).LinkAnnotations);
        PdfLinkAnnotation rotatedLink = Assert.Single(PdfInspector.Inspect(rotatedBytes).LinkAnnotations);
        var expected = RotateRectangle(flatLink, 30, 110, 60, 30, 90D);

        Assert.Equal(uri, rotatedLink.Uri);
        AssertClose(expected.X1, rotatedLink.X1);
        AssertClose(expected.Y1, rotatedLink.Y1);
        AssertClose(expected.X2, rotatedLink.X2);
        AssertClose(expected.Y2, rotatedLink.Y2);
    }

    [Fact]
    public void CanvasImage_AppliesFitAfterSourceCrop() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 100,
                PageHeight = 100,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(
                CreateMinimalRgbPng(),
                0,
                0,
                100,
                100,
                new PdfImageStyle {
                    Fit = OfficeImageFit.Contain,
                    SourceCrop = new PdfImageSourceCrop(0.5D, 0D, 0D, 0D)
                }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.Contains("100 0 0 100 -25 0 cm", raw, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0.5 1 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasEffect_WritesIsolatedFormAndTransformsSearchableLinkedText() {
        const string uri = "https://evotec.xyz/canvas-effect";
        PdfOptions options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 100,
            MarginLeft = 0,
            MarginRight = 0,
            MarginTop = 0,
            MarginBottom = 0,
            CompressContentStreams = false
        };
        byte[] flatBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Text(new[] { PdfTextRun.Link("EffectText", uri) }, 20, 20, 80, 20, fontSize: 10))
            .ToBytes();
        byte[] effectBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Effect(
                OfficeTransform.Translate(12D, 7D),
                0.5D,
                nested => nested.Text(new[] { PdfTextRun.Link("EffectText", uri) }, 20, 20, 80, 20, fontSize: 10)))
            .ToBytes();

        PdfLinkAnnotation flatLink = Assert.Single(PdfInspector.Inspect(flatBytes).LinkAnnotations);
        PdfLinkAnnotation effectLink = Assert.Single(PdfInspector.Inspect(effectBytes).LinkAnnotations);
        string raw = Encoding.ASCII.GetString(effectBytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(effectBytes));

        Assert.Contains("EffectText", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("/Group << /S /Transparency /I true /K false >>", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 12 -7 cm", raw, StringComparison.Ordinal);
        AssertClose(flatLink.X1 + 12D, effectLink.X1);
        AssertClose(flatLink.Y1 - 7D, effectLink.Y1);
        AssertClose(flatLink.X2 + 12D, effectLink.X2);
        AssertClose(flatLink.Y2 - 7D, effectLink.Y2);
    }

    [Fact]
    public void CanvasEffect_RejectsInvalidOpacity() {
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas =>
            canvas.Effect(OfficeTransform.Identity, double.NaN, _ => { })));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().Canvas(canvas =>
            canvas.Effect(OfficeTransform.Identity, 1D, null!)));
    }

    [Fact]
    public void CanvasEffect_RejectsInteractiveFieldsInNontrivialEffects() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas =>
            canvas.Effect(OfficeTransform.RotateDegrees(45D), 0.5D, nested =>
                nested.TextField("Name", "Ada", 10D, 10D, 80D, 20D))));

        byte[] identityBytes = PdfDocument.Create().Canvas(canvas =>
            canvas.Effect(OfficeTransform.Identity, 1D, nested =>
                nested.TextField("Name", "Ada", 10D, 10D, 80D, 20D))).ToBytes();
        Assert.Single(PdfInspector.Inspect(identityBytes).FormFields);
    }

    [Fact]
    public void CanvasMultiSelect_SerializesSelectionsInOptionOrder() {
        byte[] bytes = PdfDocument.Create().Canvas(canvas => canvas.ChoiceField(
            "Letters",
            new[] { "A", "B", "C" },
            new[] { "C", "A" },
            10D,
            10D,
            100D,
            50D,
            isComboBox: false,
            allowsMultipleSelection: true)).ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal(new[] { 0, 2 }, field.SelectedIndices);
        Assert.Equal(new[] { "A", "C" }, field.SelectedOptions.Select(option => option.ExportValue).ToArray());
    }

    [Fact]
    public void CanvasItem_OutsidePageBounds_ThrowsClearDiagnostic() {
        var doc = PdfDocument.Create(new PdfOptions {
                PageWidth = 100,
                PageHeight = 100,
                MarginLeft = 10,
                MarginRight = 10,
                MarginTop = 10,
                MarginBottom = 10
            })
            .Canvas(canvas => canvas.Text("Out", 90, 10, 20, 20));

        ArgumentException ex = Assert.Throws<ArgumentException>(() => doc.ToBytes());
        Assert.Contains("Canvas text exceeds the current page bounds.", ex.Message, StringComparison.Ordinal);
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static PdfOptions CreateCanvasOptions() =>
        new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            CompressContentStreams = false
        };

    private static (double X1, double Y1, double X2, double Y2) RotateRectangle(PdfLinkAnnotation rectangle, double x, double bottomY, double width, double height, double rotationAngle) {
        double angle = rotationAngle * Math.PI / 180D;
        double cos = Math.Cos(angle);
        double sin = Math.Sin(angle);
        double centerX = x + width / 2D;
        double centerY = bottomY + height / 2D;

        RotatePoint(rectangle.X1, rectangle.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
        RotatePoint(rectangle.X1, rectangle.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
        RotatePoint(rectangle.X2, rectangle.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
        RotatePoint(rectangle.X2, rectangle.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

        return (
            Math.Min(Math.Min(x1, x2), Math.Min(x3, x4)),
            Math.Min(Math.Min(y1, y2), Math.Min(y3, y4)),
            Math.Max(Math.Max(x1, x2), Math.Max(x3, x4)),
            Math.Max(Math.Max(y1, y2), Math.Max(y3, y4)));
    }

    private static void RotatePoint(double x, double y, double centerX, double centerY, double cos, double sin, out double rotatedX, out double rotatedY) {
        double dx = x - centerX;
        double dy = y - centerY;
        rotatedX = centerX + cos * dx - sin * dy;
        rotatedY = centerY + sin * dx + cos * dy;
    }

    private static void AssertClose(double expected, double actual) =>
        Assert.InRange(Math.Abs(expected - actual), 0D, 0.01D);

    private static int CountOccurrences(string value, string marker) =>
        value.Split(new[] { marker }, StringSplitOptions.None).Length - 1;

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);
}
