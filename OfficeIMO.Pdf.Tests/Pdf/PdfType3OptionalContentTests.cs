using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfType3OptionalContentTests {
    [Fact]
    public void InlineOptionalContentMembershipDictionary_ParsesReferencedGroupsAndPolicy() {
        const string content = "<< /Type /OCMD /OCGs [11 0 R % 10 0 R\n] /Extension << /P /AllOff >> /P /AnyOn /XRef 10 0 R >>";

        PdfInlineOptionalContentReferences references = PdfInlineOptionalContentReferenceParser.Parse(content, 0, content.Length);

        Assert.True(references.IsMembershipDictionary);
        Assert.Equal("AnyOn", references.Policy);
        Assert.False(references.HasInvalidPolicy);
        Assert.Equal(new[] { 11 }, references.ObjectNumbers);
    }

    [Theory]
    [InlineData("/Bad")]
    [InlineData("1")]
    public void InlineOptionalContentMembershipDictionary_RejectsMalformedPolicy(string policy) {
        string content = "<< /Type /OCMD /OCGs [11 0 R] /P " + policy + " >>";

        PdfInlineOptionalContentReferences references = PdfInlineOptionalContentReferenceParser.Parse(content, 0, content.Length);

        Assert.True(references.HasInvalidPolicy);
    }

    [Fact]
    public void RenderPage_SkipsHiddenOptionalContentInsideType3GlyphProgram() {
        byte[] pdf = BuildType3OptionalContentPdf(nestedForm: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
        Assert.Empty(drawing.Images);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingText>());
    }

    [Fact]
    public void RenderPage_SkipsHiddenOptionalContentInsideNestedType3Form() {
        byte[] pdf = BuildType3OptionalContentPdf(nestedForm: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingGroup clipped = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingShape visible = Assert.Single(clipped.Drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
        Assert.Empty(drawing.Images);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingText>());
    }

    [Theory]
    [InlineData("<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn >>", true)]
    [InlineData("<< /Type /OCMD /OCGs 10 0 R /P /AllOn >>", true)]
    [InlineData("<< /Type /O#43MD /OCGs [10 0 R] /P /Any#4Fn >>", true)]
    [InlineData("<< /Type /OCMD /OCGs [11 0 R % 10 0 R\n] /P /AllOn >>", false)]
    [InlineData("<< /Type /OCMD /OCGs [11 0 R] /P /AllOn /XRef 10 0 R >>", false)]
    [InlineData("<< /Type /OCMD /OCGs [11 0 R] /Extension << /P /AllOff >> /P /AnyOn >>", false)]
    [InlineData("<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE [/N#6Ft 10 0 R] >>", false)]
    [InlineData("<< /Type /OCMD /OCGs [11 0 R] /P /AnyOn /VE [/Or % hidden operand\n 10 0 R] >>", true)]
    public void RenderPage_EvaluatesInlineMembershipDictionaryInsideType3GlyphProgram(
        string membershipDictionary,
        bool expectHidden) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: membershipDictionary,
            includeUnsupportedConditionalContent: expectHidden);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        if (expectHidden) {
            OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
            Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
        } else {
            Assert.Equal(new OfficeColor?[] { OfficeColor.Red, OfficeColor.Lime }, drawing.Shapes.Select(shape => shape.Shape.FillColor));
        }

        Assert.Empty(drawing.Images);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingText>());
    }

    [Fact]
    public void RenderPage_SkipsHiddenInlineMembershipDictionaryInsideNestedType3Form() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: true,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingGroup clipped = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingShape visible = Assert.Single(clipped.Drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
        Assert.Empty(drawing.Images);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingText>());
    }

    [Fact]
    public void RenderPage_EvaluatesIndirectVisibilityExpressionBeforeMembershipPolicy() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [11 0 R] /P /AnyOn /VE 12 0 R >>",
            indirectVisibilityExpression: "[/Not 11 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
    }

    [Fact]
    public void RenderPage_FailsClosedForMismatchedVisibilityExpressionGeneration() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE 12 1 R >>",
            indirectVisibilityExpression: "[/Not 10 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMismatchedNestedVisibilityGeneration() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE 12 0 R >>",
            indirectVisibilityExpression: "[/Not 10 1 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMismatchedMembershipGroupGeneration() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 1 R] /P /AnyOn >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Bad")]
    [InlineData("1")]
    public void RenderPage_FailsClosedForMalformedInlineMembershipPolicy(string policy) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P " + policy + " >>",
            includeUnsupportedConditionalContent: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullInlineMembershipPolicyAsDefaultAnyOn() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P null >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
    }

    [Theory]
    [InlineData("[/Or]")]
    [InlineData("10 0 R")]
    [InlineData("10 0 Rubbish")]
    public void RenderPage_FailsClosedForMalformedInlineVisibilityExpression(string expression) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE " + expression + " >>",
            includeUnsupportedConditionalContent: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMembershipReferenceOutsideConfiguredOptionalContentGroups() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [12 0 R] /P /AnyOn >>",
            indirectVisibilityExpression: "<< /Type /OCG /Name (Unconfigured layer) >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_EvaluatesChainedIndirectVisibilityExpressionBeforeMembershipPolicy() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [11 0 R] /P /AnyOn /VE 12 0 R >>",
            indirectVisibilityExpression: "13 0 R",
            secondaryVisibilityExpression: "[/Not 11 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedIndirectNotExpression() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE 12 0 R >>",
            indirectVisibilityExpression: "[/Not 10 0 R 11 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_EvaluatesIndirectOperandInsideInlineVisibilityExpression() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [11 0 R] /P /AnyOn /VE [/Not 12 0 R] >>",
            indirectVisibilityExpression: "[/Not 10 0 R]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenInlineVisibilityExpressionExceedsNestingLimit() {
        string expression = "10 0 R";
        for (int depth = 0; depth < 129; depth++) expression = "[/Not " + expression + "]";
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE " + expression + " >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 256 }
        };

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, readOptions: options));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenIndirectVisibilityExpressionExceedsNestingLimit() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE 12 0 R >>",
            indirectVisibilityChainLength: 130);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AllowsRepeatedIndirectVisibilitySubexpressionAcrossSiblings() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOn /VE 12 0 R >>",
            includeUnsupportedConditionalContent: false,
            indirectVisibilityExpression: "[/And 13 0 R 13 0 R]",
            secondaryVisibilityExpression: "[/Not 10 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(new OfficeColor?[] { OfficeColor.Red, OfficeColor.Lime }, drawing.Shapes.Select(shape => shape.Shape.FillColor));
    }

    [Fact]
    public void RenderPage_EvaluatesInlineAnyOffWhenEveryOptionalContentGroupIsOn() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            inlineMembershipDictionary: "<< /Type /OCMD /OCGs [10 0 R] /P /AnyOff >>",
            includeUnsupportedConditionalContent: true,
            allGroupsOn: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingShape visible = Assert.Single(drawing.Shapes);
        Assert.Equal(OfficeColor.Lime, visible.Shape.FillColor);
    }

    [Fact]
    public void RenderPage_IgnoresInexactDashUsedOnlyByHiddenOptionalContent() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            hiddenExtraContent: " [3 1] 2 d 0 0 m 500 700 l S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_EvaluatesResourceMembershipVisibilityExpressionWithoutOcgs() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            resourceMembershipDictionary: "<< /Type /OCMD /VE [/And 10 0 R] >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("<< /Type /OCMD /OCGs [10 0 R] /P /Bad >>")]
    [InlineData("<< /Type /OCMD /OCGs [10 0 R] /VE [/Or] >>")]
    [InlineData("<< /Type /OCMD /OCGs [10 1 R] /P /AnyOn >>")]
    public void RenderPage_FailsClosedForMalformedResourceMembershipDictionary(string membershipDictionary) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            resourceMembershipDictionary: membershipDictionary);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMissingNamedOptionalContentProperty() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            omitPropertyResource: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/Type /Bad")]
    public void RenderPage_FailsClosedForMalformedOptionalContentGroupDeclaration(string hiddenGroupType) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            hiddenGroupType: hiddenGroupType);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/ON [10 1 R] /OFF []")]
    [InlineData("/ON [] /OFF [10 1 R]")]
    public void RenderPage_FailsClosedForMismatchedDefaultConfigurationGroupGeneration(string defaultConfigurationEntries) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            defaultConfigurationEntries: defaultConfigurationEntries);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Bad")]
    [InlineData("1")]
    public void RenderPage_FailsClosedForMalformedDefaultConfigurationBaseState(string baseState) {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            defaultConfigurationEntries: "/BaseState " + baseState + " /ON [11 0 R] /OFF [10 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullDefaultConfigurationBaseStateAsAbsent() {
        byte[] pdf = BuildType3OptionalContentPdf(
            nestedForm: false,
            defaultConfigurationEntries: "/BaseState null /ON [11 0 R] /OFF [10 0 R]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    private static byte[] BuildType3OptionalContentPdf(
        bool nestedForm,
        string? inlineMembershipDictionary = null,
        bool includeUnsupportedConditionalContent = true,
        string hiddenExtraContent = "",
        string? indirectVisibilityExpression = null,
        bool allGroupsOn = false,
        string? secondaryVisibilityExpression = null,
        int indirectVisibilityChainLength = 0,
        string? resourceMembershipDictionary = null,
        string hiddenGroupType = "/Type /OCG",
        string? defaultConfigurationEntries = null,
        bool omitPropertyResource = false) {
        string hiddenProperty = inlineMembershipDictionary ?? (resourceMembershipDictionary is null ? "/Hidden" : "/Membership");
        string unsupportedConditionalContent = includeUnsupportedConditionalContent
            ? " BT /Missing 12 Tf (Hidden) Tj ET /Missing gs /Missing Do"
            : string.Empty;
        string hiddenAndVisibleContent =
            "/OC " + hiddenProperty + " BDC 1 0 0 rg 0 0 500 700 re f" + hiddenExtraContent + unsupportedConditionalContent + " EMC " +
            "0 1 0 rg 250 0 250 700 re f";
        string type3Resources = nestedForm
            ? "<< /XObject << /Fm1 7 0 R >> >>"
            : omitPropertyResource ? "<< >>"
            : inlineMembershipDictionary is not null ? "<< >>" : resourceMembershipDictionary is not null
                ? "<< /Properties << /Membership 14 0 R >> >>"
                : "<< /Properties << /Hidden 10 0 R >> >>";
        string glyphContent = nestedForm ? "500 0 d0 /Fm1 Do" : "500 0 d0 " + hiddenAndVisibleContent;
        string defaultConfiguration = defaultConfigurationEntries ??
            "/ON [" + (allGroupsOn ? "10 0 R " : string.Empty) + "11 0 R] /OFF [" + (allGroupsOn ? string.Empty : "10 0 R") + "]";
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [10 0 R 11 0 R] /D << " + defaultConfiguration + " >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", "BT /FType3 18 Tf 20 100 Td (A) Tj ET"),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources " + type3Resources + " >>\nendobj",
            StreamObject(6, "<<", glyphContent)
        };
        if (nestedForm) {
            objects.Add(StreamObject(
                7,
                "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources " +
                (inlineMembershipDictionary is not null ? "<< >>" : resourceMembershipDictionary is not null
                    ? "<< /Properties << /Membership 14 0 R >> >>"
                    : "<< /Properties << /Hidden 10 0 R >> >>"),
                hiddenAndVisibleContent));
        }
        objects.Add("10 0 obj\n<< " + hiddenGroupType + " /Name (Hidden Type 3 layer) >>\nendobj");
        objects.Add("11 0 obj\n<< /Type /OCG /Name (Visible Type 3 layer) >>\nendobj");
        if (indirectVisibilityChainLength > 0) {
            for (int index = 0; index < indirectVisibilityChainLength; index++) {
                int objectNumber = 12 + index;
                string value = index + 1 == indirectVisibilityChainLength
                    ? "[/Not 10 0 R]"
                    : (objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R";
                objects.Add(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" + value + "\nendobj");
            }
        } else if (indirectVisibilityExpression is not null) {
            objects.Add("12 0 obj\n" + indirectVisibilityExpression + "\nendobj");
        }
        if (secondaryVisibilityExpression is not null) {
            objects.Add("13 0 obj\n" + secondaryVisibilityExpression + "\nendobj");
        }
        if (resourceMembershipDictionary is not null) {
            objects.Add("14 0 obj\n" + resourceMembershipDictionary + "\nendobj");
        }
        return Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
    }

    private static string StreamObject(int number, string dictionaryPrefix, string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        return number.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" +
               dictionaryPrefix + " /Length " + length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
               " >>\nstream\n" + content + "\nendstream\nendobj";
    }
}
