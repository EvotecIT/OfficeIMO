using System.Text;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyTests {
    [Fact]
    public void BrowserAdversarialFixtureCoversEveryClaimedSvgVisibilityLane() {
        byte[] svg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "svg-concealed-adversarial.svg"));

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE CONTROL");
        Assert.Contains(report.Findings, item => item.TextPreview == "CSS DISPLAY PAYLOAD" && item.Kind == OfficeContentConcealmentKind.HiddenByProperty);
        Assert.Contains(report.Findings, item => item.TextPreview == "DEEP SELECTOR PAYLOAD" && item.Kind == OfficeContentConcealmentKind.HiddenByProperty);
        Assert.Contains(report.Findings, item => item.TextPreview == "PRESENTATION VISIBILITY PAYLOAD" && item.Kind == OfficeContentConcealmentKind.HiddenByProperty);
        Assert.Contains(report.Findings, item => item.TextPreview == "OPACITY PAYLOAD" && item.Kind == OfficeContentConcealmentKind.TransparentText);
        Assert.Contains(report.Findings, item => item.TextPreview == "CLIPPED PAYLOAD" && item.Kind == OfficeContentConcealmentKind.ClippedContent);
        Assert.Contains(report.Findings, item => item.TextPreview == "OFF CANVAS PAYLOAD" && item.Kind == OfficeContentConcealmentKind.OffCanvas);
        Assert.Contains(report.Findings, item => item.TextPreview == "LOW CONTRAST PAYLOAD" && item.Kind == OfficeContentConcealmentKind.LowContrastText);
        Assert.Contains(report.Findings, item => item.TextPreview == "PAINT ORDER PAYLOAD" &&
            (item.Kind == OfficeContentConcealmentKind.LowContrastText || item.Kind == OfficeContentConcealmentKind.Other));
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE CSS OVERRIDE");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE VISIBILITY OVERRIDE");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE MISSING VAR OVERRIDE");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE INITIAL OVERRIDE");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE INHERITED VAR FALLBACK");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE FOREIGN NAMESPACE CONTROL");
    }

    [Fact]
    public void InspectContentSafetyResolvesPresentationAndStylesheetVisibility() {
        byte[] svg = Svg("""
            <style>
              .css-hidden { display: none; }
              #transparent { opacity: 0; }
            </style>
            <text class="css-hidden" x="10" y="25">stylesheet hidden</text>
            <text visibility="hidden" x="10" y="50">presentation hidden</text>
            <text id="transparent" x="10" y="75">transparent text</text>
            <text x="10" y="100">ordinary visible</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.HiddenByProperty && item.TextPreview == "stylesheet hidden");
        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.HiddenByProperty && item.TextPreview == "presentation hidden");
        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.TransparentText && item.TextPreview == "transparent text");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "ordinary visible");
    }

    [Fact]
    public void InspectContentSafetyUsesResolvedGeometryAndClipping() {
        byte[] svg = Svg("""
            <defs><clipPath id="empty"><rect width="0" height="0" /></clipPath></defs>
            <text x="400" y="30">off canvas</text>
            <text x="10" y="55" transform="scale(0)">zero geometry</text>
            <text x="10" y="80" clip-path="url(#empty)">clipped payload</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.OffCanvas && item.TextPreview == "off canvas");
        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.ZeroDimension && item.TextPreview == "zero geometry");
        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.ClippedContent && item.TextPreview == "clipped payload");
    }

    [Fact]
    public void InspectContentSafetyResolvesBackgroundAndDocumentPaintOrder() {
        byte[] svg = Svg("""
            <rect width="220" height="120" fill="white" />
            <text font-family="OfficeIMO Shaping Test" x="10" y="30" fill="white">white on white</text>
            <text font-family="OfficeIMO Shaping Test" x="10" y="65" fill="black">covered by later paint</text>
            <rect x="0" y="42" width="220" height="35" fill="black" />
            """);
        int[] fontScalars = "white on whitecovered by later paint"
            .Distinct()
            .Select(character => (int)character)
            .ToArray();
        var readerOptions = new OfficeSvgDrawingReaderOptions();
        readerOptions.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(fontScalars));

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.LowContrastText &&
            item.TextPreview == "white on white" && item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
        Assert.Contains(report.Findings, item => item.TextPreview == "covered by later paint" &&
            (item.Kind == OfficeContentConcealmentKind.LowContrastText || item.Kind == OfficeContentConcealmentKind.Other));
    }

    [Fact]
    public void RemoveSelectedContentRemovesOnlyTheSelectedTextNodeAndReopens() {
        byte[] svg = Svg("""
            <text style="display:none" x="10" y="35">remove only me</text>
            <text x="10" y="70">keep me</text>
            """);
        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);
        OfficeContentSafetyFinding hidden = Assert.Single(report.Findings, item => item.TextPreview == "remove only me");

        OfficeContentCleanupResult cleaned = OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { hidden.Id }));

        Assert.True(cleaned.Changed);
        Assert.DoesNotContain(cleaned.After.Findings, item => item.TextPreview == "remove only me");
        string output = Encoding.UTF8.GetString(cleaned.Output);
        Assert.DoesNotContain("remove only me", output, StringComparison.Ordinal);
        Assert.Contains("keep me", output, StringComparison.Ordinal);
        Assert.True(OfficeSvgDrawingReader.TryRead(cleaned.Output, out OfficeDrawing? reopened));
        Assert.NotNull(reopened);
    }

    [Fact]
    public void RemoveSelectedContentCanRemoveExactUnicodeWithoutDeletingVisibleText() {
        byte[] svg = Svg("""
            <text x="10" y="35">pay​load remains</text>
            """);
        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);
        OfficeContentSafetyFinding unicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u200B");

        OfficeContentCleanupResult cleaned = OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { unicode.Id }));

        string output = Encoding.UTF8.GetString(cleaned.Output);
        Assert.Contains("payload remains", output, StringComparison.Ordinal);
        Assert.DoesNotContain("pay​load", output, StringComparison.Ordinal);
        Assert.DoesNotContain(cleaned.After.Findings, item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode);
    }

    [Fact]
    public void EmptySelectionPreservesOriginalSvgBytes() {
        byte[] svg = Encoding.Unicode.GetPreamble()
            .Concat(Encoding.Unicode.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'><text x='1' y='10'>visible</text></svg>"))
            .ToArray();

        OfficeContentCleanupResult result = OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(Array.Empty<string>()));

        Assert.False(result.Changed);
        Assert.Equal(svg, result.Output);
    }

    [Fact]
    public void MetadataAndStylesheetTextRemainReportOnly() {
        byte[] svg = Svg("""
            <metadata>machine metadata</metadata>
            <style>.ordinary { fill: black; }</style>
            <text class="ordinary" x="10" y="35">visible</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item => item.TextPreview == "machine metadata" && item.CleanupCapability == OfficeContentCleanupCapability.ReportOnly);
        Assert.Contains(report.Findings, item => item.TextPreview.Contains("ordinary", StringComparison.Ordinal) && item.CleanupCapability == OfficeContentCleanupCapability.ReportOnly);
    }

    [Fact]
    public void NonPrimarySvgTextCanBeExcludedWithoutHidingPaintedTextIntegrityEvidence() {
        byte[] svg = Svg("""
            <title>machine title</title>
            <metadata>machine metadata</metadata>
            <text x="10" y="35">painted​text</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(
            svg,
            new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false });

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "machine title");
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "machine metadata");
        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u200B");
    }

    [Fact]
    public void CleanupRejectsReportOnlyAndForeignSnapshotSelections() {
        byte[] svg = Svg("""
            <metadata>machine metadata</metadata>
            <text style="display:none" x="10" y="35">first payload</text>
            """);
        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);
        OfficeContentSafetyFinding metadata = Assert.Single(report.Findings, item => item.TextPreview == "machine metadata");

        Assert.Throws<InvalidOperationException>(() => OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { metadata.Id })));

        byte[] other = Svg("""
            <text style="display:none" x="10" y="35">different payload</text>
            """);
        OfficeContentSafetyFinding foreign = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(other).Findings,
            item => item.TextPreview == "different payload");

        Assert.Throws<ArgumentException>(() => OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { foreign.Id })));
    }

    [Fact]
    public void InspectionResolvesDeepDescendantSelectors() {
        byte[] svg = Svg("""
            <style>svg g text { display: none; }</style>
            <g><text x="10" y="35">deep descendant payload</text></g>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.HiddenByProperty && item.TextPreview == "deep descendant payload");
    }

    [Fact]
    public void InlineVisibilityStillAppliesAfterTheStylesheetDeclarationBudgetIsExactlyConsumed() {
        string declarations = string.Join(";", Enumerable.Range(0, 32768).Select(index => $"--unused-{index}:0"));
        byte[] svg = Svg("<style>.never-matches { " + declarations +
            " }</style><text style='display:none' x='10' y='35'>late inline payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item => item.TextPreview == "late inline payload");
    }

    [Fact]
    public void CssBudgetOverflowFailsClosed() {
        string declarations = string.Join(";", Enumerable.Range(0, 32769).Select(index => $"--unused-{index}:0"));
        byte[] svg = Svg("<style>.never-matches { " + declarations +
            " }</style><text x='10' y='35'>ordinary visible</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CssOverridesPresentationAttributesWithoutVisibleTextFalsePositives() {
        byte[] svg = Svg("""
            <text opacity="0" style="opacity:1" x="10" y="25">opacity override visible</text>
            <text display="none" style="display:inline" x="10" y="55">display override visible</text>
            <g visibility="hidden"><text visibility="visible" x="10" y="85">visibility override visible</text></g>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        OfficeContentSafetyFinding[] falsePositives = report.Findings
            .Where(item => item.TextPreview.Contains("override visible", StringComparison.Ordinal))
            .ToArray();
        Assert.True(falsePositives.Length == 0, string.Join(" | ", falsePositives.Select(item =>
            item.TextPreview + ": " + item.Kind + " - " + item.Evidence)));
    }

    [Fact]
    public void UnsupportedCssDeclarationValuesFailClosedBeforeCleanupEvidence() {
        byte[] svg = Svg("""
            <text opacity="0" style="opacity:bogus" x="10" y="35">invalid opacity fallback</text>
            <text display="none" style="display:bogus" x="10" y="70">invalid display fallback</text>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CssOpacityPercentagesOverrideConcealingPresentationAttributes() {
        byte[] svg = Svg("""
            <text opacity="0" style="opacity:100%" x="10" y="35">percentage opacity visible</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "percentage opacity visible");
    }

    [Fact]
    public void CssCustomPropertiesInheritAcrossNeutralAncestorsAndRemainCaseSensitive() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" width="220" height="120" viewBox="0 0 220 120" style="--o:0;--O:1">
              <g><g><text style="opacity:var(--o)" x="10" y="35">inherited custom property payload</text></g></g>
            </svg>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText &&
            item.TextPreview == "inherited custom property payload");
    }

    [Fact]
    public void CssInheritMaterializesParentTransformWithoutTinyFalsePositive() {
        byte[] svg = Svg("""
            <g transform="scale(0.2,1)">
              <text style="transform:inherit" font-size="16" x="10" y="35">inherited transform payload</text>
            </g>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText && item.TextPreview == "inherited transform payload");
    }

    [Fact]
    public void CssSpecificityUsesLexicographicTuplesInsteadOfFlattenedWeights() {
        string repeatedClasses = string.Concat(Enumerable.Repeat(".match", 101));
        byte[] svg = Svg($"<style>#target {{ display:inline }} {repeatedClasses} {{ display:none }}</style>" +
            "<text id='target' class='match' x='10' y='35'>id specificity visible</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "id specificity visible");
    }

    [Theory]
    [InlineData(".target, { display:none }")]
    [InlineData("> .target { display:none }")]
    [InlineData("g > > .target { display:none }")]
    [InlineData(".target { display:none")]
    public void MalformedOrRecoveredCssFailsClosed(string css) {
        byte[] svg = Svg($"<style>{css}</style><text class='target' x='10' y='35'>recovered css payload</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void ForeignNamespaceTextAndStylesRemainReportOnlyAndDoNotChangeNativeCascade() {
        byte[] svg = Svg("""
            <app:style xmlns:app="urn:example:extension">.native { display:none }</app:style>
            <app:text xmlns:app="urn:example:extension" style="display:none">extension payload</app:text>
            <text class="native" x="10" y="35">native visible</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);
        OfficeContentSafetyFinding extension = Assert.Single(report.Findings, item => item.TextPreview == "extension payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, extension.CleanupCapability);
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "native visible");
        Assert.Throws<InvalidOperationException>(() => OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { extension.Id })));
    }

    [Fact]
    public void UnsupportedAttributeSelectorOperatorsFailClosed() {
        byte[] svg = Svg("""
            <style>[data-state^="hide"] { display:none }</style>
            <text data-state="hidden" x="10" y="35">unsupported selector payload</text>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void UnsupportedAttributeSelectorFlagsFailClosed() {
        byte[] svg = Svg("""
            <style>[data-state="SHOW" i] { display:inline }</style>
            <text data-state="show" display="none" x="10" y="35">selector flag payload</text>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void UnsupportedCssWideRevertFailsClosed() {
        byte[] svg = Svg("""
            <text display="none" style="display:revert" x="10" y="35">reverted display payload</text>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void UnsupportedCssAtRulesFailClosed() {
        byte[] svg = Svg("""
            <style>@media all { .target { display:inline } }</style>
            <text class="target" display="none" x="10" y="35">at-rule payload</text>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CssSelectorListOverflowFailsClosedEvenWithinOneRule() {
        string selectors = string.Join(",", Enumerable.Range(0, 4096).Select(index => $"#unused-{index}")) + ",.target";
        byte[] svg = Svg($"<style>{selectors} {{ display:none }}</style>" +
            "<text class='target' x='10' y='35'>late selector payload</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CssInvalidAtComputedValueAndInitialOverridePresentationAttributes() {
        byte[] svg = Svg("""
            <text display="none" style="display:var(--missing)" x="10" y="25">missing display variable visible</text>
            <text opacity="0" style="opacity:initial" x="10" y="55">initial opacity visible</text>
            <text fill="none" style="fill:var(--missing)" x="10" y="85">missing fill variable visible</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview.Contains("variable visible", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "initial opacity visible");
    }

    [Fact]
    public void TransformsContributeToTheEffectiveTinyTextThreshold() {
        byte[] svg = Svg("""
            <text transform="scale(0.05)" font-size="16" x="10" y="35">transformed tiny payload</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText && item.TextPreview == "transformed tiny payload");
    }

    [Fact]
    public void AnisotropicTransformsDoNotMakeVisiblyTallTextTiny() {
        byte[] svg = Svg("""
            <text transform="scale(0.001,1)" font-size="16" x="10" y="35">anisotropic tiny payload</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText && item.TextPreview == "anisotropic tiny payload");
    }

    [Fact]
    public void StrongShearDoesNotCreateATinyTextFalsePositive() {
        byte[] svg = Svg("""
            <text transform="matrix(1,0,100,1,-1000,0)" font-size="16" x="10" y="20">strong shear remains ordinary size</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText && item.TextPreview == "strong shear remains ordinary size");
    }

    [Fact]
    public void ForeignNamespaceShapesAndDefinitionsDoNotAffectNativeVisualEvidence() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" width="220" height="120" viewBox="0 0 220 120">
              <defs><clipPath xmlns="" id="foreign-clip"><rect width="0" height="0" /></clipPath></defs>
              <text x="10" y="35" clip-path="url(#foreign-clip)">native visible through foreign clip</text>
              <text x="10" y="75">native visible under foreign shape</text>
              <rect xmlns="" x="0" y="48" width="220" height="35" fill="black" />
            </svg>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview.StartsWith("native visible", StringComparison.Ordinal));
    }

    [Fact]
    public void MixedNamespaceDefinitionChildrenAndForeignIdsFollowBrowserOwnership() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" xmlns:app="urn:example:extension" width="220" height="120" viewBox="0 0 220 120">
              <defs>
                <clipPath app:id="foreign-id"><rect width="0" height="0" /></clipPath>
                <clipPath id="native-empty"><rect xmlns="" width="220" height="120" /></clipPath>
                <linearGradient id="native-gradient"><stop xmlns="" offset="0" stop-color="black" /></linearGradient>
              </defs>
              <text clip-path="url(#foreign-id)" x="10" y="25">foreign id does not clip</text>
              <text clip-path="url(#native-empty)" x="10" y="55">foreign clip child payload</text>
              <text fill="url(#native-gradient)" x="10" y="85">foreign gradient stop payload</text>
            </svg>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "foreign id does not clip");
        Assert.Contains(report.Findings, item => item.TextPreview == "foreign clip child payload");
        Assert.Contains(report.Findings, item => item.TextPreview == "foreign gradient stop payload");
    }

    [Fact]
    public void StandardNamespaceChildrenUnderNamespaceLessRootRemainReportOnly() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg width="220" height="120" viewBox="0 0 220 120">
              <text xmlns="http://www.w3.org/2000/svg" style="display:none" x="10" y="35">foreign nested svg text</text>
            </svg>
            """);

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "foreign nested svg text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void VisualComparisonsUseTheSamePerRenderPixelBudgetAsTheBaseline() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" width="2000" height="1000" viewBox="0 0 2000 1000">
              <rect width="2000" height="1000" fill="white" />
              <text x="100" y="300" font-size="100" fill="white">large low contrast payload</text>
            </svg>
            """);
        var readerOptions = new OfficeSvgDrawingReaderOptions {
            MaximumContentSafetyVisualComparisons = 1,
            MaximumContentSafetyVisualPixels = 2_000_000
        };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.LowContrastText && item.TextPreview == "large low contrast payload");
    }

    [Fact]
    public void CssSelectorMatchingHasABoundedAggregateWorkBudget() {
        string rules = string.Join(string.Empty, Enumerable.Range(0, 1000).Select(index => $".unused-{index}{{fill:black}}"));
        string elements = string.Join(string.Empty, Enumerable.Range(0, 3000).Select(_ => "<g/>"));
        byte[] svg = Svg($"<style>{rules}</style>{elements}<text x='10' y='35'>ordinary visible</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void VisualDocumentTransformationWorkReducesComparisonCount() {
        string padding = new string('x', 2 * 1024 * 1024);
        byte[] svg = Svg($"<!--{padding}--><text x='10' y='35'>ordinary visible</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Diagnostics, item => item.Contains("document-transformation work budgets", StringComparison.Ordinal));
    }

    [Fact]
    public void CleanupRejectsUtf8RewriteThatExceedsTheSvgReaderHardLimit() {
        string retained = new string('界', 2_800_000);
        string xml = "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120'>" +
            "<metadata>" + retained + "</metadata>" +
            "<text style='display:none' x='10' y='35'>remove me</text></svg>";
        byte[] svg = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };
        OfficeContentSafetyFinding hidden = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "remove me");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.RemoveSelectedContent(
            svg,
            new OfficeContentCleanupSelection(new[] { hidden.Id }),
            readerOptions: readerOptions));
    }

    [Fact]
    public void NestedSvgViewportMapsVisibleTextIntoTheRootCanvas() {
        byte[] svg = Svg("""
            <svg x="10" y="10" width="100" height="50" viewBox="1000 0 100 50">
              <text x="1000" y="30">nested visible text</text>
            </svg>
            """);
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "nested visible text");
    }

    [Theory]
    [InlineData("media='print'")]
    [InlineData("type='text/less'")]
    [InlineData("title='alternate'")]
    public void ConditionalOrNonCssStylesheetsFailClosed(string attributes) {
        byte[] svg = Svg($"<style {attributes}>text {{ display:none }}</style><text x='10' y='35'>visible text</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void InheritedCustomPropertyExpansionHasACumulativeBudget() {
        string value = new string('x', 3000);
        string elements = string.Concat(Enumerable.Repeat("<g/>", 6000));
        byte[] svg = Svg($"<style>svg{{--payload:{value}}}g{{opacity:1}}</style>{elements}<text x='10' y='35'>visible text</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void RecursiveCustomPropertyExpansionFailsClosedAtTheComputedValueBudget() {
        string seed = new string('1', 550_000);
        byte[] svg = Svg($"""
            <svg style="--v0:{seed};--v1:var(--v0)var(--v0);--v2:var(--v1)var(--v1);--v3:var(--v2)var(--v2);--v4:var(--v3)var(--v3);--v5:var(--v4)var(--v4)">
              <text style="opacity:var(--v5)" x="10" y="35">visible text</text>
            </svg>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Theory]
    [InlineData("[data-state==hidden]")]
    [InlineData("[data-state=hidden=again]")]
    [InlineData("[data-state='hidden'junk]")]
    public void MalformedAttributeSelectorEqualityFailsClosed(string selector) {
        byte[] svg = Svg($"<style>{selector}{{display:none}}</style><text data-state='=hidden' x='10' y='35'>visible text</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void TextRunBudgetFailsClosedBeforeOmittingLaterCandidates() {
        string fragmented = string.Join("<!---->", Enumerable.Repeat("x", 4097));
        byte[] svg = Svg($"<text opacity='0' x='10' y='35'>{fragmented}</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void ExcessiveElementNestingFailsClosedBeforeRecursiveInspection() {
        string nested = string.Concat(Enumerable.Repeat("<g>", 130)) +
            "<text x='10' y='35'>nested text</text>" +
            string.Concat(Enumerable.Repeat("</g>", 130));
        byte[] svg = Svg(nested);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void HyperlinkWrappedTextParticipatesInConcealmentInspection() {
        byte[] svg = Svg("""
            <text x="10" y="35"><a href="https://example.test" style="display:none">linked payload</a></text>
            """);

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "linked payload");

        Assert.Equal(OfficeContentConcealmentKind.HiddenByProperty, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, finding.CleanupCapability);
    }

    [Fact]
    public void UseReferencedTextIsReportOnlyBecauseOneSourceCanHaveVisibleInstances() {
        byte[] svg = Svg("""
            <g id="shared"><text x="1000" y="35">shared instance</text></g>
            <use href="#shared" x="-990" />
            """);

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "shared instance");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("use or tref", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void SwitchBranchTextIsReportOnlyAcrossRendererContexts() {
        byte[] svg = Svg("""
            <switch>
              <text systemLanguage="fr" x="10" y="35">French branch</text>
              <text x="10" y="35">fallback branch</text>
            </switch>
            """);

        OfficeContentSafetyFinding[] findings = OfficeSvgDrawingReader.InspectContentSafety(svg).Findings
            .Where(item => item.TextPreview is "French branch" or "fallback branch")
            .ToArray();

        Assert.Equal(2, findings.Length);
        Assert.All(findings, item => Assert.Equal(OfficeContentCleanupCapability.ReportOnly, item.CleanupCapability));
    }

    [Theory]
    [InlineData("inherit")]
    [InlineData("unset")]
    public void InheritedCssWideCustomPropertyKeywordsRetainTheParentValue(string keyword) {
        byte[] svg = Svg($"<svg style='--o:0'><text style='--o:{keyword};opacity:var(--o)' x='10' y='35'>hidden custom value</text></svg>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText && item.TextPreview == "hidden custom value");
    }

    [Fact]
    public void InitialCustomPropertyUsesVarFallbackInsteadOfTheInheritedValue() {
        byte[] svg = Svg("<svg style='--o:0'><text style='--o:initial;opacity:var(--o,1)' x='10' y='35'>visible custom fallback</text></svg>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible custom fallback");
    }

    [Theory]
    [InlineData("opacity='-1'")]
    [InlineData("fill-opacity='-1'")]
    [InlineData("fill='none' stroke='black' stroke-opacity='-1'")]
    [InlineData("style='opacity:-1'")]
    public void OutOfRangeOpacityClampsBeforeStructuralInspection(string attributes) {
        byte[] svg = Svg($"<text {attributes} x='10' y='35'>clamped transparent text</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText && item.TextPreview == "clamped transparent text");
    }

    [Fact]
    public void ConditionalProcessingTextIsReportOnly() {
        byte[] svg = Svg("<text systemLanguage='fr' x='10' y='35'>locale text</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "locale text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void DynamicSvgMakesStaticTextCleanupReportOnly() {
        byte[] svg = Svg("""
            <text x="10" y="35">animated text<animate attributeName="opacity" values="0;1" dur="1s" /></text>
            """);

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "animated text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void ExcludingNonPrimaryContextStillReportsDynamicTextUnicodeAsReportOnly() {
        byte[] svg = Svg("""
            <text x="10" y="35">dynamic​text<animate attributeName="opacity" values="0;1" dur="1s" /></text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(
            svg,
            new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false });

        OfficeContentSafetyFinding finding = Assert.Single(
            report.Findings,
            item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void TrefReferencedSourceTextIsReportOnly() {
        byte[] svg = Svg("""
            <text id="source" x="1000" y="35">referenced text</text>
            <text x="10" y="35"><tref href="#source" /></text>
            """);

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "referenced text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void XmlStylesheetProcessingInstructionFailsClosed() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <?xml-stylesheet type="text/css" href="data:text/css,text%7Bdisplay:none%7D"?>
            <svg xmlns="http://www.w3.org/2000/svg" width="220" height="120">
              <text x="10" y="35">stylesheet text</text>
            </svg>
            """);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void IncompleteRendererProjectionMakesVisualCleanupEvidenceReportOnly() {
        byte[] svg = Svg("""
            <image href="https://example.test/background.png" width="220" height="120" />
            <text fill="white" fill-opacity="0.1" x="10" y="35">context dependent contrast</text>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        OfficeContentSafetyFinding finding = Assert.Single(
            report.Findings,
            item => item.TextPreview == "context dependent contrast");
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains(report.Diagnostics, item => item.Contains("visual comparison was limited", StringComparison.Ordinal));
    }

    [Fact]
    public void ExcessiveTextNodeFragmentationFailsClosed() {
        string fragmented = string.Join("<!---->", Enumerable.Repeat("x", OfficeSvgDrawingReaderOptions.DefaultMaximumElements + 1));
        byte[] svg = Svg($"<text x=\"10\" y=\"35\">{fragmented}</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CallerCanDisableVisualComparisonsWithoutDisablingStructuralInspection() {
        byte[] svg = Svg("""
            <text style="display:none" x="10" y="35">structural payload</text>
            """);
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.Contains(report.Findings, item => item.TextPreview == "structural payload");
        Assert.Contains(report.Diagnostics, item => item.Contains("disabled by the caller", StringComparison.Ordinal));
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
