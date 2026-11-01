using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    public static IEnumerable<object[]> NavigationRedactionCases() {
        foreach (PdfRedactionTextSelection selection in new[] { PdfRedactionTextSelection.MatchedGlyphs, PdfRedactionTextSelection.LogicalBlocks })
            foreach (bool hybrid in new[] { false, true })
                foreach (string variant in new[] { "Title", "BodyTitle", "Hex", "Utf16", "PdfDoc", "NoMatch", "Array", "SafeNext", "SafeAA", "NamedSafe", "CaseSensitive", "CaseBody", "Regex", "PresentationNA", "PresentationPA", "SharedOutlineLink", "FormScriptOpening", "FormScriptOutline", "FormScriptBoth",
                    "LegacyDestination", "NameTreeDestination", "OutlineDestination", "GoToDestination", "JavaScript", "Launch", "SubmitForm", "ImportData", "GoToR", "GoToE", "Rendition", "URI", "Unknown", "Next", "OutlineAction", "CatalogAA", "PageAA" })
                    yield return new object[] { selection, hybrid, variant };
    }

    [Theory]
    [MemberData(nameof(NavigationRedactionCases))]
    public void Search_ClearsBookmarkTitlesAndPreservesOnlySafeNavigation(PdfRedactionTextSelection selection, bool hybrid, string variant) {
        byte[] input = BuildNavigationRedactionFixture(variant, hybrid);
        using PdfDocument source = PdfDocument.Load(input);
        string query = variant == "PdfDoc" ? "secret•" : "secret";
        PdfRedactionPlan plan = source.Redactions.Search(variant == "Regex" ? new PdfRedactionSearchOptions { TextSelection = selection }.AddRegex("se(?:cret)") : new PdfRedactionSearchOptions { TextSelection = selection, MatchCase = variant is "CaseSensitive" or "CaseBody" }.AddLiteral(query));
        if (variant.StartsWith("FormScript", StringComparison.Ordinal)) {
            Assert.True(plan.IsReviewable);
            PdfMutationBlockedException error = Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.ApplyWithEvidence(plan));
            Assert.Contains("FullRewrite.ActiveContent", error.Plan.BlockerCodes);
            return;
        }
        bool refused = variant is "LegacyDestination" or "NameTreeDestination" or "OutlineDestination" or "GoToDestination" or
            "JavaScript" or "Launch" or "SubmitForm" or "ImportData" or "GoToR" or "GoToE" or "Rendition" or "URI" or "Unknown" or "Next" or "OutlineAction" or "CatalogAA" or "PageAA" or "PresentationNA" or "PresentationPA" or "SharedOutlineLink";
        if (refused) {
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, static finding => finding.Severity == PdfDiagnosticSeverity.Error &&
                finding.Code is "redaction-navigation-unsupported" or "redaction-search-destination-unsupported");
            if (variant is not "LegacyDestination" and not "NameTreeDestination" and not "OutlineDestination" and not "GoToDestination")
                Assert.False(source.PlanMutation(PdfMutationOperation.Redact).CanExecute);
            Exception? error = Record.Exception(() => source.Redactions.ApplyWithEvidence(plan));
            if (variant is "LegacyDestination" or "NameTreeDestination" or "OutlineDestination" or "GoToDestination") Assert.IsType<InvalidOperationException>(error);
            else {
                Assert.IsType<PdfMutationBlockedException>(error);
                Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.Apply(new[] { new PdfRedactionArea(1, 60, 680, 150, 40) }));
            }
            return;
        }
        Assert.True(plan.IsReviewable);
        Assert.True(source.PlanMutation(PdfMutationOperation.Redact).CanExecute);
        if (variant is not "NoMatch" and not "CaseSensitive" and not "CaseBody") {
            Assert.True(plan.HasMatches);
            Assert.NotEmpty(plan.Areas);
            Assert.Contains(plan.Findings, static finding => finding.Code == "redaction-search-outline-titles");
            Assert.DoesNotContain(plan.Areas, area => area.Label?.Contains(query, StringComparison.OrdinalIgnoreCase) == true);
        }
        PdfRedactionApplyResult applied = source.Redactions.ApplyWithEvidence(plan, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None });
        Assert.True(applied.Evidence.IsVerified);
        if (variant is not "NoMatch" and not "CaseSensitive" and not "CaseBody")
            Assert.Contains(applied.Evidence.Items, static item => item.ReviewedMatch.Kind == PdfRedactionMatchKind.BookmarkTitle && item.ReviewedMatch.Text is null && item.Status == PdfRedactionEvidenceStatus.VerifiedAbsent);
        Assert.Empty(applied.Evidence.ResidualMatches);
        byte[] output = applied.Pdf;
        Assert.DoesNotContain(query, Encoding.GetEncoding(28591).GetString(output), (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(query, DecodeSavedResidueBytes(output), (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("0073006500630072", DecodeSavedResidueBytes(output), (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.DoesNotContain("secret", string.Join("\n", independent.GetPages().Select(static page => page.Text)), (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        AssertNavigationTree(output, variant == "NoMatch" ? "Middle public" : variant is "CaseSensitive" or "CaseBody" ? "Secret" : "");
        using PdfDocument saved = PdfDocument.Load(output);
        if (variant == "NamedSafe") Assert.Equal(2, Assert.Single(saved.Reader.NamedDestinations()).PageNumber);
        else Assert.Equal(2, saved.Reader.OpenAction()!.PageNumber);
        foreach (PdfExternalValidator validator in new[] { PdfExternalValidator.QpdfCheck(), PdfExternalValidator.QpdfDecoded(), PdfExternalValidator.PdfText() }) {
            if (!validator.IsAvailable) continue;
            PdfExternalProcessResult check = validator.Run(output, "navigation-redacted.pdf");
            Assert.Equal(0, check.ExitCode);
            Assert.DoesNotContain("secret", check.Output, (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain("0073006500630072", check.Output, (variant is "CaseSensitive" or "CaseBody" ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        }
        if (plan.Areas.Count > 0) AssertNavigationTree(source.Redactions.Apply(plan.Areas).ToBytes(), variant == "CaseBody" ? "Secret" : "");
    }

    private byte[] BuildNavigationRedactionFixture(string variant, bool hybrid, Action<Dictionary<int, string>>? customize = null) {
        string middle = variant switch {
            "Hex" => "<736563726574>", "Utf16" => "<FEFF007300650063007200650074>",
            "PdfDoc" => @"(secret\200)", "CaseSensitive" or "CaseBody" => "(Secret)", "NoMatch" => "(Middle public)", _ => @"(se\143ret)"
        };
        string action = "<< /S /GoTo /D [6 0 R /FitH 700] >>";
        if (variant == "Array") action = "[6 0 R /FitH 700]";
        if (variant == "SafeNext") action = "<< /S /GoTo /D [6 0 R /FitH 700] /Next [<< /S /GoTo /D [3 0 R /Fit] >>] >>";
        if (variant is "JavaScript" or "Launch" or "SubmitForm" or "ImportData" or "GoToR" or "GoToE" or "Rendition" or "URI" or "Unknown") action = UnsafeAction(variant);
        if (variant == "Next") action = "<< /S /GoTo /D [6 0 R /Fit] /Next [" + UnsafeAction("JavaScript") + "] >>";
        string extra = variant == "SafeAA" ? "/AA << /WC << /S /GoTo /D [6 0 R /Fit] >> >>" : variant == "CatalogAA" ? "/AA << /WC " + UnsafeAction("JavaScript") + " >>" : "";
        string dest = variant is "LegacyDestination" or "OutlineDestination" or "GoToDestination" ? "/Dests << /secret [6 0 R /Fit] >>" : variant == "NamedSafe" ? "/Dests << /Kept [6 0 R /Fit] >>" :
            variant == "NameTreeDestination" ? "/Names << /Dests << /Names [(secret) [6 0 R /Fit]] >> >>" : "";
        if (variant == "GoToDestination") action = "<< /S /GoTo /D /secret >>";
        if (variant == "NamedSafe") action = "<< /S /GoTo /D /Kept >>";
        string presentation = variant == "PresentationNA" ? "/PresSteps << /Type /NavNode /NA << /S /URI /URI (https://example.invalid/) >> >>" :
            variant == "PresentationPA" ? "/PresSteps << /PA << /S /URI /URI (https://example.invalid/) >> >>" : "";
        var objects = new Dictionary<int, string> {
            [1] = "<< /Type /Catalog /Pages 2 0 R /Outlines 20 0 R /OpenAction " + action + " " + extra + " " + dest + " >>",
            [2] = "<< /Type /Pages /Count 2 /Kids [3 0 R 6 0 R] /MediaBox [0 0 612 792] >>",
            [3] = "<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> " + (variant == "PageAA" ? "/AA << /O " + UnsafeAction("JavaScript") + " >>" : "") + " " + presentation + " >>",
            [4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            [5] = Stream("BT /F1 12 Tf 72 700 Td (" + (variant is "BodyTitle" or "CaseBody" ? "secret body" : "Public first") + ") Tj ET"),
            [6] = "<< /Type /Page /Parent 2 0 R /Contents 7 0 R /Resources << /Font << /F1 4 0 R >> >> >>",
            [7] = Stream("BT /F1 12 Tf 72 700 Td (Public second) Tj ET"),
            [20] = "<< /Type /Outlines /First 21 0 R /Last 23 0 R /Count 5 >>",
            [21] = "<< /Title (Keep first) /Parent 20 0 R /Next 22 0 R /Dest [3 0 R /Fit] >>",
            [22] = "<< /Title " + middle + " /Parent 20 0 R /Prev 21 0 R /Next 23 0 R /First 24 0 R /Last 25 0 R /Count 2 " +
                (variant == "OutlineAction" ? "/A " + UnsafeAction("JavaScript") : "/Dest " + (variant == "OutlineDestination" ? "(secret)" : "[6 0 R /Fit]")) + " >>",
            [23] = "<< /Title (Keep last) /Parent 20 0 R /Prev 22 0 R /Dest [6 0 R /XYZ null 700 1] >>",
            [24] = "<< /Title (Keep child) /Parent 22 0 R /Next 25 0 R /Dest [3 0 R /Fit] >>",
            [25] = "<< /Title (Keep nested) /Parent 22 0 R /Prev 24 0 R /Dest [6 0 R /Fit] >>"
        };
        if (variant == "SharedOutlineLink") {
            objects[3] = objects[3].Replace("/Type /Page", "/Type /Page /Annots [22 0 R]");
            objects[22] = objects[22].Replace("/Title", "/Subtype /Link /Rect [300 300 400 320] /A 32 0 R /Title");
            objects[22] = objects[22].Replace(" /Dest [6 0 R /Fit]", "");
            objects[32] = "<< /Type /Action /S /URI /URI (https://example.invalid/) >>";
        }
        if (variant.StartsWith("FormScript", StringComparison.Ordinal)) {
            objects[1] = objects[1].Replace("/Type /Catalog", "/Type /Catalog /AcroForm << /Fields [30 0 R] >>");
            objects[3] = objects[3].Replace("/Type /Page", "/Type /Page /Annots [31 0 R]");
            objects[30] = "<< /FT /Tx /T (Name) /V (secret) /Kids [31 0 R] /AA << /K << /S /JavaScript /JS (public) >> >> >>";
            objects[31] = "<< /Type /Annot /Subtype /Widget /Parent 30 0 R /Rect [100 500 300 520] >>";
            if (variant == "FormScriptOpening") objects[1] = objects[1].Replace(" /Outlines 20 0 R", "");
            if (variant == "FormScriptOutline") objects[1] = objects[1].Replace(" /OpenAction " + action, "");
        }
        customize?.Invoke(objects);
        return hybrid ? BuildWordHybridRedactionFixture(false, objects) : BuildPdf(objects.Select(pair => pair.Key.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + pair.Value + "\nendobj").ToArray(), 1);
        static string Stream(string content) => "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream";
        static string UnsafeAction(string type) => "<< /S /" + type + (type == "JavaScript" ? " /JS (alert)" : type == "URI" ? " /URI (https://example.invalid/)" : " /F (public)") + " >>";
    }

    private static void AssertNavigationTree(byte[] output, string middleTitle) {
        PdfReadDocument document = PdfReadDocument.Open(output);
        PdfDictionary root = Resolve(document.CatalogDictionary!.Items["Outlines"]);
        PdfDictionary first = Resolve(root.Items["First"]), middle = Resolve(first.Items["Next"]), last = Resolve(middle.Items["Next"]);
        Assert.Same(last, Resolve(root.Items["Last"])); Assert.Equal(5, root.Get<PdfNumber>("Count")!.Value);
        Assert.Equal("Keep first", first.Get<PdfStringObj>("Title")!.Value); Assert.Equal(middleTitle, middle.Get<PdfStringObj>("Title")!.Value); Assert.Equal("Keep last", last.Get<PdfStringObj>("Title")!.Value);
        Assert.Same(first, Resolve(middle.Items["Prev"])); Assert.Same(middle, Resolve(last.Items["Prev"]));
        foreach (PdfDictionary item in new[] { first, middle, last }) Assert.Same(root, Resolve(item.Items["Parent"]));
        PdfDictionary child = Resolve(middle.Items["First"]), nested = Resolve(child.Items["Next"]);
        Assert.Same(nested, Resolve(middle.Items["Last"])); Assert.Same(child, Resolve(nested.Items["Prev"]));
        Assert.Same(middle, Resolve(child.Items["Parent"])); Assert.Same(middle, Resolve(nested.Items["Parent"])); Assert.Equal(2, middle.Get<PdfNumber>("Count")!.Value);
        Assert.Equal(new[] { 1, 2, 2, 1, 2 }, new[] { first, middle, last, child, nested }.Select(item => {
            PdfArray destination = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(document.Objects, item.Items["Dest"]));
            return document.Pages.ToList().FindIndex(page => page.ObjectNumber == Assert.IsType<PdfReference>(destination.Items[0]).ObjectNumber) + 1;
        }));
        PdfDictionary Resolve(PdfObject value) => Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(document.Objects, value));
    }
}
