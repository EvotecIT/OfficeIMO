using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    public static IEnumerable<object[]> NavigationFollowupCases() {
        foreach (PdfRedactionTextSelection selection in new[] { PdfRedactionTextSelection.MatchedGlyphs, PdfRedactionTextSelection.LogicalBlocks })
            foreach (bool hybrid in new[] { false, true })
                foreach (string variant in new[] { "NextPage", "PrevPage", "FirstPage", "LastPage", "Print", "StringCommand", "SyntaxCommand", "MetadataOnly", "MetadataBody", "MetadataImage", "MetadataPattern", "MetadataAppearance", "TaggedLink", "TaggedReply", "TaggedReplyChain", "UnknownReference", "UnknownSubtypeReference" })
                    yield return new object[] { selection, hybrid, variant };
    }

    [Theory]
    [MemberData(nameof(NavigationFollowupCases))]
    public void SearchPreservesNavigationAndRepairsOnlyKnownRemovedOwners(PdfRedactionTextSelection selection, bool hybrid, string variant) {
        byte[] input = BuildNavigationFollowupFixture(variant, hybrid);
        using PdfDocument source = PdfDocument.Load(input);
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral(variant == "SyntaxCommand" ? "NextPage" : "secret"));
        if (variant is "Print" or "StringCommand") {
            Assert.False(plan.IsReviewable);
            Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.Apply(new[] { new PdfRedactionArea(1, 70, 690, 150, 30) }));
            if (variant == "StringCommand") Assert.Contains(source.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions()).Issues, static issue => issue.Feature == "SearchResidue");
            return;
        }
        Assert.True(plan.IsReviewable);
        if (variant == "SyntaxCommand") Assert.Empty(plan.Areas);
        else Assert.NotEmpty(plan.Areas);
        if (variant is "UnknownReference" or "UnknownSubtypeReference") {
            PdfMutationBlockedException blocked = Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.ApplyWithEvidence(plan));
            Assert.Contains("FullRewrite.ProofFailed.RewritePreservation", blocked.Plan.BlockerCodes);
            Assert.Equal(input, source.ToBytes());
            return;
        }
        var result = source.Redactions.ApplyWithEvidence(plan, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None });
        Assert.True(result.Evidence.IsVerified);
        Assert.DoesNotContain("secret", Encoding.Latin1.GetString(result.Pdf), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret", DecodeSavedResidueBytes(result.Pdf), StringComparison.OrdinalIgnoreCase);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(result.Pdf);
        Assert.DoesNotContain("secret", string.Join("\n", independent.GetPages().Select(static page => page.Text)), StringComparison.OrdinalIgnoreCase);
        foreach (PdfExternalValidator validator in new[] { PdfExternalValidator.QpdfCheck(), PdfExternalValidator.QpdfDecoded(), PdfExternalValidator.PdfText() }) {
            if (!validator.IsAvailable) continue;
            PdfExternalProcessResult check = validator.Run(result.Pdf, "navigation-owners-redacted.pdf");
            Assert.Equal(0, check.ExitCode);
            Assert.DoesNotContain("secret", check.Output, StringComparison.OrdinalIgnoreCase);
        }
        PdfReadDocument saved = PdfReadDocument.Open(result.Pdf);
        if (variant.StartsWith("Metadata", StringComparison.Ordinal)) {
            Assert.Contains("keep metadata", DecodeSavedResidueBytes(result.Pdf));
            PdfDictionary custom = Dictionary(saved.CatalogDictionary!.Items["PrivateData"] is PdfArray array ? array.Items[0] : throw new InvalidOperationException());
            Assert.False(custom.Items.ContainsKey("Metadata"));
            Assert.False(Dictionary(custom.Items["Nested"]).Items.ContainsKey("Metadata"));

        } else if (variant.StartsWith("Tagged", StringComparison.Ordinal)) {
            PdfDictionary tree = Dictionary(saved.CatalogDictionary!.Items["StructTreeRoot"]);
            PdfDictionary document = Dictionary(Assert.Single(Assert.IsType<PdfArray>(tree.Items["K"]).Items));
            PdfDictionary paragraph = Dictionary(Assert.Single(Assert.IsType<PdfArray>(document.Items["K"]).Items));
            Assert.Equal("P", paragraph.Get<PdfName>("S")!.Name);
            Assert.Equal(0, paragraph.Get<PdfNumber>("K")!.Value);
            PdfDictionary parents = Dictionary(tree.Items["ParentTree"]);
            PdfArray nums = Assert.IsType<PdfArray>(parents.Items["Nums"]);
            Assert.Equal(2, nums.Items.Count);
            Assert.Equal(0, Assert.IsType<PdfNumber>(nums.Items[0]).Value);
            Assert.Same(paragraph, PdfObjectLookup.Resolve(saved.Objects, Assert.Single(Assert.IsType<PdfArray>(nums.Items[1]).Items)));
            PdfDictionary page = Assert.IsType<PdfDictionary>(saved.Objects[saved.Pages[0].ObjectNumber].Value);
            PdfDictionary reply = Dictionary(Assert.IsType<PdfArray>(page.Items["Annots"]).Items[0]);
            Assert.Equal("kept reply", reply.Get<PdfStringObj>("Contents")!.Value);
            Assert.False(reply.Items.ContainsKey("IRT"));
            Assert.False(reply.Items.ContainsKey("RT"));
            if (variant == "TaggedReplyChain") {
                Assert.Equal("Accepted", reply.Get<PdfStringObj>("State")!.Value);
                PdfDictionary child = Dictionary(Assert.IsType<PdfArray>(page.Items["Annots"]).Items[1]);
                Assert.Same(reply, PdfObjectLookup.Resolve(saved.Objects, child.Items["IRT"]));
                Assert.Equal("R", child.Get<PdfName>("RT")!.Name);
            }
        } else {
            PdfDictionary action = Dictionary(saved.CatalogDictionary!.Items["OpenAction"]);
            Assert.Equal(variant == "SyntaxCommand" ? "NextPage" : variant, action.Get<PdfName>("N")!.Name);
            Assert.Equal("Named", action.Get<PdfName>("S")!.Name);
        }
        // Area-based application uses the same ownership repair and action proof.
        if (plan.Areas.Count > 0) Assert.DoesNotContain("secret", DecodeSavedResidueBytes(source.Redactions.Apply(plan.Areas, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None }).ToBytes()), StringComparison.OrdinalIgnoreCase);
        PdfDictionary Dictionary(PdfObject value) => Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(saved.Objects, value));
    }

    private byte[] BuildNavigationFollowupFixture(string variant, bool hybrid) {
        return BuildNavigationRedactionFixture(variant is "MetadataOnly" or "StringCommand" or "SyntaxCommand" ? "NoMatch" : "BodyTitle", hybrid, objects => {
            // Keep the fixture independent of a document generator, using the existing classic/Word layout writers.
            string Stream(string dictionary, string data) => "<< " + dictionary + " /Length " + Encoding.ASCII.GetByteCount(data) + " >>\nstream\n" + data + "\nendstream";
            if (variant.StartsWith("Metadata", StringComparison.Ordinal)) {
                objects[1] = objects[1].TrimEnd()[..^2] + " /PrivateData [30 0 R 31 0 R] /StructTreeRoot 33 0 R /Kept << /Metadata 39 0 R >> >>";
                objects[30] = "<< /Metadata 38 0 R /Nested << /Metadata 38 0 R >> >>";
                objects[31] = Stream("/Metadata 38 0 R", "public stream");
                objects[33] = "<< /Type /StructTreeRoot /K [34 0 R] >>";
                objects[34] = "<< /Type /StructElem /S /Document /P 33 0 R /Metadata 38 0 R /K [] >>";
                objects[38] = Stream("/Type /Metadata /Subtype /XML /Filter /ASCIIHexDecode", Convert.ToHexString(Encoding.ASCII.GetBytes("<value>secret</value>")) + ">");
                objects[39] = Stream("/Type /Metadata /Subtype /XML", "<value>keep metadata</value>");
                if (variant == "MetadataImage") {
                    objects[6] = objects[6].Replace("/Resources <<", "/Resources << /XObject << /Im0 40 0 R >>", StringComparison.Ordinal);
                    objects[7] = Stream("", "BT /F1 12 Tf 72 700 Td (Public second) Tj ET q 50 0 0 50 200 400 cm /Im0 Do Q");
                    objects[40] = Stream("/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Metadata 38 0 R", "FF0000>");
                }
                if (variant == "MetadataPattern") {
                    objects[6] = objects[6].Replace("/Resources <<", "/Resources << /Pattern << /P1 41 0 R >>", StringComparison.Ordinal);
                    objects[7] = Stream("", "BT /F1 12 Tf 72 700 Td (Public second) Tj ET /Pattern cs /P1 scn 200 400 50 50 re f");
                    objects[41] = Stream("/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >> /Metadata 38 0 R", "1 0 0 rg 0 0 10 10 re f");
                }
                if (variant == "MetadataAppearance") {
                    objects[6] = objects[6].TrimEnd()[..^2] + " /Annots [40 0 R] >>";
                    objects[40] = "<< /Type /Annot /Subtype /Stamp /Rect [200 400 250 450] /Contents (keep appearance) /AP << /N 41 0 R >> >>";
                    objects[41] = Stream("/Type /XObject /Subtype /Form /BBox [0 0 50 50] /Resources << >> /Metadata 38 0 R", "1 0 0 rg 0 0 50 50 re f");
                }
            } else if (variant is "TaggedLink" or "TaggedReply" or "TaggedReplyChain" or "UnknownReference" or "UnknownSubtypeReference") {
                objects[1] = objects[1].TrimEnd()[..^2] + " /MarkInfo << /Marked true >> /StructTreeRoot 33 0 R >>";
                objects[3] = objects[3].TrimEnd()[..^2] + " /Annots [30 0 R 31 0 R 32 0 R] /StructParents 0 >>";
                objects[30] = "<< /Type /Annot /Subtype /Link /Rect [72 600 140 615] /StructParent 1 /Contents (secret) /A << /S /URI /URI (https://example.invalid/secret) >> /Popup 31 0 R >>";
                objects[31] = "<< /Type /Annot /Subtype /Popup /Rect [100 600 140 615] /Parent 30 0 R >>";
                objects[32] = "<< /Type /Annot /Subtype /Text /Rect [72 500 140 515] /Contents (kept reply) " + (variant is "TaggedReply" or "TaggedReplyChain" ? "/IRT 30 0 R /RT /R" : "") + " >>";
                objects[33] = "<< /Type /StructTreeRoot /K [34 0 R] /ParentTree 38 0 R /ParentTreeNextKey 2 >>";
                objects[34] = "<< /Type /StructElem /S /Document /P 33 0 R /K [35 0 R 36 0 R] " + (variant == "UnknownReference" ? "/PrivateRef 30 0 R" : "") + " >>";
                objects[35] = "<< /Type /StructElem /S /Link /P 34 0 R /Pg 3 0 R /K 37 0 R >>";
                objects[36] = "<< /Type /StructElem /S /P /P 34 0 R /Pg 3 0 R /K 0 >>";
                objects[37] = "<< /Type /OBJR /Pg 3 0 R /Obj 30 0 R >>";
                objects[38] = "<< /Nums [0 [36 0 R] 1 35 0 R] /Limits [0 1] >>";
                if (variant == "TaggedReplyChain") {
                    objects[3] = objects[3].Replace("32 0 R]", "32 0 R 39 0 R]", StringComparison.Ordinal);
                    objects[32] = objects[32].Replace("/RT /R", "/RT /R /State (Accepted) /StateModel (Review)", StringComparison.Ordinal);
                    objects[39] = "<< /Type /Annot /Subtype /Text /Rect [72 400 140 415] /Contents (kept child) /IRT 32 0 R /RT /R >>";
                }
                if (variant == "UnknownSubtypeReference") {
                    objects[1] = objects[1].TrimEnd()[..^2] + " /PrivateData 40 0 R >>";
                    objects[40] = "<< /Subtype /Private /IRT 30 0 R /RT /R /Popup 31 0 R >>";
                }
            } else {
                string action = "<< /S /Named /N " + (variant == "StringCommand" ? "(secret)" : "/" + (variant == "SyntaxCommand" ? "NextPage" : variant)) + " >>";
                objects[1] = objects[1].Replace("/OpenAction", "/KeptOpening", StringComparison.Ordinal).TrimEnd()[..^2] + " /OpenAction " + action + " >>";
                objects[22] = objects[22].Replace(" /Dest [6 0 R /Fit]", "", StringComparison.Ordinal).TrimEnd()[..^2] + " /A " + action + " >>";
                objects[3] = objects[3].TrimEnd()[..^2] + " /Annots [30 0 R] >>";
                objects[30] = "<< /Type /Annot /Subtype /Link /Rect [72 600 140 615] /A " + action + " >>";
                objects[23] = objects[23].Replace(" /Dest [6 0 R /XYZ null 700 1]", "", StringComparison.Ordinal).TrimEnd()[..^2] + " /A << /S /GoTo /D [6 0 R /Fit] /Next [" + action + "] >> >>";
            }
        });
    }
}
