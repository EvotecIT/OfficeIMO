using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAccessibilityCheckTests {
    private const string Xmp =
        "<?xpacket begin=\"\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?><x:xmpmeta xmlns:x=\"adobe:ns:meta/\">" +
        "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"><rdf:Description rdf:about=\"\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:pdfuaid=\"http://www.aiim.org/pdfua/ns/id/\">" +
        "<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">Accessibility fixture</rdf:li></rdf:Alt></dc:title>" +
        "<pdfuaid:part>1</pdfuaid:part></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end=\"w\"?>";

    public static TheoryData<string> Fixtures() => new() { "tagged-with-defects", "untagged" };

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void CheckAccessibility_FailsExactlyTheDefectsBuiltIntoTheFixture(string fixture) {
        (byte[] pdf, string[] expected) = CreateFixture(fixture);

        PdfAccessibilityReport report = PdfDocument.Load(new MemoryStream(pdf)).Proof.CheckAccessibility();

        Assert.Equal(expected.OrderBy(static id => id, StringComparer.Ordinal),
            report.FailedConditions.Select(static result => result.Condition.ConditionId).OrderBy(static id => id, StringComparer.Ordinal));
        Assert.Equal(PdfAccessibilityConditions.All.Select(static condition => condition.ConditionId),
            report.Conditions.Select(static result => result.Condition.ConditionId));
        foreach (PdfAccessibilityConditionResult result in report.Conditions.Where(static result => result.Condition.Coverage == PdfAccessibilityCoverage.ManualReview)) {
            Assert.Equal(PdfAccessibilityConditionStatus.NeedsReview, result.Status);
            Assert.Contains(report.Findings, finding => finding.ConditionId == result.Condition.ConditionId && finding.RequiresHumanReview);
        }
        AssertAgreesWithVeraPdf(pdf, report, fixture);
    }

    [Fact]
    public void CheckAccessibility_LocatesFindingsByPageObjectAndStructureElement() {
        (byte[] pdf, _) = CreateFixture("tagged-with-defects");

        PdfAccessibilityReport report = PdfDocument.Load(new MemoryStream(pdf)).Proof.CheckAccessibility();

        PdfAccessibilityFinding untagged = Assert.Single(report.Findings, static finding => finding.ConditionId == "01-005");
        Assert.Equal((1, 3, "7.1-3", PdfAccessibilityFindingSeverity.Error), (untagged.PageNumber, untagged.ObjectNumber, untagged.VeraPdfRule, untagged.Severity));
        PdfAccessibilityFinding figure = Assert.Single(report.Findings, static finding => finding.ConditionId == "13-004");
        Assert.Equal((9, "Figure", 1), (figure.StructureElementObjectNumber, figure.StructureType, figure.PageNumber));
        PdfAccessibilityFinding link = Assert.Single(report.Findings, static finding => finding.ConditionId == "28-012");
        Assert.Equal((1, 11), (link.PageNumber, link.ObjectNumber));
        PdfAccessibilityFinding font = Assert.Single(report.Findings, static finding => finding.ConditionId == "31-009");
        Assert.Equal(7, font.ObjectNumber);
        Assert.Equal(1, report.DeclaredPdfUaPart);
    }

    [Theory]
    [InlineData(PdfStandardPermissions.Print, true)]
    [InlineData(PdfStandardPermissions.Print | PdfStandardPermissions.Accessibility, false)]
    public void CheckAccessibility_ReportsWhatTheAccessibilityPermissionAllows(PdfStandardPermissions permissions, bool accessibilityDenied) {
        (byte[] plain, _) = CreateFixture("untagged");
        byte[] encrypted = PdfDocument.Load(new MemoryStream(plain)).Security
            .Encrypt(PdfStandardEncryptionOptions.ForPermissions("owner", permissions)).Pdf;

        PdfAccessibilityReport report = PdfDocument.Load(new MemoryStream(encrypted)).Proof.CheckAccessibility();

        PdfAccessibilityConditionStatus Status(string id) => report.Conditions.Single(result => result.Condition.ConditionId == id).Status;
        Assert.Equal(accessibilityDenied ? PdfAccessibilityConditionStatus.Failed : PdfAccessibilityConditionStatus.Passed, Status("26-002"));
        Assert.Equal(accessibilityDenied ? PdfAccessibilityConditionStatus.Indeterminate : PdfAccessibilityConditionStatus.Failed, Status("01-005"));
    }

    private static void AssertAgreesWithVeraPdf(byte[] pdf, PdfAccessibilityReport report, string fixture) {
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf("ua1");
        if (!validator.IsAvailable) {
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }
        PdfExternalProcessResult result = validator.Run(pdf, "accessibility-" + fixture + ".pdf");
        var veraPdfFailed = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match rule in Regex.Matches(result.Output, "<rule\\b[^>]*>")) {
            string text = rule.Value;
            string? Attribute(string name) {
                Match match = Regex.Match(text, "\\b" + name + "=\"([^\"]*)\"");
                return match.Success ? match.Groups[1].Value : null;
            }
            if (Attribute("status") == "failed" && Attribute("clause") is string clause && Attribute("testNumber") is string test) {
                veraPdfFailed.Add(clause + "-" + test);
            }
        }
        var covered = new HashSet<string>(PdfAccessibilityConditions.All
            .Where(static condition => condition.Coverage == PdfAccessibilityCoverage.Automatic)
            .SelectMany(static condition => condition.VeraPdfRules), StringComparer.Ordinal);
        var checkerFailed = new HashSet<string>(report.Findings
            .Where(static finding => finding.Severity == PdfAccessibilityFindingSeverity.Error && finding.VeraPdfRule != null)
            .Select(static finding => finding.VeraPdfRule!), StringComparer.Ordinal);
        Assert.Equal(veraPdfFailed.Where(covered.Contains).OrderBy(static rule => rule, StringComparer.Ordinal),
            checkerFailed.OrderBy(static rule => rule, StringComparer.Ordinal));
    }

    private static (byte[] Pdf, string[] ExpectedFailures) CreateFixture(string name) {
        string pageContent;
        switch (name) {
            case "tagged-with-defects":
                pageContent =
                    "/P <</MCID 0>> BDC BT /F1 12 Tf 10 150 Td (Tagged) Tj ET /Artifact BMC 0 0 1 1 re f EMC EMC\n" +
                    "/Figure <</MCID 1>> BDC 20 20 10 10 re f EMC\n" +
                    "BT /F1 12 Tf 10 100 Td (Untagged) Tj ET";
                return (Build(
                    "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R /Lang (en-US) " +
                    "/ViewerPreferences << /DisplayDocTitle true >> /Metadata 12 0 R >>",
                    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 7 0 R >> >> " +
                    "/Contents 4 0 R /Annots [11 0 R] /StructParents 0 >>",
                    Stream(string.Empty, pageContent),
                    "<< /Type /StructTreeRoot /K 6 0 R /RoleMap << /Custom /Bogus >> /ParentTree 13 0 R /ParentTreeNextKey 1 >>",
                    "<< /Type /StructElem /S /Document /P 5 0 R /K [8 0 R 9 0 R 10 0 R 14 0 R] >>",
                    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
                    "<< /Type /StructElem /S /P /P 6 0 R /Pg 3 0 R /K 0 >>",
                    "<< /Type /StructElem /S /Figure /P 6 0 R /Pg 3 0 R /K 1 >>",
                    "<< /Type /StructElem /S /Custom /P 6 0 R >>",
                    "<< /Type /Annot /Subtype /Link /Rect [10 10 50 30] /Border [0 0 0] /A << /S /URI /URI (https://example.com) >> >>",
                    Stream("/Type /Metadata /Subtype /XML", Xmp),
                    "<< /Nums [0 [8 0 R 9 0 R]] >>",
                    "<< /Type /StructElem /S /TR /P 6 0 R >>"),
                    new[] { "01-003", "01-004", "01-005", "02-001", "09-004", "13-004", "28-004", "28-008", "28-011", "28-012", "31-009" });
            case "untagged":
                pageContent = "BT /F1 12 Tf 10 150 Td (Plain) Tj ET 20 20 30 30 re f";
                return (Build(
                    "<< /Type /Catalog /Pages 2 0 R >>",
                    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
                    Stream(string.Empty, pageContent),
                    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
                    new[] { "UA1:6.2-1", "UA1:7.1-11", "01-005", "06-001", "06-002", "06-003", "07-001", "11-001", "31-009" });
            default:
                throw new ArgumentOutOfRangeException(nameof(name), name, null);
        }
    }

    [Theory]
    [InlineData("xfa-prefixed-dynamic", "25-001", PdfAccessibilityConditionStatus.Failed)]
    [InlineData("xfa-commented-dynamic", "25-001", PdfAccessibilityConditionStatus.Passed)]
    [InlineData("font-from-graphics-state", "31-009", PdfAccessibilityConditionStatus.Failed)]
    [InlineData("font-inherited-by-form", "31-009", PdfAccessibilityConditionStatus.Failed)]
    [InlineData("table-with-direct-content", "09-004", PdfAccessibilityConditionStatus.Failed)]
    [InlineData("enclosing-marked-content-language", "11-002", PdfAccessibilityConditionStatus.NotApplicable)]
    public void CheckAccessibility_ReadsTheConstructThatDecidesTheCondition(string construct, string conditionId, PdfAccessibilityConditionStatus expected) {
        string catalog = string.Empty, resources = string.Empty, content;
        var extra = new List<string>();
        switch (construct) {
            case "xfa-prefixed-dynamic":
            case "xfa-commented-dynamic":
                catalog = "/AcroForm << /Fields [] /XFA [(config) 5 0 R] >>";
                content = string.Empty;
                extra.Add(Stream(string.Empty, construct == "xfa-prefixed-dynamic"
                    ? "<c:config xmlns:c=\"http://www.xfa.org/schema/xci/3.0/\"><c:present><c:pdf><c:dynamicRender>required</c:dynamicRender></c:pdf></c:present></c:config>"
                    : "<config><!-- <dynamicRender>required</dynamicRender> --><present><pdf><dynamicRender>forbidden</dynamicRender></pdf></present></config>"));
                break;
            case "font-from-graphics-state":
                resources = "/ExtGState << /G1 << /Font [5 0 R 12] >> >>";
                content = "BT /G1 gs 10 10 Td (x) Tj ET";
                extra.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
                break;
            case "font-inherited-by-form":
                resources = "/Font << /F1 5 0 R >> /XObject << /X1 6 0 R >>";
                content = "BT /F1 12 Tf ET /X1 Do";
                extra.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
                extra.Add(Stream("/Type /XObject /Subtype /Form /BBox [0 0 100 100]", "BT 10 10 Td (x) Tj ET"));
                break;
            case "table-with-direct-content":
                catalog = "/MarkInfo << /Marked true >> /StructTreeRoot 5 0 R";
                content = "/Table <</MCID 0>> BDC 0 0 1 1 re f EMC";
                extra.Add("<< /Type /StructTreeRoot /K 6 0 R >>");
                extra.Add("<< /Type /StructElem /S /Table /P 5 0 R /Pg 3 0 R /K 0 >>");
                break;
            case "enclosing-marked-content-language":
                content = "/Span <</Lang (en-US)>> BDC /Span <</ActualText (x)>> BDC 0 0 1 1 re f EMC EMC";
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(construct), construct, null);
        }
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R " + catalog + " >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << " + resources + " >> /Contents 4 0 R >>",
            Stream(string.Empty, content)
        };
        objects.AddRange(extra);

        PdfAccessibilityReport report = PdfDocument.Load(new MemoryStream(Build(objects.ToArray()))).Proof.CheckAccessibility();

        Assert.Equal(expected, report.Conditions.Single(result => result.Condition.ConditionId == conditionId).Status);
    }

    private static string Stream(string dictionary, string data) =>
        "<< " + dictionary + " /Length " + Encoding.ASCII.GetByteCount(data).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + data + "\nendstream";

    private static byte[] Build(params string[] objects) {
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>(objects.Length);
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n")
            .Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
