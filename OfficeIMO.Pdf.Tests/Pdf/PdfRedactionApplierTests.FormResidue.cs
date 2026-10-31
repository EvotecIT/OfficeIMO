using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    public static IEnumerable<object[]> FormResidueCases() {
        foreach (PdfRedactionTextSelection selection in new[] { PdfRedactionTextSelection.MatchedGlyphs, PdfRedactionTextSelection.LogicalBlocks })
            foreach (string holder in new[] { "TU", "TM", "Contents", "Name", "CatalogXmp", "PageXmp", "Info", "UnknownString", "UnknownStream", "UnknownStringLive", "UnknownStreamLive", "AA", "JsStream", "NamesJavaScript", "Historical", "SyntaxNames" })
                yield return new object[] { selection, holder };
    }

    [Theory]
    [MemberData(nameof(FormResidueCases))]
    public void SearchClearsSavedFormResidueOrRefusesProof(PdfRedactionTextSelection selection, string holder) {
        byte[] bytes = BuildFormResidueSource(holder);
        PdfDocument source = PdfDocument.Load(bytes);
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
        bool blockedPlan = holder is "Name" or "Historical" or "UnknownString" or "UnknownStream";
        bool refusedOutput = holder is "UnknownStringLive" or "UnknownStreamLive" or "AA" or "JsStream" or "NamesJavaScript";
        if (blockedPlan || refusedOutput) {
            Assert.Equal(!blockedPlan, plan.IsReviewable);
            if (holder is "AA" or "JsStream" or "NamesJavaScript") {
                PdfMutationBlockedException rejected = Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.ApplyWithEvidence(plan));
                Assert.Contains("FullRewrite.ActiveContent", rejected.Plan.BlockerCodes);
                PdfDocument scriptsOnly = PdfDocument.Load(BuildFormResidueSource(holder, scriptOnly: true));
                PdfRedactionPlan scriptPlan = scriptsOnly.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
                Assert.False(scriptPlan.IsReviewable);
                Assert.Contains(scriptsOnly.Redactions.VerifyAppliedPlan(scriptPlan, new PdfRedactionVerificationOptions()).Issues,
                    static issue => issue.Feature == "SearchResidue");
            } else Assert.Throws<InvalidOperationException>(() => source.Redactions.ApplyWithEvidence(plan));
            Assert.False(source.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions()).IsVerified);
            if (plan.Areas.Count > 0) {
                if (holder is "AA" or "JsStream" or "NamesJavaScript") Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.Apply(plan.Areas));
                else Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan.Areas));
            }
            Assert.Equal(bytes, source.ToBytes());
            return;
        }
        Assert.True(plan.IsReviewable);
        if (holder is "TU" or "TM" or "Contents") {
            Assert.Equal(new[] { "Choice", "Text" }, plan.Matches.Where(static match => match.Kind == PdfRedactionMatchKind.FormField)
                .Select(static match => match.Text).OrderBy(static name => name).ToArray());
        }
        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None });
        Assert.True(result.Evidence.IsVerified);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(result.Pdf);
        Assert.Single(independent.GetPages());
        Assert.Equal(new[] { "Choice", "Text" }, result.ToDocument().Reader.FormFields().Select(static field => field.Name).OrderBy(static name => name).ToArray());
        if (holder == "SyntaxNames") return; // PDF names and keys legitimately retain this spelling.
        string decoded = DecodeSavedResidueBytes(result.Pdf);
        Assert.DoesNotContain("secret", decoded, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("007300650063007200650074", decoded, StringComparison.OrdinalIgnoreCase);
        PdfExternalValidator qpdf = PdfExternalValidator.QpdfDecoded();
        if (qpdf.IsAvailable) {
            PdfExternalProcessResult external = qpdf.Run(result.Pdf, "redacted-form.pdf");
            Assert.Equal(0, external.ExitCode);
            Assert.DoesNotContain("secret", external.Output, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("007300650063007200650074", external.Output, StringComparison.OrdinalIgnoreCase);
        }
    }

    private byte[] BuildFormResidueSource(string holder, bool scriptOnly = false) {
        string data = holder is "TU" or "TM" or "Contents" ? "/" + holder + " (the secret data)" : "";
        string name = holder == "Name" ? "secret-field" : "Text";
        string catalog = holder == "CatalogXmp" ? "/Metadata 14 0 R" : holder.StartsWith("Unknown", StringComparison.Ordinal) ? "/PrivateData 14 0 R" :
            holder == "NamesJavaScript" ? "/Names << /JavaScript << /Names [(OnOpen) << /S /JavaScript /JS 14 0 R >>] >> >>" :
            holder == "SyntaxNames" ? "/secretKey /secretName" : "";
        string value = !scriptOnly && (holder.EndsWith("Live", StringComparison.Ordinal) || holder is "AA" or "JsStream" or "NamesJavaScript" or "Historical") ? "secret" : "plain";
        string action = holder is "AA" or "JsStream" ? "/AA << /K << /S /JavaScript /JS " + (holder == "AA" ? "(event.value = 'secret';)" : "14 0 R") + " >> >>" : "";
        byte[] metadata = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'><title>secret</title></x:xmpmeta>");
        byte[] streamData = holder is "CatalogXmp" or "PageXmp" ? metadata : Encoding.UTF8.GetBytes("var value = 'secret';");
        if (holder == "UnknownStreamLive") streamData = Encoding.ASCII.GetBytes("/Private (se\\143ret) def");
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [10 0 R 11 0 R] >> " + catalog + " >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Annots [10 0 R 11 0 R] /Resources << /Font << /F1 4 0 R " + (holder == "SyntaxNames" ? "/secretFont 4 0 R " : "") + ">> >> " + (holder == "PageXmp" ? "/Metadata 14 0 R" : "") + " >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Public page) Tj ET")),
            TextWidget(10, "/FT /Tx /T (" + name + ") /V (" + value + ") " + data + action),
            TextWidget(11, "/FT /Ch /T (Choice) /Opt [(plain) (other)] /V (plain) " + data)
        };
        if (holder == "Info") objects.Add("14 0 obj\n<< /Producer (secret) /Custom (secret) >>\nendobj");
        else if (holder.StartsWith("UnknownString", StringComparison.Ordinal)) objects.Add("14 0 obj\n[(public) <FEFF007300650063007200650074>]\nendobj");
        else if (holder is "CatalogXmp" or "PageXmp" or "UnknownStream" or "UnknownStreamLive" or "JsStream" or "NamesJavaScript") objects.Add(BuildStreamObject(14, streamData));
        byte[] result = BuildPdf(objects.ToArray(), rootObjectNumber: 1);
        if (holder == "Info") {
            string raw = Encoding.ASCII.GetString(result).Replace("/Root 1 0 R", "/Root 1 0 R /Info 14 0 R");
            result = Encoding.ASCII.GetBytes(raw); // trailer size changes do not move any object or xref offset.
        }
        if (holder != "Historical") return result;
        // An obsolete definition contains the only match; the selected latest definition is clean.
        string original = Encoding.ASCII.GetString(result);
        result = Encoding.ASCII.GetBytes(original);
        int prev = int.Parse(original.Substring(original.LastIndexOf("startxref", StringComparison.Ordinal) + 9).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)[0], CultureInfo.InvariantCulture);
        string update = TextWidget(10, "/FT /Tx /T (Text) /V (plain)") + "\n" + TextWidget(11, "/FT /Ch /T (Choice) /V (plain)") + "\n";
        int second = update.IndexOf("11 0 obj", StringComparison.Ordinal);
        update += "xref\n10 2\n" + result.Length.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n" +
            (result.Length + second).ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \ntrailer\n<< /Size 15 /Root 1 0 R /Prev " + prev.ToString(CultureInfo.InvariantCulture) +
            " >>\nstartxref\n" + (result.Length + Encoding.ASCII.GetByteCount(update)).ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n";
        return result.Concat(Encoding.ASCII.GetBytes(update)).ToArray();
    }

    // Deliberately independent of OfficeIMO's object parser and stream decoder.
    private static string DecodeSavedResidueBytes(byte[] pdf) {
        string raw = Encoding.GetEncoding(28591).GetString(pdf);
        var decoded = new StringBuilder(raw);
        foreach (Match match in Regex.Matches(raw, @"(?s)(<<[^>]*>>)[\r\n ]*stream\r?\n(.*?)\r?\nendstream")) {
            if (!match.Groups[1].Value.Contains("/FlateDecode", StringComparison.Ordinal)) continue;
            byte[] compressed = Encoding.GetEncoding(28591).GetBytes(match.Groups[2].Value);
            using var input = new MemoryStream(compressed, 2, compressed.Length - 6);
            using var deflate = new DeflateStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            deflate.CopyTo(output);
            decoded.Append(Encoding.GetEncoding(28591).GetString(output.ToArray()));
        }
        return decoded.ToString();
    }
}
