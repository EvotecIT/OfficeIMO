using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    public static IEnumerable<object[]> CompressedRedactionCases() {
        foreach (PdfRedactionTextSelection selection in new[] { PdfRedactionTextSelection.MatchedGlyphs, PdfRedactionTextSelection.LogicalBlocks })
            foreach (string variant in new[] { "TextChoice", "EscapedTypes", "HarmlessNames", "EscapedTypesDuplicate", "CompactStreamDelimiter", "XrefOnly", "NoMatch", "NamesOnly", "DuplicateContainer", "ExcludedPage", "UnindexedSecret", "ByteRange", "Sig", "SigFlags", "DocMDP", "UR3", "XFA", "AA", "NamesJavaScript", "OpenAction", "Name", "Extends", "WinningRevision", "TruncatedClassicTable", "MalformedHistoryTrailer", "MissingCompressedPredecessor" })
                yield return new object[] { selection, variant };
    }

    [Theory]
    [MemberData(nameof(CompressedRedactionCases))]
    public void SearchNormalizesProvenCompressedObjectsAndRefusesUnprovenSources(PdfRedactionTextSelection selection, string variant) {
        byte[] input = BuildCompressedRedactionFixture(variant);
        PdfDocument source = PdfDocument.Load(input);
        if (variant == "Extends") Assert.Equal("plain", source.Reader.FormFields().Single(static field => field.Name == "Text").Value);
        var options = new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret");
        if (variant == "ExcludedPage") options.PageNumbers.Add(1);
        PdfRedactionPlan plan = source.Redactions.Search(options);
        bool accepted = variant is "TextChoice" or "EscapedTypes" or "HarmlessNames" or "CompactStreamDelimiter" or "XrefOnly" or "NoMatch" or "NamesOnly" or "ExcludedPage" or "Extends" or "WinningRevision";
        if (!accepted) {
            if (variant is "ByteRange" or "Sig" or "SigFlags" or "DocMDP" or "UR3" or "AA" or "NamesJavaScript" or "OpenAction") {
                PdfMutationPlan gate = PdfMutationPlanner.Plan(input, PdfMutationOperation.Redact);
                Assert.False(gate.CanExecute);
                if (variant == "ByteRange") Assert.True(gate.Preflight.Probe.Security.HasByteRange);
                if (variant is "Sig" or "SigFlags") Assert.True(gate.Preflight.Probe.Security.HasSignatures);
                if (variant == "DocMDP") Assert.True(gate.Preflight.Probe.Security.HasDocMDPPermissions);
                if (variant == "UR3") Assert.True(gate.Preflight.Probe.Security.HasUsageRights);
                if (variant is "AA" or "NamesJavaScript" or "OpenAction") Assert.Contains("FullRewrite.ActiveContent", gate.BlockerCodes);
            }
            Assert.NotNull(Record.Exception(() => source.Redactions.ApplyWithEvidence(plan)));
            Assert.NotNull(Record.Exception(() => source.Redactions.Apply(plan)));
            if (plan.Areas.Count > 0) Assert.NotNull(Record.Exception(() => source.Redactions.Apply(plan.Areas)));
            Assert.Equal(input, source.ToBytes());
            return;
        }
        Assert.True(plan.IsReviewable);
        PdfMutationPlan mutation = PdfMutationPlanner.Plan(input, PdfMutationOperation.Redact);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, mutation.ExecutionMode);
        Assert.Contains(PdfMutationProof.RewritePreservation, mutation.RequiredProofs);
        Assert.Contains(PdfMutationProof.RedactionResidue, mutation.RequiredProofs);
        Assert.Contains(PdfMutationProof.FormFieldReadback, mutation.RequiredProofs);
        Assert.Contains("Input.CrossReferenceAndObjectStreamsWillBeNormalized", mutation.Warnings);
        if (variant is "TextChoice" or "XrefOnly") Assert.Equal(new[] { "Choice", "Text" }, plan.Matches
            .Where(static match => match.Kind == PdfRedactionMatchKind.FormField).Select(static match => match.Text).OrderBy(static name => name));
        PdfRedactionApplyResult applied = source.Redactions.ApplyWithEvidence(plan, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None });
        Assert.True(applied.Evidence.IsVerified);
        byte[] output = applied.Pdf;
        string decoded = DecodeSavedResidueBytes(output);
        Assert.DoesNotMatch(@"/Type\s*/(?:ObjStm|XRef)\b", decoded);
        Assert.DoesNotContain("/Prev", decoded, StringComparison.Ordinal);
        using var reader = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(variant == "ExcludedPage" ? 2 : 1, reader.NumberOfPages);
        if (variant == "ExcludedPage") Assert.Contains("secret", reader.GetPage(2).Text, StringComparison.Ordinal);
        else Assert.DoesNotContain("secret", variant == "NamesOnly" ? Regex.Replace(decoded, @"/[^\s()<>\[\]{}/%]+", "") : decoded, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(new[] { "Choice", "Text" }, applied.ToDocument().Reader.FormFields().Select(static field => field.Name).OrderBy(static name => name));
        PdfExternalValidator validator = PdfExternalValidator.QpdfCheck();
        if (validator.IsAvailable) Assert.Equal(0, validator.Run(output, "classic-redaction.pdf").ExitCode);
        PdfExternalValidator qdf = PdfExternalValidator.QpdfDecoded();
        if (qdf.IsAvailable && variant is not "ExcludedPage" and not "NamesOnly") Assert.DoesNotContain("secret", qdf.Run(output, "decoded-redaction.pdf").Output, StringComparison.OrdinalIgnoreCase);
        if (plan.Areas.Count > 0) {
            byte[] areas = source.Redactions.Apply(plan.Areas).ToBytes();
            Assert.DoesNotContain("secret", DecodeSavedResidueBytes(areas), StringComparison.OrdinalIgnoreCase);
        }
    }

    private byte[] BuildCompressedRedactionFixture(string variant) {
        string holder = variant == "Name" ? "Name" : variant == "ExcludedPage" ? "ExcludedPage" : variant is "NoMatch" or "NamesOnly" or "HarmlessNames" or "DuplicateContainer" or "EscapedTypesDuplicate" or "UnindexedSecret" or "MalformedHistoryTrailer" ? "SyntaxNames" : "TU";
        string raw = Encoding.ASCII.GetString(BuildFormResidueSource(holder));
        // These names are harmless; keep the empty-search fixture free of the requested spelling.
        if (variant is "NoMatch" or "HarmlessNames" or "DuplicateContainer" or "EscapedTypesDuplicate" or "UnindexedSecret" or "MalformedHistoryTrailer") raw = raw.Replace("secret", "public");
        string catalog = variant switch {
            "DocMDP" => "/Perms << /DocMDP 20 0 R >>",
            "UR3" => "/Perms << /UR3 20 0 R >>",
            "AA" => "/AA << /WC << /S /JavaScript /JS (secret) >> >>",
            "OpenAction" => "/OpenAction << /S /JavaScript /JS (secret) >>",
            "NamesJavaScript" => "/Names << /JavaScript << /Names [(init) << /S /JavaScript /JS (secret) >>] >> >>",
            _ => ""
        };
        raw = raw.Replace("/Type /Catalog", "/Type /Catalog " + catalog);
        if (variant == "HarmlessNames") raw = raw.Replace("/F1 4 0 R", "/F1 4 0 R /ObjStm 4 0 R /XRef 4 0 R");
        if (variant == "SigFlags") raw = raw.Replace("/AcroForm <<", "/AcroForm << /SigFlags 3");
        if (variant == "XFA") raw = raw.Replace("/AcroForm << /Fields [10 0 R 11 0 R] >>", "/AcroForm << /XFA (dynamic secret) >>")
            .Replace("/Annots [10 0 R 11 0 R]", "/Annots []");
        var objects = Regex.Matches(raw, @"(?s)(\d+) 0 obj\s*\n(.*?)\nendobj").Cast<Match>()
            .ToDictionary(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), static match => match.Groups[2].Value);
        if (variant == "XFA") { objects.Remove(10); objects.Remove(11); }
        if (variant == "CompactStreamDelimiter") objects[5] = objects[5].Replace("\nendstream", "endstream");
        if (variant == "ByteRange") objects.Add(20, "<< /ByteRange [0 2 4 2] >>");
        if (variant is "Sig" or "DocMDP" or "UR3") objects.Add(20, "<< /Type /Sig /Contents <0000> >>");
        if (variant == "UnindexedSecret") objects.Add(20, "<< /V (secret) >>");
        if (variant == "Extends") {
            string previous = "10 0 << /FT /Tx /T (Text) /V (older extended value) >>";
            string hex = string.Concat(Encoding.ASCII.GetBytes(previous).Select(static value => value.ToString("X2", CultureInfo.InvariantCulture))) + ">";
            objects.Add(99, "<< /Type /ObjStm /N 1 /First 5 /Filter /ASCIIHexDecode /Length " + hex.Length + " >>\nstream\n" + hex + "\nendstream");
        }
        int container = objects.Keys.Max() + 1, xref = container + 1;
        int[] packed = variant is "XrefOnly" or "TruncatedClassicTable" or "MalformedHistoryTrailer" ? Array.Empty<int>() : objects.Keys.Where(number => !objects[number].Contains("\nstream\n", StringComparison.Ordinal)).ToArray();
        var header = new StringBuilder();
        var body = new StringBuilder();
        foreach (int number in packed) { header.Append(number).Append(' ').Append(body.Length).Append(' '); body.Append(objects[number]).Append('\n'); }
        if (packed.Length > 0) {
            byte[] payload = Encoding.ASCII.GetBytes(header.ToString() + body);
            using var compressed = new MemoryStream();
            // Raw Deflate plus an ASCIIHex filter works on all supported test targets.
            using (var deflate = new DeflateStream(compressed, CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(payload, 0, payload.Length);
            string hex = string.Concat(compressed.ToArray().Select(static value => value.ToString("X2", CultureInfo.InvariantCulture))) + ">";
            objects.Add(container, "<< /Type /ObjStm /N " + packed.Length + " /First " + header.Length +
                " /Filter [/ASCIIHexDecode /FlateDecode] /Length " + hex.Length + (variant == "Extends" ? " /Extends 99 0 R" : "") + " >>\nstream\n" + hex + "\nendstream");
        }
        var output = new StringBuilder("%PDF-1.7\n");
        var offsets = new Dictionary<int, int>();
        foreach (var entry in objects.Where(entry => !packed.Contains(entry.Key))) {
            if (variant is "DuplicateContainer" or "EscapedTypesDuplicate" && entry.Key == container) {
                string oldPayload = "99 0 << /V (secret) >>";
                string oldHex = string.Concat(Encoding.ASCII.GetBytes(oldPayload).Select(static value => value.ToString("X2", CultureInfo.InvariantCulture))) + ">";
                output.Append(container).Append(variant == "EscapedTypesDuplicate" ? " 0 obj\n<< /Type /Obj#53tm /N 1 /First 5 /Filter /ASCIIHexDecode /Length " : " 0 obj\n<< /Type /ObjStm /N 1 /First 5 /Filter /ASCIIHexDecode /Length ").Append(oldHex.Length)
                    .Append(" >>\nstream\n").Append(oldHex).Append("\nendstream\nendobj\n");
            }
            offsets[entry.Key] = output.Length;
            output.Append(entry.Key).Append(" 0 obj\n").Append(variant.StartsWith("EscapedTypes", StringComparison.Ordinal) ? entry.Value.Replace("/Type /ObjStm", "/Type /Obj#53tm") : entry.Value).Append("\nendobj\n");
        }
        offsets[xref] = output.Length;
        var rows = new StringBuilder();
        for (int number = 0; number <= xref; number++) {
            int index = Array.IndexOf(packed, number);
            int type = index >= 0 && !(variant == "UnindexedSecret" && number == 20) ? 2 : offsets.ContainsKey(number) ? 1 : 0;
            uint first = type == 2 ? (uint)container : type == 1 ? (uint)offsets[number] : 0;
            rows.Append(type.ToString("X2", CultureInfo.InvariantCulture)).Append(first.ToString("X8", CultureInfo.InvariantCulture))
                .Append((type == 2 ? index : number == 0 ? 65535 : 0).ToString("X4", CultureInfo.InvariantCulture));
        }
        rows.Append('>');
        output.Append(xref).Append(variant.StartsWith("EscapedTypes", StringComparison.Ordinal) ? " 0 obj\n<< /Type /XR#65f /Root 1 0 R /Size " : " 0 obj\n<< /Type /XRef /Root 1 0 R /Size ").Append(xref + 1)
            .Append(" /W [1 4 2] /Filter /ASCIIHexDecode /Length ").Append(rows.Length).Append(" >>\nstream\n").Append(rows)
            .Append("\nendstream\nendobj\nstartxref\n").Append(offsets[xref]).Append("\n%%EOF\n");
        if (variant == "WinningRevision") {
            int valueOffset = output.Length;
            output.Append(TextWidget(10, "/FT /Tx /T (Text) /V (new secret value)")).Append('\n');
            int nextXref = output.Length;
            string entries = "01" + valueOffset.ToString("X8", CultureInfo.InvariantCulture) + "0000" +
                "01" + nextXref.ToString("X8", CultureInfo.InvariantCulture) + "0000>";
            output.Append("200 0 obj\n<< /Type /XRef /Root 1 0 R /Size 201 /Prev ").Append(offsets[xref])
                .Append(" /Index [10 1 200 1] /W [1 4 2] /Filter /ASCIIHexDecode /Length ").Append(entries.Length)
                .Append(" >>\nstream\n").Append(entries).Append("\nendstream\nendobj\nstartxref\n").Append(nextXref).Append("\n%%EOF\n");
        }
        if (variant == "TruncatedClassicTable") {
            int tableOffset = output.Length;
            output.Append("xref\n0 110000\n");
            for (int number = 0; number < 110000; number++) {
                if (offsets.TryGetValue(number, out int objectOffset)) output.Append(objectOffset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
                else output.Append("0000000000 65535 f \n");
            }
            output.Append("trailer\n<< /Root 1 0 R /Size 110000 /Prev ").Append(offsets[xref]).Append(" >>\nstartxref\n")
                .Append(tableOffset).Append("\n%%EOF\n");
        }
        if (variant == "MalformedHistoryTrailer") {
            int tableOffset = output.Length;
            output.Append("xref\n0 ").Append(xref + 1).Append('\n');
            for (int number = 0; number <= xref; number++) {
                if (number != xref && offsets.TryGetValue(number, out int objectOffset)) output.Append(objectOffset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
                else output.Append("0000000000 65535 f \n");
            }
            output.Append("trailer junk << /Root 1 0 R /Size ").Append(xref + 1).Append(" /Prev ").Append(offsets[xref])
                .Append(" >>\nstartxref\n").Append(tableOffset).Append("\n%%EOF\n");
        }
        if (variant == "MissingCompressedPredecessor") output.Append("xref\n0 1\n0000000000 65535 f \ntrailer\n<< /Root 1 0 R /Size ")
            .Append(xref + 1).Append(" /Prev ").Append(offsets[xref]).Append(" >>\nstartxref\n")
            .Append(output.ToString().LastIndexOf("xref\n0", StringComparison.Ordinal)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(output.ToString());
    }
}
