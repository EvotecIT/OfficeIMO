using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    private byte[] BuildWordHybridRedactionFixture(bool collision) {
        string raw = Encoding.ASCII.GetString(BuildFormResidueSource("TU"));
        var objects = Regex.Matches(raw, @"(?s)(\d+) 0 obj\s*\n(.*?)\nendobj").Cast<Match>()
            .ToDictionary(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), static match => match.Groups[2].Value);
        objects[1] = objects[1].Replace("/Type /Catalog", "/Type /Catalog /MarkInfo << /Marked true >> /StructTreeRoot 20 0 R");
        objects[20] = "<< /Type /StructTreeRoot /K [29 0 R] /ParentTree << /Nums [] >> >>";
        objects[29] = "<< /Type /StructElem /S /Document /P 20 0 R /K [] >>";
        objects[28] = "<< /Producer (Synthetic Word hybrid) >>";
        for (int number = 1; number < 45; number++) {
            if (!objects.ContainsKey(number)) objects.Add(number, "<< /Value (public) >>");
        }
        if (collision) {
            objects[1] = objects[1].Replace("/Type /Catalog", "/Type /Catalog /PrivateData 30 0 R");
            objects[5] = objects[5].Replace("Public", "secret");
            objects[30] = "<< /Missing 47 0 R /Array [47 0 R] /IndirectArray 31 0 R /Stream 32 0 R >>";
            objects[31] = "[47 0 R]";
            objects[32] = "<< /Length 6 /Missing 47 0 R >>\nstream\npublic\nendstream";
        }
        int[] packed = objects.Keys.OrderBy(static number => number).Where(number => !objects[number].Contains("\nstream\n", StringComparison.Ordinal)).Take(25).ToArray();
        var header = new StringBuilder();
        var members = new StringBuilder();
        foreach (int number in packed) {
            header.Append(number).Append(' ').Append(members.Length).Append(' ');
            members.Append(objects[number]).Append("\r\n");
        }
        Encoding latin1 = Encoding.GetEncoding(28591);
        var output = new StringBuilder("%PDF-1.5\r\n");
        var offsets = new Dictionary<int, int>();
        foreach (int number in objects.Keys.OrderBy(static number => number)) {
            if (!packed.Contains(number)) Emit(number, objects[number].Replace("\n", "\r\n"));
        }
        Emit(45, CompressedStream("/Type /ObjStm /N 25 /First " + header.Length, Encoding.ASCII.GetBytes(header.ToString() + members)));
        const string identity = "/Root 1 0 R /Info 28 0 R /ID [<00112233445566778899AABBCCDDEEFF> <00112233445566778899AABBCCDDEEFF>]";
        using var rows = new MemoryStream();
        for (int number = 0; number < 46; number++) {
            int member = Array.IndexOf(packed, number);
            int type = member >= 0 ? 2 : offsets.ContainsKey(number) ? 1 : 0;
            uint first = type == 2 ? 45U : type == 1 ? (uint)offsets[number] : 0U;
            int second = type == 2 ? member : type == 0 ? 65535 : 0;
            rows.WriteByte((byte)type);
            for (int shift = 24; shift >= 0; shift -= 8) rows.WriteByte((byte)(first >> shift));
            rows.WriteByte((byte)(second >> 8)); rows.WriteByte((byte)second);
        }
        Emit(46, CompressedStream("/Type /XRef /Size 46 /W[1 4 2] " + identity, rows.ToArray()));
        int classic = output.Length;
        output.Append("xref\r\n0 47\r\n");
        for (int number = 0; number < 47; number++) {
            if (offsets.TryGetValue(number, out int offset)) output.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n\r\n");
            else output.Append("0000000000 65535 f\r\n");
        }
        output.Append("trailer\r\n<< /Size 47 ").Append(identity).Append(" >>\r\nstartxref\r\n").Append(classic).Append("\r\n%%EOF\r\n");
        int final = output.Length;
        output.Append("xref\r\n0 0\r\ntrailer\r\n<< /Size 47 ").Append(identity).Append(" /Prev ").Append(classic)
            .Append(" /XRefStm ").Append(offsets[46]).Append(" >>\r\nstartxref\r\n").Append(final).Append("\r\n%%EOF\r\n");
        return latin1.GetBytes(output.ToString());

        void Emit(int number, string body) {
            offsets[number] = output.Length;
            output.Append(number).Append(" 0 obj\r\n").Append(body).Append("\r\nendobj\r\n");
        }
        string CompressedStream(string dictionary, byte[] data) {
            byte[] compressed = PdfFlateEncoder.Compress(data);
            return "<< " + dictionary + " /Filter /FlateDecode /Length " + compressed.Length + " >>\r\nstream\r\n" + latin1.GetString(compressed) + "\r\nendstream";
        }
    }

    private static void AssertWordHybridOffsets(byte[] input) {
        string raw = Encoding.GetEncoding(28591).GetString(input);
        PdfReadDocument read = PdfReadDocument.Open(input);
        Assert.Equal(46, read.Objects.Count);
        Assert.Equal(raw.IndexOf("45 0 obj\r\n", StringComparison.Ordinal), read.Objects[10].SourceOffset);
        Assert.Equal(raw.IndexOf("46 0 obj\r\n", StringComparison.Ordinal), read.Objects[46].SourceOffset);
        foreach (int number in new[] { 45, 46 }) {
            PdfStream stream = Assert.IsType<PdfStream>(read.Objects[number].Value);
            int start = raw.IndexOf("stream\r\n", read.Objects[number].SourceOffset, StringComparison.Ordinal) + 8;
            int end = start + (int)stream.Dictionary.Get<PdfNumber>("Length")!.Value;
            Assert.Equal("\r\nendstream", raw.Substring(end, 11));
        }
    }

    private static void AssertSavedReferenceTarget(byte[] output, bool live) {
        PdfReadDocument read = PdfReadDocument.Open(output);
        PdfDictionary catalog = PdfSyntax.FindCatalog(read.Objects, read.TrailerRaw)!;
        PdfDictionary owner = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, catalog.Items["PrivateData"]));
        PdfObject? value = PdfObjectLookup.Resolve(read.Objects, owner.Items["Missing"]);
        if (live) Assert.Equal("kept live value", Assert.IsType<PdfStringObj>(Assert.IsType<PdfDictionary>(value).Items["Value"]).Value);
        else Assert.IsType<PdfNull>(value);
        if (owner.Items.TryGetValue("Array", out PdfObject? array)) {
            Assert.IsType<PdfNull>(Assert.Single(Assert.IsType<PdfArray>(array).Items));
            Assert.IsType<PdfNull>(Assert.Single(Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(read.Objects, owner.Items["IndirectArray"])).Items));
            PdfStream stream = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(read.Objects, owner.Items["Stream"]));
            Assert.IsType<PdfNull>(stream.Dictionary.Items["Missing"]);
        }
    }
}
