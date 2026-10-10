using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfProvenanceManifestSummaryTests {
    [Fact]
    public void CompressedFilespecAliasesRetainTheirIdentityAndReachTheMutationSafetyGate() {
        byte[] pdf = RevisionFixture(sameRevision: true, compressedAliases: true);
        var report = PdfProvenance.Inspect(pdf);
        Assert.Equal(3, report.Evidence.Count); // Two aliases plus the separate current store.
        Assert.Throws<PdfMutationBlockedException>(() => PdfProvenance.Remove(pdf));
    }

    [Fact]
    public void DistinctDirectFilespecsRetainEvidenceAndRejectUnsafeRemoval() {
        byte[] pdf = RevisionFixture(sameRevision: true, directAliases: true);
        var report = PdfProvenance.Inspect(pdf);
        Assert.Equal(2, report.Evidence.Count);
        Assert.All(report.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.Throws<InvalidDataException>(() => PdfProvenance.Remove(pdf,
            new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false }));
    }
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HistoricalStoresFollowStreamAndHybridCrossReferences(bool streamXref, bool hybrid) {
        var report = PdfProvenance.Inspect(RevisionFixture(replaceAssociations: true, streamXref: streamXref, hybrid: hybrid));
        Assert.Equal(2, report.Evidence.Count);
        var active = Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!;
        Assert.Equal("Current editor", active.ClaimGenerator);
        Assert.Equal(2, active.ManifestCount);
    }

    [Fact]
    public void CompetingLaterStoresFallBackToTheEarlierAssociatedStore() {
        var report = PdfProvenance.Inspect(RevisionFixture(replaceAssociations: true, competing: true));
        Assert.Equal(3, report.Evidence.Count);
        var active = Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!;
        Assert.Equal("Previous editor", active.ClaimGenerator);
        Assert.Equal(1, active.ManifestCount);
        Assert.Contains(report.Diagnostics, message => message.Contains("multiple C2PA stores"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FullRewriteRemovesHistoricalCredentialsAndPreservesPages(bool historicalOnly) {
        byte[] pdf = RevisionFixture(replaceAssociations: true, historicalOnly: historicalOnly);
        var removed = PdfProvenance.Remove(pdf);
        Assert.NotEmpty(removed.Changes);
        Assert.Empty(removed.After.Evidence);
        Assert.DoesNotContain("Previous editor", Encoding.ASCII.GetString(removed.ToArray()));
        Assert.Single(PdfReadDocument.Open(removed.ToArray()).Pages);
    }

    [Fact]
    public void HistoricalInspectionChargesItsPrefixCopiesToTheCumulativeBudget() {
        Assert.ThrowsAny<InvalidDataException>(() => PdfProvenance.Inspect(RevisionFixture(replaceAssociations: true),
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 1024 }));
    }

    [Fact]
    public void RetainedOldNameWithoutItsOldAssociationUsesTheOwningRevisionProfile() {
        var report = PdfProvenance.Inspect(RevisionFixture(replaceAssociations: true, retainOldName: true));
        Assert.Equal(2, report.Evidence.Count);
        Assert.All(report.Evidence, item => Assert.True(item.IsStructurallyValid));
        Assert.Equal(2, Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!.ManifestCount);
    }
    [Fact]
    public void ReplacedCatalogAssociationsRetainEarlierLogicalStores() {
        OfficeProvenanceReport report = PdfProvenance.Inspect(RevisionFixture(replaceAssociations: true));
        Assert.Equal(2, report.Evidence.Count);
        OfficeC2paManifestSummary active = Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!;
        Assert.Equal("Current editor", active.ClaimGenerator);
        Assert.Equal(2, active.ManifestCount);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncrementalStoreIsActiveRegardlessOfNameOrderAndObjectNumber(bool payloadMarkers) {
        byte[] pdf = RevisionFixture(payloadMarkers: payloadMarkers);
        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf);
        Assert.Equal(2, report.Evidence.Count);
        Assert.All(report.Evidence, item => Assert.True(item.IsStructurallyValid));
        OfficeC2paManifestSummary active = Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!;
        Assert.Equal("Current editor", active.ClaimGenerator);
        Assert.Equal(2, active.ManifestCount);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void AliasesToOneEmbeddedStoreDoNotCountAsAdditionalManifests() {
        OfficeProvenanceReport report = PdfProvenance.Inspect(RevisionFixture(alias: true));
        Assert.Equal(3, report.Evidence.Count);
        OfficeC2paManifestSummary active = Assert.Single(report.Evidence, item => item.Manifest != null).Manifest!;
        Assert.Equal("Current editor", active.ClaimGenerator);
        Assert.Equal(2, active.ManifestCount);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void MultipleStoresInOneUpdateDoNotAcquireAnArbitraryActiveSummary() {
        OfficeProvenanceReport report = PdfProvenance.Inspect(RevisionFixture(sameRevision: true));
        Assert.Equal(2, report.Evidence.Count);
        Assert.All(report.Evidence, item => Assert.Null(item.Manifest));
        Assert.Contains(report.Diagnostics, message => message.Contains("multiple C2PA stores"));
    }

    // An independent, minimal classic-xref PDF. The newer embedded stream deliberately uses
    // a lower object number; the old store remains first in both the name tree and AF list.
    private static byte[] RevisionFixture(bool sameRevision = false, bool payloadMarkers = false, bool alias = false,
        bool replaceAssociations = false, bool streamXref = false, bool hybrid = false, bool competing = false, bool historicalOnly = false,
        bool retainOldName = false, bool compressedAliases = false, bool directAliases = false) {
        alias |= compressedAliases;
        streamXref |= compressedAliases;
        using var output = new MemoryStream();
        var offsets = new Dictionary<int, long>();
        void Text(string value) { byte[] bytes = Encoding.ASCII.GetBytes(value); output.Write(bytes, 0, bytes.Length); }
        void Object(int id, string body, byte[]? stream = null) {
            offsets[id] = output.Position;
            Text($"{id} 0 obj\n{body}");
            if (stream != null) { Text("\nstream\n"); output.Write(stream, 0, stream.Length); Text("\nendstream"); }
            Text("\nendobj\n");
        }
        string DirectSpec() => "<< /Type /Filespec /F (a-old.c2pa) /AFRelationship /C2PA_Manifest /EF << /F 5 0 R >> >>";
        string Catalog(bool current) => directAliases
            ? $"<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(first) {DirectSpec()} (second) {DirectSpec()}] >> >> >>"
            : current && historicalOnly ? "<< /Type /Catalog /Pages 2 0 R >>" : current && replaceAssociations
            ? "<< /Type /Catalog /Pages 2 0 R /AF [8 0 R" + (competing ? " 10 0 R" : "") + "] /Names << /EmbeddedFiles << /Names [" + (retainOldName ? "(a-old.c2pa) 6 0 R " : "") + "(z-current.c2pa) 8 0 R" + (competing ? " (second.c2pa) 10 0 R" : "") + "] >> >> >>"
            : "<< /Type /Catalog /Pages 2 0 R /AF [6 0 R" + (alias ? " 7 0 R" : "") + (current ? " 8 0 R" : "") +
            "] /Names << /EmbeddedFiles << /Names [(a-old.c2pa) 6 0 R" + (alias ? " (b-alias.c2pa) 7 0 R" : "") + (current ? " (z-current.c2pa) 8 0 R" : "") + "] >> >> >>";
        void Store(int streamId, int specId, string name, string generator) {
            byte[] store = ManifestStore(generator);
            Object(streamId, $"<< /Type /EmbeddedFile /Subtype /application#2Fc2pa /Length {store.Length} >>", store);
            if (!compressedAliases || specId != 6)
                Object(specId, $"<< /Type /Filespec /F ({name}.c2pa) /AFRelationship /C2PA_Manifest /EF << /F {streamId} 0 R >> >>");
        }
        long StreamXref(int[] ids, bool terminal, long? previous = null) {
            const int xrefId = 18;
            offsets[xrefId] = output.Position;
            int[] entries = ids.Append(xrefId).Distinct().OrderBy(id => id).ToArray();
            using var data = new MemoryStream();
            foreach (int id in entries) {
                bool compressed = compressedAliases && (id == 6 || id == 7);
                data.WriteByte(compressed ? (byte)2 : (byte)1);
                long offset = compressed ? 9 : offsets[id];
                for (int shift = 24; shift >= 0; shift -= 8) data.WriteByte((byte)(offset >> shift));
                data.WriteByte(0); data.WriteByte(compressed && id == 7 ? (byte)1 : (byte)0);
            }
            string indices = string.Join(" ", entries.Select(id => $"{id} 1"));
            Object(xrefId, $"<< /Type /XRef /Size 19 /Root 1 0 R /W [1 4 2] /Index [{indices}] /Length {data.Length}" +
                (previous.HasValue ? $" /Prev {previous}" : "") + " >>", data.ToArray());
            long start = offsets[xrefId];
            if (terminal) Text($"startxref\n{start}\n%%EOF\n");
            return start;
        }
        long Xref(int[] ids, long? previous = null, long? hybridOffset = null) {
            long position = output.Position;
            Text("xref\n0 1\n0000000000 65535 f \n");
            foreach (int id in ids.OrderBy(id => id)) Text($"{id} 1\n{offsets[id]:D10} 00000 n \n");
            if (previous.HasValue && (streamXref || hybrid)) Text("18 1\n0000000000 00001 f \n"); // Retire the old xref stream.
            Text($"trailer\n<< /Size 19 /Root 1 0 R" + (previous.HasValue ? $" /Prev {previous}" : "") +
                (hybridOffset.HasValue ? $" /XRefStm {hybridOffset}" : "") + $" >>\nstartxref\n{position}\n%%EOF\n");
            return position;
        }
        Text("%PDF-1.7\n");
        Object(1, Catalog(sameRevision));
        Object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
        Object(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>");
        Store(5, 6, "a-old", payloadMarkers ? "Previous editor startxref\n999999\n" : "Previous editor");
        if (compressedAliases) {
            string spec = "<< /Type /Filespec /F (a-old.c2pa) /AFRelationship /C2PA_Manifest /EF << /F 5 0 R >> >>";
            string header = $"6 0 7 {spec.Length + 1} ";
            byte[] data = Encoding.ASCII.GetBytes(header + spec + " " + spec);
            Object(9, $"<< /Type /ObjStm /N 2 /First {header.Length} /Length {data.Length} >>", data);
            offsets[6] = offsets[7] = offsets[9];
        } else if (alias) Object(7, "<< /Type /Filespec /F (b-alias.c2pa) /AFRelationship /C2PA_Manifest /EF << /F 5 0 R >> >>");
        if (sameRevision) Store(4, 8, "z-current", "Current editor");
        long? xrefStream = streamXref || hybrid ? StreamXref(offsets.Keys.ToArray(), streamXref && !hybrid) : null;
        long previous = streamXref && !hybrid ? xrefStream!.Value : Xref(offsets.Keys.ToArray(), hybridOffset: xrefStream);
        if (!sameRevision) {
            Store(4, 8, "z-current", "Current editor");
            if (competing) Store(9, 10, "second", "Competing editor");
            Object(1, Catalog(true));
            int[] changed = competing ? new[] { 1, 4, 8, 9, 10 } : new[] { 1, 4, 8 };
            if (streamXref && hybrid) StreamXref(changed, true, previous); // Replace the old xref-stream object.
            else Xref(changed, previous);
        }
        return output.ToArray();
    }

    private static byte[] ManifestStore(string generator) {
        byte[] Join(params byte[][] chunks) => chunks.SelectMany(chunk => chunk).ToArray();
        byte[] Box(string type, byte[] value) {
            int length = value.Length + 8;
            return Join(new[] { (byte)(length >> 24), (byte)(length >> 16), (byte)(length >> 8), (byte)length }, Encoding.ASCII.GetBytes(type), value);
        }
        byte[] Description(string uuid, string label) => Box("jumd", Join(Encoding.ASCII.GetBytes(uuid),
            new byte[] { 0, 17, 0, 16, 128, 0, 0, 170, 0, 56, 155, 113, 3 }, Encoding.ASCII.GetBytes(label + "\0")));
        byte[] CborText(string value) {
            byte[] text = Encoding.UTF8.GetBytes(value);
            return Join(text.Length < 24 ? new[] { (byte)(0x60 + text.Length) } : new[] { (byte)0x78, (byte)text.Length }, text);
        }
        byte[] claim = Join(new byte[] { 0xA1 }, CborText("claim_generator"), CborText(generator));
        byte[] manifest = Box("jumb", Join(Description("c2ma", "manifest"),
            Box("jumb", Join(Description("c2as", "c2pa.assertions"),
                Box("jumb", Join(Description("cbor", "c2pa.actions"), Box("cbor", new byte[] { 0xA0 }))))),
            Box("jumb", Join(Description("c2cl", "c2pa.claim"), Box("cbor", claim))),
            Box("jumb", Join(Description("c2cs", "c2pa.signature"), Box("cbor", new byte[] { 0xA0 })))));
        byte[] store = Box("jumb", Join(Description("c2pa", "c2pa"), manifest));
        Assert.True(OfficeC2paManifestStore.IsValid(store, 0, store.Length, store.Length, 1024, out _));
        Assert.NotNull(OfficeC2paManifestStore.TryDescribe(store, 0, store.Length));
        return store;
    }
}
