using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures the catalog /Dests form separately from name-tree bookmarks.</summary>
[MemoryDiagnoser]
public class PdfDirectDestinationScalingBenchmarks {
    private byte[] _source = null!;
    private byte[] _plainSource = null!;
    private int[] _selectedPages = null!;

    [Params(5, 6, 50, 500, 2000)]
    public int PageCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _source = BuildSource(PageCount, includeDestinations: true);
        _plainSource = BuildSource(PageCount, includeDestinations: false);
        PdfDocumentPreflight preflight = PdfDocument.Preflight(_source);
        if (!preflight.CanRewrite || preflight.DocumentInfo?.PageCount != PageCount || !preflight.Probe.HasNamedDestinations) {
            throw new InvalidOperationException("Direct-destination source is not rewrite-safe.");
        }

        byte[][] parts = Split();
        if (parts.Length != PageCount) throw new InvalidOperationException("Split output count changed.");
        for (int i = 0; i < parts.Length; i++) {
            Validate(parts[i], new[] { i + 1 });
        }

        byte[][] plainParts = SplitPlain();
        if (plainParts.Length != PageCount) throw new InvalidOperationException("Plain split output count changed.");
        for (int i = 0; i < plainParts.Length; i++) {
            PdfReadDocument plainReadback = PdfReadDocument.Open(plainParts[i]);
            if (plainReadback.Pages.Count != 1 ||
                plainReadback.Pages[0].GetPageSize().Width != 200 + i + 1 ||
                plainReadback.NamedDestinations.Count != 0) {
                throw new InvalidOperationException("Plain split output changed at page " + (i + 1) + ".");
            }
        }

        int selectedCount = Math.Max(2, PageCount / 4);
        _selectedPages = new int[selectedCount];
        for (int i = 0; i < selectedCount; i++) {
            _selectedPages[i] = PageCount - (int)Math.Round(i * (PageCount - 1D) / (selectedCount - 1D));
        }
        Validate(Select(), _selectedPages);
    }

    [Benchmark]
    public byte[][] Split() => PdfDocument.Load(_source).Pages.Split().Select(static part => part.ToBytes()).ToArray();

    [Benchmark]
    public byte[][] SplitPlain() => PdfDocument.Load(_plainSource).Pages.Split().Select(static part => part.ToBytes()).ToArray();

    [Benchmark]
    public byte[] Select() => PdfDocument.Load(_source).Pages.Extract(_selectedPages).ToBytes();

    [Benchmark]
    public PdfReadDocument OpenSource() => PdfReadDocument.Open(_source);

    [Benchmark]
    public PdfReadDocument OpenPlainSource() => PdfReadDocument.Open(_plainSource);

    [Benchmark]
    public PdfDocumentPreflight PreflightSource() => PdfDocument.Preflight(_source);

    [Benchmark]
    public PdfDocumentInfo InspectSource() => PdfDocument.Load(_source).Inspect();

    private static void Validate(byte[] output, IReadOnlyList<int> sourcePages) {
        PdfReadDocument readback = PdfReadDocument.Open(output);
        if (readback.Pages.Count != sourcePages.Count || readback.NamedDestinations.Count != sourcePages.Count) {
            throw new InvalidOperationException("Direct destinations or selected pages were lost.");
        }
        for (int i = 0; i < sourcePages.Count; i++) {
            int sourcePage = sourcePages[i];
            if (readback.Pages[i].GetPageSize().Width != 200 + sourcePage ||
                !readback.NamedDestinations.Any(destination =>
                    destination.Name == "Dest" + sourcePage.ToString("D4", CultureInfo.InvariantCulture) &&
                    destination.PageNumber == i + 1)) {
                throw new InvalidOperationException("Direct destination or page order changed at output page " + (i + 1) + ".");
            }
        }
    }

    private static byte[] BuildSource(int pageCount, bool includeDestinations) {
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>(pageCount + (includeDestinations ? 4 : 3)) { 0 };
        int destinationsId = pageCount + 3;
        offsets.Add(pdf.Length);
        pdf.Append("1 0 obj\n<< /Type /Catalog /Pages 2 0 R");
        if (includeDestinations) pdf.Append(" /Dests ").Append(destinationsId).Append(" 0 R");
        pdf.Append(" >>\nendobj\n");
        offsets.Add(pdf.Length);
        pdf.Append("2 0 obj\n<< /Type /Pages /Count ").Append(pageCount).Append(" /Kids [");
        for (int page = 1; page <= pageCount; page++) pdf.Append(page + 2).Append(" 0 R ");
        pdf.Append("] >>\nendobj\n");
        for (int page = 1; page <= pageCount; page++) {
            offsets.Add(pdf.Length);
            pdf.Append(page + 2).Append(" 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ")
                .Append(200 + page).Append(" 200] >>\nendobj\n");
        }
        if (includeDestinations) {
            offsets.Add(pdf.Length);
            pdf.Append(destinationsId).Append(" 0 obj\n<< ");
            for (int page = 1; page <= pageCount; page++) {
                pdf.Append("/Dest").Append(page.ToString("D4", CultureInfo.InvariantCulture))
                    .Append(" [").Append(page + 2).Append(" 0 R /Fit] ");
            }
            pdf.Append(">>\nendobj\n");
        }
        int xrefOffset = pdf.Length;
        pdf.Append("xref\n0 ").Append(offsets.Count).Append("\n0000000000 65535 f \n");
        for (int objectId = 1; objectId < offsets.Count; objectId++) {
            pdf.Append(offsets[objectId].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        pdf.Append("trailer\n<< /Root 1 0 R /Size ").Append(offsets.Count)
            .Append(" >>\nstartxref\n").Append(xrefOffset).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
