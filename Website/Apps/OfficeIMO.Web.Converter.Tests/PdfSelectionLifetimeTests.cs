using System.Reflection;
using Microsoft.AspNetCore.Components.Forms;
using OfficeIMO.Web.Converter.Components;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed partial class ProvenanceSelectionLifetimeTests {
    [Theory]
    [InlineData("merge", true)]
    [InlineData("compare", true)]
    [InlineData("inspect", false)]
    public async Task PdfPickerRetainsCarriedResultForMultiFileTools(string tool, bool append) {
        var session = new BrowserDocumentSession();
        session.Open([Pdf("original.pdf")]);
        session.SetResult([4, 5, 6], "converted.pdf", session.Revision);
        session.UseResult();
        var carried = Assert.Single(session.Current);
        var component = await PdfComponent(session, tool);
        await SelectPdf(component, new UploadedPdf());
        Assert.Equal(append ? 2 : 1, session.Current.Count);
        if (append) Assert.Same(carried, session.Current[0]);
        Assert.Equal("additional.pdf", session.Current.Last().Name);
        Assert.Equal(session.Current, PdfFiles(component));
        if (tool == "merge") {
            await InvokePdf(component, "MoveFileAsync", new PdfFileMoveRequest(1, -1));
            Assert.Equal("additional.pdf", session.Current[0].Name);
            Assert.Same(carried, session.Current[1]);
            await InvokePdf(component, "RemoveFileAsync", 0);
            Assert.Same(carried, Assert.Single(session.Current));
        }
        session.RestoreOriginals();
        Assert.Equal(append ? "original.pdf" : "additional.pdf", Assert.Single(session.Current).Name);
    }

    [Theory]
    [InlineData("compare", 2, false)]
    [InlineData("merge", 10, false)]
    [InlineData("merge", 3, true)]
    public async Task PdfPickerRejectsCombinedLimitsWithoutReplacingCarriedFiles(string tool, int count, bool aggregateLimit) {
        byte[] bytes = aggregateLimit ? new byte[BrowserConversionService.MaxPackageBytes] : [1];
        var originals = Enumerable.Range(0, count).Select(i => new SelectedDocument($"{i}.pdf", ".pdf", "PDF", bytes.LongLength, bytes)).ToArray();
        var session = new BrowserDocumentSession(); session.Open(originals);
        var component = await PdfComponent(session, tool);
        var upload = new UploadedPdf();
        await SelectPdf(component, upload);
        Assert.Equal(0, upload.Reads);
        Assert.Equal(originals, session.Current);
        Assert.Equal(originals, PdfFiles(component));
    }

    [Fact]
    public async Task PdfAppendRejectsAnInvalidBatchWithoutCommittingItsValidFirstFile() {
        var session = new BrowserDocumentSession(); session.Open([Pdf("carried.pdf")]);
        var carried = Assert.Single(session.Current);
        var component = await PdfComponent(session, "merge");
        var valid = new UploadedPdf();
        var invalid = new UploadedPdf("document.docx");
        await InvokePdf(component, "HandleFilesSelectedAsync", new InputFileChangeEventArgs([valid, invalid]));
        Assert.Equal(1, valid.Reads);
        Assert.Equal(0, invalid.Reads);
        Assert.Same(carried, Assert.Single(session.Current));
        Assert.Same(carried, Assert.Single(PdfFiles(component)));
    }

    [Fact]
    public async Task PdfAppendCannotResurrectSelectionAfterCleanupAndSessionClear() {
        var session = new BrowserDocumentSession(); session.Open([Pdf("carried.pdf")]);
        var component = await PdfComponent(session, "merge");
        var js = new DelayedRevocation();
        Set(component, "_interop", new ConverterInterop(js)); Set(component, "ArtifactUrl", "blob:old");
        var upload = new UploadedPdf();
        var pending = SelectPdf(component, upload);
        await js.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        session.Clear();
        await InvokePdf(component, "OnParametersSetAsync");
        js.Release.SetResult(); await pending;
        Assert.Equal(0, upload.Reads);
        Assert.Empty(session.Current); Assert.Empty(PdfFiles(component));
    }

    private static SelectedDocument Pdf(string name) => new(name, ".pdf", "PDF", 1, [1]);
    private static async Task<PdfWorkbench> PdfComponent(BrowserDocumentSession session, string tool) {
        var component = new PdfWorkbench(); Set(component, "Session", session); Set(component, "ToolId", tool);
        await InvokePdf(component, "OnParametersSetAsync"); return component;
    }
    private static List<SelectedDocument> PdfFiles(PdfWorkbench component) =>
        (List<SelectedDocument>)typeof(PdfWorkbench).GetProperty("Files", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(component)!;
    private static Task SelectPdf(PdfWorkbench component, IBrowserFile file) => InvokePdf(component, "HandleFilesSelectedAsync", new InputFileChangeEventArgs([file]));
    private static Task InvokePdf(PdfWorkbench component, string method, params object[] args) =>
        (Task)typeof(PdfWorkbench).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(component, args)!;
    private sealed class UploadedPdf(string name = "additional.pdf") : IBrowserFile {
        public string Name => name;
        public DateTimeOffset LastModified => DateTimeOffset.UnixEpoch;
        public long Size => 1;
        public string ContentType => "application/pdf";
        public int Reads { get; private set; }
        public Stream OpenReadStream(long maxAllowedSize = 512000, CancellationToken cancellationToken = default) { Reads++; return new MemoryStream([2]); }
    }
}
