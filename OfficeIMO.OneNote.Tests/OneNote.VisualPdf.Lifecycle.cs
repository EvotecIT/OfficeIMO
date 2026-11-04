using System.Reflection;
using System.Threading;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Tests;

public sealed class OneNoteVisualPdfLifecycleTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancelledVisualOperationsStopBeforeInvalidRenderingConfiguration(bool notebookSource) {
        object source = notebookSource ? new OneNoteNotebook() : new OneNoteSection();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var stream = new MemoryStream();
        var options = new OneNoteVisualPdfExportOptions { RasterScale = 0D };

        foreach (MethodInfo method in typeof(OneNoteVisualPdfExtensions).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.GetParameters()[0].ParameterType == source.GetType())) {
            object?[] arguments = method.GetParameters().Select(parameter =>
                parameter.ParameterType == source.GetType() ? source :
                parameter.ParameterType == typeof(Stream) ? stream :
                parameter.ParameterType == typeof(string) ? string.Empty :
                parameter.ParameterType == typeof(OneNoteVisualPdfExportOptions) ? options :
                parameter.ParameterType == typeof(CancellationToken) ? (object)cancellation.Token : null).ToArray();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
                object? result;
                try { result = method.Invoke(null, arguments); }
                catch (TargetInvocationException exception) when (exception.InnerException != null) {
                    System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                    throw;
                }
                if (result is Task task) await task;
            });
        }
        Assert.Equal(0, stream.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task VisualSavesRejectInvalidRenderingSettingsAndRetainDestination(bool notebookSource) {
        var section = new OneNoteSection();
        var notebook = new OneNoteNotebook();
        var options = new OneNoteVisualPdfExportOptions { RasterScale = 0D };
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        using var stream = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() => notebookSource
            ? notebook.SaveAsVisualPdf(path, options)
            : section.SaveAsVisualPdf(path, options));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => notebookSource
            ? notebook.SaveAsVisualPdfAsync(stream, options)
            : section.SaveAsVisualPdfAsync(stream, options));
        Assert.False(File.Exists(path));
        Assert.Equal(0, stream.Length);
        Assert.Throws<ArgumentOutOfRangeException>(() => {
            if (notebookSource) notebook.SaveAsVisualPdf(stream, options);
            else section.SaveAsVisualPdf(stream, options);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VisualByteOutputProducesReadablePdf(bool notebookSource) {
        var section = new OneNoteSection { Name = "Visual lifecycle" };
        section.Pages.Add(new OneNotePage { Title = "A", PageSize = OneNotePageSize.IndexCard });
        var notebook = new OneNoteNotebook();
        notebook.Sections.Add(section);
        var options = new OneNoteVisualPdfExportOptions { RasterScale = 0.1D };
        byte[] bytes = notebookSource ? notebook.ToVisualPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }) : section.ToVisualPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        using var stream = new MemoryStream(bytes);
        PdfDocument document = PdfDocument.Load(stream);
        Assert.True(document.Analyze().Preflight.CanRead);
        Assert.Equal(1, document.Inspect().PageCount);
    }
}
