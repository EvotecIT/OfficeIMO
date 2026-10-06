using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Sign;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Mobile;

internal sealed partial class MobileDocumentController {
    private static readonly StudioFileType PdfFiles = new("PDF documents", ["pdf"], "application/pdf");
    private static readonly StudioFileType WorkflowFiles = new("Documents and images", [
        "docx", "xlsx", "pptx", "pdf", "html", "htm", "pages", "numbers", "key",
        "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "ico", "pcx", "zip"]);

    private MainWindowViewModel CreateDocument(Func<string, CancellationToken, Task> openInTab) {
        MainWindowViewModel? document = null;
        string Text(string key) => _services.Localizer.Get(key);
        Task<string?> SavePdf(CancellationToken token) => Host.PickSaveFileAsync(Text("Picker.SavePdf"),
            Path.GetFileNameWithoutExtension(document!.DocumentName.TrimEnd(' ', '*')) + ".pdf", PdfFiles, token);
        document = new MainWindowViewModel(
            pickPdf: token => PickWorkingCopyAsync(document!, token),
            pickSavePdf: SavePdf,
            pickImportPdfs: token => Host.PickFilesAsync(Text("Picker.AddPdfDocuments"), PdfFiles, true, token),
            pickOutputFolder: token => Host.PickFolderAsync(Text("Picker.ChooseOutputFolder"), token),
            openUri: uri => Host.OpenUriAsync(uri),
            confirmUnsavedChanges: () => ConfirmUnsavedChangesAsync?.Invoke(document!) ?? Task.FromResult(UnsavedChangesDecision.Cancel),
            confirmProviderWrite: location => Host.ConfirmWriteAsync(location),
            confirmWorkflowProviderWrite: location => Host.ConfirmWriteAsync(location, workflow: true),
            pickImage: token => Host.PickImageAsync(token),
            confirmPageDeletion: static _ => Task.FromResult(true),
            reviewPageMove: preview => Host.ShowAsync<bool>(new PageMoveDialogContent(preview)),
            reviewPageSplit: preview => Host.ShowAsync<bool>(new PageSplitDialogContent(preview)),
            showPageSplitResult: result => Host.ShowAsync<object>(new PageSplitDialogContent(result)),
            reviewPageExtraction: preview => Host.ShowAsync<bool>(new PageExtractionDialogContent(preview)),
            showPageExtractionResult: result => Host.ShowAsync<object>(new PageExtractionDialogContent(result)),
            reviewPageImport: preview => Host.ShowAsync<bool>(new PageImportDialogContent(preview)),
            reviewProtection: preview => Host.ShowAsync<bool>(new PdfProtectionDialogContent(preview)),
            showProtectionResult: result => Host.ShowAsync<object>(new PdfProtectionDialogContent(result)),
            reviewSigning: preview => Host.ShowAsync<bool>(new PdfSigningDialogContent(preview)),
            showSigningResult: result => Host.ShowAsync<object>(new PdfSigningDialogContent(result)),
            reviewWatermark: preview => Host.ShowAsync<bool>(new WatermarkDialogContent(preview)),
            pickWorkflowFiles: token => Host.PickFilesAsync(Text("Picker.AddDocumentsOrImages"), WorkflowFiles, true, token),
            pickOcrFiles: token => Host.PickFilesAsync(Text("OcrSession.Add"),
                new("PDFs and images", ["pdf", "png", "jpg", "jpeg", "bmp", "tif", "tiff", "gif", "webp"]), true, token),
            pickAssemblyFolder: token => Host.PickFolderAsync(Text("Picker.AddSourceFolder"), token),
            pickProvenanceFile: async token => {
                string? copy = await Host.PickWorkingCopyAsync(Text("Provenance.ChooseFile"),
                    new("Provenance assets", OfficeProvenanceWorkflowCatalog.All.SelectMany(item => item.Extensions).ToArray()), token);
                if (copy is not null) document!.ProvenanceWorkbench.OutputFolder = Path.GetDirectoryName(copy)!;
                return copy;
            },
            pickSaveRedactionReport: token => Host.PickSaveFileAsync(Text("Redaction.ExportReport"), "redaction-evidence.json",
                new("JSON", ["json"], "application/json"), token),
            promptPdfPassword: (name, invalid, token) => Host.PromptPasswordAsync(name, invalid, token),
            recentDocumentStore: _services.DocumentHistory.RecentDocuments,
            canSaveAsPath: path => document is not null && Tabs.CanDocumentOwnPath(document, path),
            openDocumentInTab: openInTab,
            services: _services,
            supportsFolderNavigation: false,
            canPublishPath: path => Tabs.CanPublishPath(path),
            publicationGuard: new StudioWorkflowPublicationGuard((path, directory) =>
                directory ? Tabs.CanPublishDirectory(path) : Tabs.CanPublishPath(path)),
            confirmBookChanges: () => Host.ShowAsync<UnsavedChangesDecision>(new UnsavedChangesDialogContent(
                document?.BookWorkbench.BookTitle ?? "Book", _services.Localizer)),
            bookPublicationGuard: new StudioWorkflowPublicationGuard((path, _) => Tabs.CanPublishBookPath(document, path)));
        document.ProvenanceWorkbench.UsesWorkingCopies = true;
        if (_host is not null) {
            document.FileDialogs = _host;
            _host.ConfigureAssistant(document, path => Tabs.CanPublishPath(path));
        }
        document.CreateSignatureDialog = kind => Host.ShowAsync<StudioSignatureDraft>(new SignatureDialogContent(kind, _services.Localizer));
        document.PropertyChanged += OnDocumentChanged;
        _observed.Add(document);
        return document;
    }
}
