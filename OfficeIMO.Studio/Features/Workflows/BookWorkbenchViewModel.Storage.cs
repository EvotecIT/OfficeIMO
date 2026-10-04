using System.Security.Cryptography;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class BookWorkbenchViewModel {
    private string? _sourceIdentity, _projectIdentity;
    internal IEnumerable<string> OwnedLocations => new[] { _sourceLocation, _projectLocation }.OfType<string>().Distinct();
    internal bool OwnsPath(string path) => OwnedLocations.Any(source => OfficeStorageIdentity.AreEquivalent(source, path));
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task OpenBookAsync() {
        string? location = await _dialogs().PickOpenFileAsync(Title,
            new StudioFileType("Book or manuscript", ["oibook", "epub", "docx", "md", "markdown", "html", "htm"]), CancellationToken.None);
        if (location == null || !await PrepareCloseAsync()) return;
        await RunAsync(token => OpenLocationAsync(location, token));
    }
    internal async Task OpenLocationAsync(string location, CancellationToken token) {
        string extension = Path.GetExtension(_storage.Describe(location).Name).ToLowerInvariant();
        StudioStorageSnapshot input = await _storage.ReadSnapshotAsync(location, token, extension == ".oibook" ? 130L * 1024 * 1024 : 128L * 1024 * 1024);
        BookProject project;
        if (extension == ".oibook") project = await Task.Run(() => BookProject.LoadProject(input.Bytes, token), token);
        else if (extension == ".epub") project = await Task.Run(() => BookProject.FromEpub(input.Bytes, token), token);
        else {
            var options = new EpubManuscriptOptions();
            string? local = OfficeStorageIdentity.GetLocalPath(location);
            HtmlConversionDocumentOptions? htmlOptions = null;
            if (local != null) {
                options.ResourceResolver = BookManuscriptImporter.CreateLocalResourceResolver(local);
                var policy = new HtmlUrlPolicy { DisallowFileUrls = false, AllowDataUrls = true, RestrictUrlSchemes = true };
                policy.AllowedUrlSchemes.Clear(); policy.AllowedUrlSchemes.Add("file"); policy.AllowedUrlSchemes.Add("data");
                htmlOptions = new HtmlConversionDocumentOptions { BaseUri = new Uri(Path.GetFullPath(local)), ResourceUrlPolicy = policy };
            }
            EpubManuscriptResult imported = await Task.Run(() => BookManuscriptImporter.ImportBytesAsync(input.Bytes, extension, options, htmlOptions, token), token);
            project = BookProject.FromImport(imported);
        }
        token.ThrowIfCancellationRequested();
        _project = project; ResetLocation(); _sourceLocation = location; _sourceIdentity = input.Identity;
        if (extension == ".oibook") { _projectLocation = location; _projectFingerprint = Convert.ToHexString(SHA256.HashData(input.Bytes)); _projectIdentity = input.Identity; }
        ProjectName = _storage.Describe(location).Name;
        IsDirty = extension != ".oibook"; RefreshBook();
        Status = project.CanExport ? T("Imported", "Book opened. Review the chapters and preview before export.") :
            T("ReviewRequired", "Review the import findings. Failed dependencies require source repair and re-import.");
    }
    [RelayCommand(CanExecute = nameof(CanEditBook))]
    private Task SaveProjectAsync() => RunAsync(async token => {
        string? destination = await _dialogs().PickSaveFileAsync(T("SaveProject", "Save book project"),
            "book.oibook", new StudioFileType("OfficeIMO book project", ["oibook"]), token);
        if (destination == null) return;
        await ApplyDraftsAsync(token);
        byte[] bytes = await Task.Run(() => _project!.ToProjectBytes(token), token);
        string? expected = SameLocation(destination, _projectLocation) ? _projectFingerprint : await GetDestinationFingerprintAsync(destination, token);
        StudioStoragePublication saved = await _storage.PublishAsync(destination, bytes, expected,
            ct => AuthorizeAsync(destination, allowProjectSource: true, ct), token);
        _projectLocation = destination; _projectFingerprint = saved.Fingerprint;
        _projectIdentity = saved.Identity;
        ProjectName = _storage.Describe(destination).Name; IsDirty = false;
        Status = T("Saved", "Book project saved and verified.");
    });
    [RelayCommand(CanExecute = nameof(CanExport))]
    private Task ExportBookAsync() => RunAsync(async token => {
        string? destination = await _dialogs().PickSaveFileAsync(T("Export", "Export EPUB"),
            "book.epub", new StudioFileType("EPUB publication", ["epub"], "application/epub+zip"), token);
        if (destination == null) return;
        await ApplyDraftsAsync(token);
        byte[] bytes = await Task.Run(() => _project!.Export(token).Bytes, token);
        string? expected = await GetDestinationFingerprintAsync(destination, token);
        await _storage.PublishAsync(destination, bytes, expected, ct => AuthorizeAsync(destination, false, ct), token);
        Status = T("Exported", "EPUB exported and verified. Save the book project to retain your editing and review state.");
    });
    private static bool SameLocation(string location, string? other) => other != null &&
        OfficeStorageIdentity.Normalize(location) == OfficeStorageIdentity.Normalize(other);
    private async Task<string?> GetDestinationFingerprintAsync(string location, CancellationToken token) {
        try { return await _storage.FingerprintAsync(location, token); }
        catch (FileNotFoundException) { return null; }
    }
    private async Task AuthorizeAsync(string destination, bool allowProjectSource, CancellationToken token) {
        _storage.EnsureWritableLocation(destination);
        bool ownProject = allowProjectSource && SameLocation(destination, _projectLocation);
        if (!ownProject && SameLocation(destination, _sourceLocation))
            throw new IOException(T("ProtectSource", "Choose another destination to preserve the opened source."));
        if (!ownProject && (_sourceIdentity != null || _projectIdentity != null)) {
            try {
                string identity = await _storage.ReadIdentityAsync(destination, token);
                if (identity == _sourceIdentity || identity == _projectIdentity)
                    throw new IOException(T("ProtectSource", "Choose another destination to preserve the opened source."));
            } catch (FileNotFoundException) { }
        }
        if (_guard != null && !await _guard.CanPublishAsync(destination, false, token))
            throw new IOException(T("Protected", "The destination is protected by another open document or recovery operation."));
    }
    [RelayCommand(CanExecute = nameof(CanEditBook))]
    private Task ChooseCoverAsync() => RunAsync(async token => {
        string? location = await _dialogs().PickOpenFileAsync(T("Cover", "Choose cover image"), new StudioFileType("Cover image", ["png", "jpg", "jpeg", "gif", "svg"]), token);
        if (location == null) return;
        byte[] bytes = (await _storage.ReadSnapshotAsync(location, token, 10L * 1024 * 1024)).Bytes;
        string mime = Path.GetExtension(_storage.Describe(location).Name).ToLowerInvariant() switch {
            ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", ".gif" => "image/gif", ".svg" => "image/svg+xml",
            _ => throw new NotSupportedException("Cover import supports PNG, JPEG, GIF and static SVG.")
        };
        await Task.Run(() => _project!.SetCoverImage(bytes, mime, token), token);
        IsDirty = true; Status = T("CoverSelected", "Package cover selected. Existing cover pages and resources are retained.");
    });
    [RelayCommand(CanExecute = nameof(CanPreview))]
    private Task PreviewChapterAsync() => RunAsync(async token => {
        int index = SelectedChapter!.Index; await ApplyDraftsAsync(token); ClearPreview();
        var results = await Task.Run(() => _project!.PreviewChapter(index, cancellationToken: token), token);
        token.ThrowIfCancellationRequested();
        if (_disposed) return;
        var result = results.FirstOrDefault() ?? throw new InvalidOperationException("No chapter preview was produced.");
        using var stream = new MemoryStream(result.Bytes!); Preview = new Bitmap(stream);
        PreviewStatus = T("PreviewHint", "Sample chapter rendered by OfficeIMO. EPUB readers can reflow the book differently.") +
            " " + string.Join(" ", result.Diagnostics.Select(item => item.Message).Distinct());
    });
}
