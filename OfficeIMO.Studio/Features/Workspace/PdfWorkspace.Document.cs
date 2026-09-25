using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>Document-level structure: properties, bookmarks, attachments, headers and footers, and exports.</summary>
internal sealed partial class PdfWorkspace {
    internal const long MaximumAttachmentBytes = 64L * 1024 * 1024;
    internal bool CanEditMetadata => CanPlan(PdfMutationOperation.UpdateMetadata);

    internal bool CanEditBookmarks => CanPlan(PdfMutationOperation.ModifyCatalog);

    internal bool CanEditAttachments => CanPlan(PdfMutationOperation.ModifyAttachments);

    internal PdfMetadata? Metadata => DocumentInfo?.Metadata;

    internal IReadOnlyList<PdfAttachmentInfo> Attachments => DocumentInfo?.Attachments ?? [];

    internal Task UpdateMetadataAsync(string title, string author, string subject, string keywords,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        // Empty strings clear a field; the engine keeps the Info dictionary and XMP packet in step.
        MutateBytesAsync(PdfWorkspaceOperationKind.Metadata, "Updated document properties", [],
            bytes => LoadDocument(bytes).UpdateMetadata(title, author, subject, keywords).ToBytes(),
            cancellationToken, progress);

    internal Task EditBookmarksAsync(string description, IReadOnlyList<int> pages, Action<PdfBookmarkEditSession> edit,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.Bookmarks, description, pages,
            bytes => LoadDocument(bytes).Bookmarks.Edit(edit).ToBytes(), cancellationToken, progress);

    internal Task AddAttachmentAsync(string fileName, byte[] data, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (Attachments.Any(item => string.Equals(item.FileName, fileName, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException($"An attachment named {fileName} already exists. Remove it first or rename the file.");
        var file = new PdfEmbeddedFile(fileName, data, GuessMimeType(fileName), modificationDate: DateTimeOffset.Now);
        return MutateBytesAsync(PdfWorkspaceOperationKind.Attachments, "Attached " + fileName, [],
            bytes => RequireValid(LoadDocument(bytes).Attachments.Add(file)).ToBytes(), cancellationToken, progress);
    }

    internal Task RemoveAttachmentAsync(PdfAttachmentInfo attachment, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.Attachments, "Removed attachment " + attachment.FileName, [],
            bytes => RequireValid(LoadDocument(bytes).Attachments.Remove(attachment)).ToBytes(), cancellationToken, progress);

    internal async Task SaveAttachmentAsync(PdfAttachmentInfo attachment, string destination, CancellationToken cancellationToken) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        byte[] payload = await RunCancellableCpuWorkAsync(() => snapshot.Attachments.Extract(
                attachment, MaximumAttachmentBytes, cancellationToken).Bytes,
            cancellationToken).ConfigureAwait(false);
        await WriteWorkspaceOutputAsync(destination,
            (stream, token) => stream.WriteAsync(payload.AsMemory(), token).AsTask(), cancellationToken).ConfigureAwait(false);
    }

    internal Task ApplyHeaderFooterAsync(PdfHeaderFooterOptions options, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!options.HasContent) throw new ArgumentException("Enter text for at least one header or footer position.", nameof(options));
        string fileName = System.IO.Path.GetFileNameWithoutExtension(FileName);
        string title = Metadata?.Title is { Length: > 0 } documentTitle ? documentTitle : fileName;
        string date = DateTime.Now.ToString("d", CultureInfo.CurrentCulture);
        int pageCount = Pages.Count;
        int[] targets = options.ResolvePages(pageCount);
        var stampOptions = new PdfCanvasStampOptions();
        if (!string.IsNullOrWhiteSpace(options.PageRange)) stampOptions.UseTargetPages(options.PageRange!.Trim());
        return MutateBytesAsync(PdfWorkspaceOperationKind.HeaderFooter, "Added header and footer", targets,
            bytes => LoadDocument(bytes).Stamp.Content((canvas, context) => {
                string Expand(string template) => template
                    .Replace("{page}", (context.PageNumber + options.StartNumber - 1).ToString(CultureInfo.CurrentCulture), StringComparison.OrdinalIgnoreCase)
                    .Replace("{pages}", (context.PageCount + options.StartNumber - 1).ToString(CultureInfo.CurrentCulture), StringComparison.OrdinalIgnoreCase)
                    .Replace("{date}", date, StringComparison.OrdinalIgnoreCase)
                    .Replace("{file}", fileName, StringComparison.OrdinalIgnoreCase)
                    .Replace("{title}", title, StringComparison.OrdinalIgnoreCase);
                double width = Math.Max(1D, context.Width - options.Margin * 2D);
                double lineHeight = options.FontSize * 1.6D;
                double headerY = options.Margin * 0.6D;
                double footerY = Math.Max(0D, context.Height - options.Margin * 0.6D - lineHeight);
                void Draw(string? template, double y, PdfAlign align) {
                    if (string.IsNullOrWhiteSpace(template)) return;
                    canvas.Text(Expand(template), options.Margin, y, width, lineHeight, fontSize: options.FontSize,
                        color: PdfColor.FromRgb(71, 84, 103), align: align);
                }
                Draw(options.HeaderLeft, headerY, PdfAlign.Left);
                Draw(options.HeaderCenter, headerY, PdfAlign.Center);
                Draw(options.HeaderRight, headerY, PdfAlign.Right);
                Draw(options.FooterLeft, footerY, PdfAlign.Left);
                Draw(options.FooterCenter, footerY, PdfAlign.Center);
                Draw(options.FooterRight, footerY, PdfAlign.Right);
            }, stampOptions).ToBytes(),
            cancellationToken, progress);
    }

    /// <summary>Changes the paper size of pages, scaling their content to fit, fill or stretch.</summary>
    internal Task ResizePagesAsync(IReadOnlyList<int> pageNumbers, PageSize size, PdfPageResizeMode mode, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.Resize, "Resized pages", pageNumbers,
            bytes => LoadDocument(bytes).Pages.Resize(new PdfPageResizeOptions(size) { Mode = mode }, pageNumbers.ToArray()).ToBytes(),
            cancellationToken, progress);

    internal async Task<int> ExportDocumentAsync(PdfExportKind kind, string destination, CancellationToken cancellationToken,
        IReadOnlyDictionary<string, PdfFormFieldValue>? formDrafts = null) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        if (kind == PdfExportKind.Images) {
            IReadOnlyList<PdfExtractedImage> images = await RunCancellableCpuWorkAsync(
                () => PdfReadDocument.Open(snapshot.ToBytes(), _readOptions).ExtractImages().Where(image => image.IsImageFile).ToArray(),
                cancellationToken).ConfigureAwait(false);
            if (images.Count == 0) throw new InvalidOperationException("This document has no embedded images that can be saved as image files.");
            destination = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(destination)
                ?? throw new System.IO.IOException("Choose a folder on this computer to save the images.");
            System.IO.Directory.CreateDirectory(destination);
            string baseName = OfficeIMO.Core.Internal.OfficePortableFileName.SanitizeBaseName(
                System.IO.Path.GetFileNameWithoutExtension(FileName), maximumLength: 80);
            if (baseName.Length == 0) baseName = "document";
            string[] names = new string[images.Count];
            for (int batch = 0; ; batch++) {
                if (batch == int.MaxValue) throw new IOException("No available image export names remain in this folder.");
                string prefix = batch == 0 ? baseName : string.Create(CultureInfo.InvariantCulture, $"{baseName}-{batch}");
                for (int index = 0; index < images.Count; index++) {
                    PdfExtractedImage image = images[index];
                    names[index] = string.Create(CultureInfo.InvariantCulture,
                        $"{prefix}-p{image.PageNumber}-{index + 1}{image.FileExtension ?? ".bin"}");
                }
                if (names.All(name => !System.IO.File.Exists(System.IO.Path.Combine(destination, name)) &&
                                      !System.IO.Directory.Exists(System.IO.Path.Combine(destination, name)))) break;
            }
            for (int index = 0; index < images.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfExtractedImage image = images[index];
                byte[] payload = image.Bytes;
                await WriteWorkspaceOutputAsync(System.IO.Path.Combine(destination, names[index]),
                    (stream, token) => stream.WriteAsync(payload.AsMemory(), token).AsTask(), cancellationToken,
                    conflictPolicy: OfficeIMO.Core.Internal.OfficeFileCommit.ConflictPolicy.FailIfExists).ConfigureAwait(false);
            }
            return images.Count;
        }
        string content = await RunCancellableCpuWorkAsync(() => kind switch {
            PdfExportKind.Markdown => snapshot.Read(cancellationToken: cancellationToken).ExportStructured(PdfStructuredExportFormat.Markdown),
            PdfExportKind.Json => snapshot.Read(cancellationToken: cancellationToken).ExportStructured(PdfStructuredExportFormat.Json),
            PdfExportKind.Text => snapshot.Read(cancellationToken: cancellationToken).Text,
            PdfExportKind.FormData => ExportFormData(snapshot, formDrafts),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        }, cancellationToken).ConfigureAwait(false);
        byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content);
        await WriteWorkspaceOutputAsync(destination,
            (stream, token) => stream.WriteAsync(bytes.AsMemory(), token).AsTask(), cancellationToken).ConfigureAwait(false);
        return 1;
    }

    private static string ExportFormData(PdfDocument snapshot, IReadOnlyDictionary<string, PdfFormFieldValue>? formDrafts) {
        if (formDrafts is { Count: > 0 }) {
            PdfMutationPlan plan = snapshot.PlanMutation(PdfMutationOperation.FillFormFields, formDrafts.Keys);
            snapshot = plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
                ? snapshot.Forms.AppendRevision(formDrafts) : snapshot.Forms.Fill(formDrafts);
        }
        return snapshot.Forms.ExportXfdf();
    }

    /// <summary>Checks the current revision against a PDF/A, PDF/UA or PDF/X profile without changing it.</summary>
    internal Task<PdfComplianceReadinessReport> AssessComplianceAsync(PdfComplianceProfile profile, CancellationToken cancellationToken) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        return RunCancellableCpuWorkAsync(() => snapshot.AssessCompliance(profile), cancellationToken);
    }

    internal async Task ImportFormDataAsync(string source, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        var snapshot = await _storage.ReadSnapshotAsync(source, cancellationToken,
            PdfFormDataSet.DefaultMaxXfdfDocumentBytes).ConfigureAwait(false);
        PdfFormDataSet data = await RunCancellableCpuWorkAsync(
            () => PdfFormDataSet.ParseXfdfBytes(snapshot.Bytes), cancellationToken).ConfigureAwait(false);
        await MutateBytesAsync(PdfWorkspaceOperationKind.FormFill, "Imported form data from " + _storage.Describe(source).Name, [],
            bytes => LoadDocument(bytes).Forms.ImportData(data).ToBytes(), cancellationToken, progress).ConfigureAwait(false);
    }

    internal Task WriteTextOutputAsync(string destination, string content, CancellationToken cancellationToken) {
        byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content);
        return WriteWorkspaceOutputAsync(destination, (stream, token) => stream.WriteAsync(bytes.AsMemory(), token).AsTask(), cancellationToken);
    }

    private static PdfAttachmentEditResult RequireValid(PdfAttachmentEditResult result) =>
        result.IsValid ? result : throw new InvalidOperationException("The attachment could not be verified after the change.");

    private static string? GuessMimeType(string fileName) => System.IO.Path.GetExtension(fileName).ToLowerInvariant() switch {
        ".pdf" => "application/pdf",
        ".txt" => "text/plain",
        ".csv" => "text/csv",
        ".xml" => "application/xml",
        ".json" => "application/json",
        ".png" => "image/png",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".zip" => "application/zip",
        _ => null
    };
}

internal enum PdfExportKind {
    Markdown,
    Text,
    Json,
    Images,
    FormData
}

/// <summary>Header and footer text per position. Templates accept {page}, {pages}, {date}, {file} and {title}.</summary>
internal sealed record PdfHeaderFooterOptions {
    public string? HeaderLeft { get; init; }
    public string? HeaderCenter { get; init; }
    public string? HeaderRight { get; init; }
    public string? FooterLeft { get; init; }
    public string? FooterCenter { get; init; }
    public string? FooterRight { get; init; }
    public double FontSize { get; init; } = 9D;
    public double Margin { get; init; } = 36D;
    public int StartNumber { get; init; } = 1;
    public string? PageRange { get; init; }

    public bool HasContent => new[] { HeaderLeft, HeaderCenter, HeaderRight, FooterLeft, FooterCenter, FooterRight }
        .Any(value => !string.IsNullOrWhiteSpace(value));

    public int[] ResolvePages(int pageCount) => string.IsNullOrWhiteSpace(PageRange)
        ? Enumerable.Range(1, pageCount).ToArray()
        : PdfPageSelector.Parse(PageRange!.Trim()).Resolve(pageCount).ToArray();
}
