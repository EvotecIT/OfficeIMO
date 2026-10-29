using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Conversion;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// Document-level structure and page production: properties, the outline, attachments, cropping and page sizes, page
/// numbers, headers and footers, watermarks and Bates numbers, flattening forms and markup, and outputs written beside the
/// document such as split parts, saved attachments, text exports and conversions to Office documents and HTML. Each edit is one undoable change in the shared workspace.
/// </summary>
internal sealed partial class StudioNativeDocument {
    /// <summary>The document's properties, outline and attachments, with which of them this document lets the host change.</summary>
    internal async Task<StudioNativeStructure> GetStructureAsync(CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            PdfMetadata? metadata = _workspace.Metadata;
            var properties = new StudioNativeProperties(metadata?.Title, metadata?.Author, metadata?.Subject, metadata?.Keywords,
                Timestamp(metadata?.CreationDate), Timestamp(metadata?.ModificationDate));
            int nextId = 0;
            IReadOnlyList<StudioNativeOutlineItem> outline = Outline(_workspace.DocumentInfo?.Outlines ?? [], editable: true, ref nextId);
            StudioNativeAttachment[] attachments = _workspace.Attachments
                .Select(static attachment => new StudioNativeAttachment(attachment.UnicodeFileName ?? attachment.FileName, attachment.SizeBytes,
                    attachment.Description, attachment.MimeType, AttachmentId(attachment))).ToArray();
            bool hasForms = _workspace.HasForms;
            return new StudioNativeStructure(StudioNativeAbi.Version, _workspace.Revision, properties, _workspace.CanEditMetadata, outline,
                _workspace.CanEditBookmarks, attachments, _workspace.CanEditAttachments, _workspace.CanCropPages, _workspace.CanResizePages,
                _workspace.CanEditPageContent, hasForms, hasForms && _workspace.CanFlattenForms, HasMarkup() && _workspace.CanEditAnnotations,
                _workspace.CanRedact);
        } finally { _gate.Release(); }

        static string? Timestamp(DateTimeOffset? value) => value?.ToString("o", CultureInfo.InvariantCulture);
    }

    // Numbers entries the way the engine's bookmark edit session does: resolvable entries depth-first, skipping the
    // subtrees of entries without a page, which the session drops and the host shows without an id.
    private static StudioNativeOutlineItem[] Outline(IReadOnlyList<PdfOutlineItem> items, bool editable, ref int nextId) {
        var result = new StudioNativeOutlineItem[items.Count];
        for (int index = 0; index < items.Count; index++) {
            PdfOutlineItem item = items[index];
            bool resolved = editable && item.PageNumber.HasValue;
            string? id = resolved ? "bookmark-" + (++nextId).ToString(CultureInfo.InvariantCulture) : null;
            result[index] = new StudioNativeOutlineItem(id, item.Title, item.PageNumber, Outline(item.Children, resolved, ref nextId));
        }
        return result;
    }

    /// <summary>
    /// Applies one edit as one undoable change. The request names exactly one edit: <c>properties</c>, <c>outline</c>,
    /// <c>crop</c>, <c>pageSize</c>, <c>headerFooter</c>, <c>watermark</c>, <c>bates</c>, <c>flatten</c>,
    /// <c>addAttachment</c> or <c>removeAttachment</c>.
    /// </summary>
    internal Task EditAsync(StudioNativeEditRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        int named = new object?[] { request.Properties, request.Outline, request.Crop, request.PageSize, request.HeaderFooter, request.Watermark,
            request.Bates, request.Flatten, request.AddAttachment, request.RemoveAttachment }.Count(static edit => edit is not null);
        if (named != 1) throw new ArgumentException("Name exactly one edit.", nameof(request));
        return RunAsync(() => request switch {
            { Properties: { } properties } => EditProperties(properties, cancellationToken),
            { Outline: { } outline } => EditOutline(outline, cancellationToken),
            { Crop: { } crop } => Crop(crop, cancellationToken),
            { PageSize: { } size } => Resize(size, cancellationToken),
            { HeaderFooter: { } headerFooter } => StampHeaderFooter(headerFooter, cancellationToken),
            { Watermark: { } watermark } => StampWatermark(watermark, cancellationToken),
            { Bates: { } bates } => StampBates(bates, cancellationToken),
            { Flatten: { } flatten } => Flatten(flatten, cancellationToken),
            { AddAttachment: { } attachment } => AddAttachment(attachment, cancellationToken),
            _ => RemoveAttachment(request.RemoveAttachment!, cancellationToken)
        }, cancellationToken);
    }

    private Task EditProperties(StudioNativeProperties properties, CancellationToken cancellationToken) {
        if (!_workspace.CanEditMetadata) throw new NotSupportedException("This document's properties can't be changed.");
        // Fields the request leaves out keep their value; an empty string clears one.
        PdfMetadata? current = _workspace.Metadata;
        return _workspace.UpdateMetadataAsync(properties.Title ?? current?.Title ?? string.Empty, properties.Author ?? current?.Author ?? string.Empty,
            properties.Subject ?? current?.Subject ?? string.Empty, properties.Keywords ?? current?.Keywords ?? string.Empty, cancellationToken);
    }

    private Task EditOutline(IReadOnlyList<StudioNativeOutlineEdit> edits, CancellationToken cancellationToken) {
        if (!_workspace.CanEditBookmarks) throw new NotSupportedException("This document's bookmarks can't be changed.");
        if (edits.Count == 0) throw new ArgumentException("Give at least one bookmark change.", nameof(edits));
        int count = _workspace.Pages.Count;
        var pages = new List<int>();
        foreach (StudioNativeOutlineEdit edit in edits) {
            if (edit.Page is { } page) {
                if (page < 1 || page > count) throw new ArgumentOutOfRangeException(nameof(edits), "The page does not exist.");
                pages.Add(page);
            }
        }
        string description = edits.Count == 1 ? edits[0].Op switch {
            "add" => "Added bookmark", "rename" => "Renamed bookmark", "remove" => "Removed bookmark", "move" => "Moved bookmark",
            "retarget" => "Changed bookmark page", "rebuild" => "Made bookmarks from headings", _ => "Edited bookmarks"
        } : "Edited bookmarks";
        return _workspace.EditBookmarksAsync(description, pages, session => {
            foreach (StudioNativeOutlineEdit edit in edits) {
                try { Apply(session, edit); }
                catch (KeyNotFoundException exception) {
                    throw new ArgumentException("The bookmark isn't in the outline any more; read the structure again.", nameof(edits), exception);
                }
            }
        }, cancellationToken);

        static void Apply(PdfBookmarkEditSession session, StudioNativeOutlineEdit edit) {
            switch (edit.Op) {
                case "add":
                    PdfBookmarkNode added = session.Add(RequireTitle(edit.Title), edit.Page ?? throw new ArgumentException("A new bookmark needs a page."), edit.ParentId);
                    if (edit.Index is { } position) session.Move(added.Id, edit.ParentId, position);
                    break;
                case "rename": session.Rename(RequireId(edit.Id), RequireTitle(edit.Title)); break;
                case "remove": session.Remove(RequireId(edit.Id)); break;
                case "move": session.Move(RequireId(edit.Id), edit.ParentId, edit.Index ?? -1); break;
                case "retarget": session.Retarget(RequireId(edit.Id), edit.Page ?? throw new ArgumentException("Give the bookmark's new page.")); break;
                case "rebuild":
                    session.RebuildFromHeadings();
                    if (session.Roots.Count == 0) throw new InvalidOperationException("The document has no headings to make bookmarks from.");
                    break;
                default: throw new ArgumentException($"“{edit.Op}” isn't a bookmark change.");
            }
        }

        static string RequireId(string? id) => string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Name the bookmark to change.") : id;

        static string RequireTitle(string? title) => string.IsNullOrWhiteSpace(title) ? throw new ArgumentException("A bookmark needs a title.") : title.Trim();
    }

    private Task Crop(StudioNativeCropEdit crop, CancellationToken cancellationToken) {
        if (!_workspace.CanCropPages) throw new NotSupportedException("This document's pages can't be cropped.");
        int[] pages = AllOrChosenPages(crop.Pages);
        foreach (int page in pages) {
            (double width, double height) = DisplayedSize(page);
            if (crop.Left + crop.Right >= width - 1D || crop.Top + crop.Bottom >= height - 1D)
                throw new ArgumentException($"The margins leave nothing of page {page}.", nameof(crop));
        }
        return _workspace.CropByMarginsAsync(pages, crop.Left, crop.Bottom, crop.Right, crop.Top, cancellationToken);
    }

    private (double Width, double Height) DisplayedSize(int page) {
        (double width, double height) = CurrentSession().GetDisplaySize(page);
        return (width, height);
    }

    private Task Resize(StudioNativePageSizeEdit size, CancellationToken cancellationToken) {
        if (!_workspace.CanResizePages) throw new NotSupportedException("This document's pages can't be resized.");
        if (!double.IsFinite(size.Width) || !double.IsFinite(size.Height) || size.Width < 72D || size.Height < 72D || size.Width > 14400D || size.Height > 14400D)
            throw new ArgumentOutOfRangeException(nameof(size), "A page is between 72 and 14,400 points on each side.");
        PdfPageResizeMode mode = size.Mode switch {
            null or "fit" => PdfPageResizeMode.Fit, "fill" => PdfPageResizeMode.Fill, "stretch" => PdfPageResizeMode.Stretch,
            _ => throw new ArgumentException($"“{size.Mode}” isn't a way to fit content.", nameof(size))
        };
        return _workspace.ResizePagesAsync(AllOrChosenPages(size.Pages), new PdfPageSize(OfficeIMO.OfficeLength.Points(size.Width),OfficeIMO.OfficeLength.Points(size.Height)), mode, cancellationToken);
    }

    private int[] AllOrChosenPages(IReadOnlyList<int>? pages) => pages is { Count: > 0 }
        ? RequirePages(pages) : Enumerable.Range(1, _workspace.Pages.Count).ToArray();

    private Task StampHeaderFooter(StudioNativeHeaderFooterEdit edit, CancellationToken cancellationToken) {
        RequirePageContent();
        var options = new PdfHeaderFooterOptions {
            HeaderLeft = edit.HeaderLeft, HeaderCenter = edit.HeaderCenter, HeaderRight = edit.HeaderRight,
            FooterLeft = edit.FooterLeft, FooterCenter = edit.FooterCenter, FooterRight = edit.FooterRight,
            FontSize = Bounded(edit.FontSize ?? 9D, 4D, 72D, "font size"), Margin = Bounded(edit.Margin ?? 36D, 0D, 288D, "margin"),
            StartNumber = edit.StartNumber ?? 1, PageRange = RangeOf(edit.Pages)
        };
        return _workspace.ApplyHeaderFooterAsync(options, cancellationToken);
    }

    private async Task StampWatermark(StudioNativeWatermarkEdit edit, CancellationToken cancellationToken) {
        RequirePageContent();
        if (string.IsNullOrWhiteSpace(edit.Text)) throw new ArgumentException("Enter the watermark's text.", nameof(edit));
        var options = new PdfWatermarkOptions {
            Text = edit.Text.Trim(), FontSize = Bounded(edit.FontSize ?? 42D, 6D, 288D, "font size"),
            Opacity = Bounded(edit.Opacity ?? 0.35D, 0.05D, 1D, "opacity"), RotationDegrees = Bounded(edit.Rotation ?? -35D, -180D, 180D, "angle"),
            BehindContent = edit.BehindContent
        };
        if (edit.Color is { } color) options.Color = global::OfficeIMO.OfficeColor.FromRgb(Channel(color.Red), Channel(color.Green), Channel(color.Blue));
        int[] pages = edit.Pages is { Count: > 0 } chosen ? RequirePages(chosen) : Enumerable.Range(1, _workspace.Pages.Count).ToArray();
        if (edit.Pages is { Count: > 0 }) options.TargetPages = PdfPageSelector.Parse(RangeOf(pages)!);
        PdfWatermarkPreview prepared = await _workspace.PrepareWatermarkAsync(options, pages[0], cancellationToken).ConfigureAwait(false);
        await _workspace.ApplyWatermarkAsync(prepared, cancellationToken).ConfigureAwait(false);

        static byte Channel(double value) => (byte)Math.Round(Math.Clamp(value, 0D, 1D) * 255D);
    }

    private Task StampBates(StudioNativeBatesEdit edit, CancellationToken cancellationToken) {
        RequirePageContent();
        if (edit.Start < 0) throw new ArgumentOutOfRangeException(nameof(edit), "Bates numbers start at zero or more.");
        var options = new PdfBatesNumberingOptions {
            Prefix = edit.Prefix ?? string.Empty, Suffix = edit.Suffix ?? string.Empty, StartNumber = edit.Start ?? 1L,
            MinimumDigits = (int)Bounded(edit.Digits ?? 6, 1, 15, "number of digits"),
            Position = edit.Position switch {
                null or "bottomRight" => PdfBatesPosition.BottomRight, "bottomCenter" => PdfBatesPosition.BottomCenter,
                "bottomLeft" => PdfBatesPosition.BottomLeft, "topLeft" => PdfBatesPosition.TopLeft, "topCenter" => PdfBatesPosition.TopCenter,
                "topRight" => PdfBatesPosition.TopRight, _ => throw new ArgumentException($"“{edit.Position}” isn't a Bates position.", nameof(edit))
            }
        };
        return _workspace.ApplyBatesNumberingAsync(options, cancellationToken);
    }

    private Task Flatten(StudioNativeFlattenEdit edit, CancellationToken cancellationToken) {
        if (edit.Markup == true) {
            if (edit.Fields is { Count: > 0 }) throw new ArgumentException("Flatten the markup or the form, not both.", nameof(edit));
            if (!_workspace.CanEditAnnotations) throw new NotSupportedException("This document's annotations can't be changed.");
            if (!HasMarkup()) throw new ArgumentException("This document has no markup to flatten.", nameof(edit));
            return _workspace.FlattenAllAnnotationsAsync(cancellationToken);
        }
        if (!_workspace.CanFlattenForms) throw new NotSupportedException("This document's form can't be flattened.");
        if (edit.Fields is not { Count: > 0 } fields) return _workspace.FlattenFormFieldsAsync(cancellationToken);
        if (fields.Count > 1) throw new ArgumentException("Flatten one field, or the whole form.", nameof(edit));
        return _workspace.FlattenFormFieldAsync(fields[0], cancellationToken);
    }

    private Task AddAttachment(StudioNativeAttachmentAdd attachment, CancellationToken cancellationToken) {
        if (!_workspace.CanEditAttachments) throw new NotSupportedException("This document's attachments can't be changed.");
        if (string.IsNullOrWhiteSpace(attachment.Path)) throw new ArgumentException("Choose a file to attach.", nameof(attachment));
        var file = new FileInfo(attachment.Path);
        if (!file.Exists) throw new FileNotFoundException("The file to attach wasn't found.", attachment.Path);
        if (file.Length > PdfWorkspace.MaximumAttachmentBytes) throw new ArgumentException("Attachments are up to 64 MB.", nameof(attachment));
        string name = string.IsNullOrWhiteSpace(attachment.FileName) ? file.Name : Path.GetFileName(attachment.FileName.Trim());
        return _workspace.AddAttachmentAsync(name, File.ReadAllBytes(file.FullName), cancellationToken);
    }

    private Task RemoveAttachment(string fileName, CancellationToken cancellationToken) {
        if (!_workspace.CanEditAttachments) throw new NotSupportedException("This document's attachments can't be changed.");
        return _workspace.RemoveAttachmentAsync(FindAttachment(fileName), cancellationToken);
    }

    private static string AttachmentId(PdfAttachmentInfo attachment) =>
        $"attachment:{attachment.FileSpecObjectNumber}:{attachment.Source}:{attachment.Name}";

    private PdfAttachmentInfo FindAttachment(string identifier) {
        PdfAttachmentInfo? identified = _workspace.Attachments.FirstOrDefault(item => AttachmentId(item) == identifier);
        if (identified is not null) return identified;
        PdfAttachmentInfo[] matches = _workspace.Attachments.Where(item =>
            string.Equals(item.UnicodeFileName ?? item.FileName, identifier, StringComparison.Ordinal)).ToArray();
        return matches.Length == 1 ? matches[0] : throw new ArgumentException(matches.Length == 0
            ? $"The document has no attachment “{identifier}”." : "Choose an attachment by its identifier; this filename is shared by several attachments.", nameof(identifier));
    }

    private void RequirePageContent() {
        if (!_workspace.CanEditPageContent) throw new NotSupportedException("This document's pages can't be stamped.");
    }

    private string? RangeOf(IReadOnlyList<int>? pages) =>
        pages is not { Count: > 0 } ? null : string.Join(",", RequirePages(pages).Select(static page => page.ToString(CultureInfo.InvariantCulture)));

    private static double Bounded(double value, double minimum, double maximum, string name) =>
        double.IsFinite(value) && value >= minimum && value <= maximum ? value
            : throw new ArgumentOutOfRangeException(name, $"The {name} is between {minimum} and {maximum}.");

    /// <summary>
    /// Writes an output beside the open document, which is unchanged: the request names exactly one of <c>split</c>,
    /// <c>attachment</c>, <c>export</c>, <c>optimize</c>, <c>repair</c> or <c>sanitize</c>. Returns the files written, in order.
    /// </summary>
    internal async Task<StudioNativeWriteResult> WriteAsync(StudioNativeWriteRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        int named = new object?[] { request.Split, request.Attachment, request.Export, request.Optimize, request.Repair, request.Sanitize }
            .Count(static output => output is not null);
        if (named != 1) throw new ArgumentException("Name exactly one output.", nameof(request));
        StudioNativeHealthCopy? health = request.Optimize ?? request.Repair ?? request.Sanitize;
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (health is not null) {
                health = health with { Path = _workspace.ValidateCopyDestination(health.Path!) };
                string kind = request.Optimize is not null ? "optimize" : request.Repair is not null ? "repair" : "sanitize";
                (IReadOnlyList<string> copies, string summary) = await WriteHealthCopyAsync(health, kind, cancellationToken).ConfigureAwait(false);
                return new StudioNativeWriteResult(StudioNativeAbi.Version, _workspace.Revision, copies, summary);
            }
            if (request.Export is { } export) {
                (IReadOnlyList<string> exported, string? notes) = await ExportAsync(export, cancellationToken).ConfigureAwait(false);
                return new StudioNativeWriteResult(StudioNativeAbi.Version, _workspace.Revision, exported, notes);
            }
            IReadOnlyList<string> files = request.Split is { } split
                ? await SplitAsync(split, cancellationToken).ConfigureAwait(false)
                : await SaveAttachmentAsync(request.Attachment!, cancellationToken).ConfigureAwait(false);
            return new StudioNativeWriteResult(StudioNativeAbi.Version, _workspace.Revision, files);
        } finally { _gate.Release(); }
    }

    private async Task<IReadOnlyList<string>> SplitAsync(StudioNativeSplit split, CancellationToken cancellationToken) {
        if (!_workspace.CanExtractPages) throw new NotSupportedException("This document's pages can't be split.");
        if (string.IsNullOrWhiteSpace(split.Directory)) throw new ArgumentException("Choose a folder for the parts.", nameof(split));
        Directory.CreateDirectory(split.Directory);
        int count = _workspace.Pages.Count;
        IReadOnlyList<StudioNativeSplitPart> parts;
        if (split.Parts is { Count: > 0 } given) {
            parts = given;
        } else if (split.PagesPerPart is { } size and > 0) {
            parts = Enumerable.Range(0, (count + size - 1) / size)
                .Select(index => new StudioNativeSplitPart(Enumerable.Range(index * size + 1, Math.Min(size, count - index * size)).ToArray())).ToArray();
        } else {
            throw new ArgumentException("Give the pages per part or the parts.", nameof(split));
        }
        if (parts.Count < 2) throw new ArgumentException("A split makes at least two parts.", nameof(split));
        string baseName = Path.GetFileNameWithoutExtension(_workspace.FileName);
        var written = new List<string>(parts.Count);
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < parts.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int[] pages = RequirePages(parts[index].Pages ?? []);
            string name = SafeFileName(parts[index].Name) ?? $"{baseName} {Describe(pages)}";
            while (true) {
                string path = FreePath(split.Directory, name, taken);
                try {
                    await _workspace.ExtractToLocalFileAsync(pages, path, cancellationToken, replaceExisting: false).ConfigureAwait(false);
                    written.Add(path);
                    break;
                } catch (IOException) when (File.Exists(path) || Directory.Exists(path)) {
                    // Another writer took this name after selection. Publication never replaced it; try the next name.
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }
        }
        return written;

        static string Describe(int[] pages) => pages.Length == 1 || pages[^1] - pages[0] != pages.Length - 1
            ? (pages.Length == 1 ? $"p{pages[0]}" : $"p{pages[0]}+{pages.Length - 1}")
            : $"p{pages[0]}-{pages[^1]}";
    }

    // Never replaces a file: a name already in the folder, or used by an earlier part, gets a number.
    private static string FreePath(string directory, string name, HashSet<string> taken) {
        for (int attempt = 1; ; attempt++) {
            string candidate = Path.Combine(directory, (attempt == 1 ? name : $"{name} {attempt}") + ".pdf");
            if (!File.Exists(candidate) && !Directory.Exists(candidate) && taken.Add(candidate)) return candidate;
        }
    }

    private static string? SafeFileName(string? name) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        char[] invalid = Path.GetInvalidFileNameChars();
        string cleaned = new string(name.Trim().Select(character => invalid.Contains(character) || character is '/' or '\\' or ':' ? '-' : character).ToArray()).Trim('.', ' ');
        return cleaned.Length == 0 ? null : cleaned.Length > 120 ? cleaned[..120] : cleaned;
    }

    private async Task<IReadOnlyList<string>> SaveAttachmentAsync(StudioNativeAttachmentSave save, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(save.Path)) throw new ArgumentException("A destination is required.", nameof(save));
        await _workspace.SaveAttachmentAsync(FindAttachment(save.Id ?? save.FileName ?? string.Empty), save.Path, cancellationToken).ConfigureAwait(false);
        return [save.Path];
    }

    /// <summary>
    /// Writes the open revision as text, Markdown, JSON, Word, Excel, PowerPoint or HTML at the path, or its embedded
    /// images into the folder the path names. Returns the files written and the converter's notes, one per line.
    /// </summary>
    private async Task<(IReadOnlyList<string> Files, string? Notes)> ExportAsync(StudioNativeExport export, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(export.Path)) throw new ArgumentException("A destination is required.", nameof(export));
        string format = export.Format ?? string.Empty;
        bool allowed = format == "text" ? _workspace.ViewInfo.CanExtractText : _workspace.ViewInfo.CanExtractContent;
        if (!allowed) throw new NotSupportedException("This document doesn't allow its content to be copied.");
        if (PdfConversion.PdfTargets.ContainsKey(format)) {
            PdfConversionResult result = await PdfConversion.ConvertPdfAsync(_workspace.CopyBytes(), _workspace.ReadOptions?.Password, format,
                Path.GetFullPath(export.Path), cancellationToken).ConfigureAwait(false);
            string? notes = result.Notes.Count == 0 ? null : string.Join("\n", result.Notes.Select(static note => note.Message).Distinct(StringComparer.Ordinal));
            return ([export.Path], notes);
        }
        if (format == "images") {
            Directory.CreateDirectory(export.Path);
            var before = new HashSet<string>(Directory.EnumerateFiles(export.Path), StringComparer.Ordinal);
            await _workspace.ExportDocumentAsync(PdfExportKind.Images, export.Path, cancellationToken).ConfigureAwait(false);
            return (Directory.EnumerateFiles(export.Path).Where(file => !before.Contains(file)).Order(StringComparer.Ordinal).ToArray(), null);
        }
        PdfExportKind kind = format switch {
            "text" => PdfExportKind.Text, "markdown" => PdfExportKind.Markdown, "json" => PdfExportKind.Json,
            _ => throw new ArgumentException($"“{export.Format}” isn't an export format.", nameof(export))
        };
        await _workspace.ExportDocumentAsync(kind, export.Path, cancellationToken).ConfigureAwait(false);
        return ([export.Path], null);
    }

    // Markup is any annotation but links, form widgets and popups, which flattening leaves alone.
    private bool HasMarkup() =>
        _workspace.AnnotationMetadata.Any(static annotation => annotation.Subtype is not ("Link" or "Widget" or "Popup"));
}

internal sealed record StudioNativeProperties(string? Title, string? Author, string? Subject, string? Keywords, string? Created = null,
    string? Modified = null);

/// <summary>An outline entry; entries without an id have no page and can't be changed, nor can their children.</summary>
internal sealed record StudioNativeOutlineItem(string? Id, string Title, int? Page, IReadOnlyList<StudioNativeOutlineItem> Children);

internal sealed record StudioNativeAttachment(string FileName, long Size, string? Description, string? MimeType, string Id);

internal sealed record StudioNativeStructure(int AbiVersion, long Revision, StudioNativeProperties Properties, bool CanEditProperties,
    IReadOnlyList<StudioNativeOutlineItem> Outline, bool CanEditOutline, IReadOnlyList<StudioNativeAttachment> Attachments,
    bool CanEditAttachments, bool CanCrop, bool CanResize, bool CanStampPages, bool HasForms, bool CanFlattenForms, bool CanFlattenMarkup,
    bool CanRedact);

/// <summary>One bookmark change: <c>op</c> is add, rename, remove, move, retarget or rebuild.</summary>
internal sealed record StudioNativeOutlineEdit(string Op, string? Id = null, string? Title = null, int? Page = null, string? ParentId = null,
    int? Index = null);

/// <summary>Margins in points cut from each edge of the pages as displayed.</summary>
internal sealed record StudioNativeCropEdit(IReadOnlyList<int>? Pages, double Left, double Top, double Right, double Bottom);

/// <summary>A new paper size in points; <c>mode</c> is fit (default), fill or stretch.</summary>
internal sealed record StudioNativePageSizeEdit(IReadOnlyList<int>? Pages, double Width, double Height, string? Mode = null);

/// <summary>Header and footer text; templates accept {page}, {pages}, {date}, {file} and {title}. No pages means every page.</summary>
internal sealed record StudioNativeHeaderFooterEdit(string? HeaderLeft = null, string? HeaderCenter = null, string? HeaderRight = null,
    string? FooterLeft = null, string? FooterCenter = null, string? FooterRight = null, double? FontSize = null, double? Margin = null,
    int? StartNumber = null, IReadOnlyList<int>? Pages = null);

internal sealed record StudioNativeWatermarkEdit(string? Text, double? FontSize = null, double? Opacity = null, double? Rotation = null,
    bool BehindContent = false, StudioNativeColor? Color = null, IReadOnlyList<int>? Pages = null);

/// <summary>Bates numbers on every page; <c>position</c> is bottomRight (default), bottomCenter, bottomLeft, topLeft, topCenter or topRight.</summary>
internal sealed record StudioNativeBatesEdit(string? Prefix = null, string? Suffix = null, long? Start = null, int? Digits = null,
    string? Position = null);

/// <summary>Flattens the named field, or every field when none is named; or, with markup, every markup annotation.</summary>
internal sealed record StudioNativeFlattenEdit(IReadOnlyList<string>? Fields = null, bool? Markup = null);

internal sealed record StudioNativeAttachmentAdd(string? Path, string? FileName = null);

internal sealed record StudioNativeEditRequest(StudioNativeProperties? Properties = null, IReadOnlyList<StudioNativeOutlineEdit>? Outline = null,
    StudioNativeCropEdit? Crop = null, StudioNativePageSizeEdit? PageSize = null, StudioNativeHeaderFooterEdit? HeaderFooter = null,
    StudioNativeWatermarkEdit? Watermark = null, StudioNativeBatesEdit? Bates = null, StudioNativeFlattenEdit? Flatten = null,
    StudioNativeAttachmentAdd? AddAttachment = null, string? RemoveAttachment = null);

internal sealed record StudioNativeSplitPart(IReadOnlyList<int>? Pages, string? Name = null);

/// <summary>Parts written into <c>directory</c>: every <c>pagesPerPart</c> pages, or the listed parts, such as one per top-level bookmark.</summary>
internal sealed record StudioNativeSplit(string? Directory, int? PagesPerPart = null, IReadOnlyList<StudioNativeSplitPart>? Parts = null);

internal sealed record StudioNativeAttachmentSave(string? FileName, string? Path, string? Id = null);

/// <summary>The document's text as <c>text</c>, <c>markdown</c> or structured <c>json</c>.</summary>
internal sealed record StudioNativeExport(string? Format, string? Path);

internal sealed record StudioNativeWriteRequest(StudioNativeSplit? Split = null, StudioNativeAttachmentSave? Attachment = null,
    StudioNativeExport? Export = null, StudioNativeHealthCopy? Optimize = null, StudioNativeHealthCopy? Repair = null,
    StudioNativeHealthCopy? Sanitize = null);

/// <summary>The files written, in order, and for a health copy what changed.</summary>
internal sealed record StudioNativeWriteResult(int AbiVersion, long Revision, IReadOnlyList<string> Files, string? Summary = null);
