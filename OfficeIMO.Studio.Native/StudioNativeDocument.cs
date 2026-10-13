using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Security;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Scan;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// One open document behind a native handle. It wraps the same workspace the desktop Studio uses, so editing, undo,
/// identity checks and verified saving behave identically in every host. Calls on one document are serialized; the
/// host may call from any thread and should not call from its main thread.
/// </summary>
internal sealed partial class StudioNativeDocument : IDisposable {
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly PdfWorkspace _workspace;
    private PdfDocumentSession? _session;
    private long _sessionRevision = -1;
    private bool _disposed;

    private StudioNativeDocument(PdfWorkspace workspace) => _workspace = workspace;

    /// <summary>
    /// Opens a PDF the host can currently read. On Apple platforms the host resolves the security-scoped URL and keeps
    /// access open for the duration of each call that reads or writes the file.
    /// </summary>
    internal static async Task<StudioNativeDocument> OpenAsync(string path, string? password, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A document path is required.", nameof(path));
        PdfWorkspace workspace = await PdfWorkspace.OpenAsync(path, cancellationToken, StudioNativeEngine.CreateRecoveryStore(), password).ConfigureAwait(false);
        return new StudioNativeDocument(workspace);
    }

    internal async Task<StudioNativeDocumentInfo> GetInfoAsync(CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            PdfDocumentViewInfo view = _workspace.ViewInfo;
            return new StudioNativeDocumentInfo(
                StudioNativeAbi.Version, _workspace.FileName, _workspace.Path, _workspace.FileSize, _workspace.Revision,
                _workspace.IsDirty, _workspace.CanUndo, _workspace.CanRedo, _workspace.RecoveryPath is not null, view.CanExtractText, view.CanExtractContent,
                _workspace.Pages.Select(page => {
                    (double width, double height) = PdfDocumentSession.GetPhysicalSize(page);
                    return new StudioNativePageInfo(page.PageNumber, width, height, page.RotationDegrees);
                }).ToArray(),
                StudioPdfSecurityPolicy.MaximumRasterPixels, _workspace.CanMutatePages, _workspace.CanExtractPages, _workspace.CanEditAnnotations,
                !view.Security.HasEncryption || view.Security.HasOwnerAuthorization || view.Security.AllowsPrinting == true,
                _workspace.CanImportPages, _workspace.CanEditPageContent && view.CanExtractText, _workspace.CanFillForms,
                _workspace.IsEncrypted, _workspace.CanChangeEncryption(null), _workspace.CanSign);
        } finally { _gate.Release(); }
    }

    /// <summary>
    /// Renders one page as PNG. Scale is pixels per PDF point; the host applies its display scale and keeps the page
    /// within <see cref="StudioPdfSecurityPolicy.MaximumRasterPixels"/>, refused otherwise as an invalid argument.
    /// </summary>
    internal async Task<PdfRenderedPage> RenderPageAsync(int pageNumber, double scale, CancellationToken cancellationToken) {
        if (!double.IsFinite(scale) || scale <= 0D) throw new ArgumentOutOfRangeException(nameof(scale), "The scale must be a positive number.");
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            return await CurrentSession().RenderPageAsync(pageNumber, scale, cancellationToken).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    /// <summary>Renders a rectangle of a page as opaque RGBA pixels; see <see cref="PdfDocumentSession.RenderRegionAsync"/>.</summary>
    internal async Task<PdfRenderedPixels> RenderRegionAsync(int pageNumber, double scale, int x, int y, int width, int height,
        CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            return await CurrentSession().RenderRegionAsync(pageNumber, scale, x, y, width, height, cancellationToken).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    internal Task RotatePagesAsync(IReadOnlyList<int> pageNumbers, int degrees, CancellationToken cancellationToken) {
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        if (degrees % 90 != 0) throw new ArgumentOutOfRangeException(nameof(degrees), "Pages rotate in quarter turns.");
        return RunAsync(() => _workspace.RotateAsync(pageNumbers, degrees, cancellationToken), cancellationToken);
    }

    /// <summary>Deletes pages as one undoable change. A document keeps at least one page.</summary>
    internal Task DeletePagesAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken) =>
        RunAsync(() => {
            int[] pages = RequirePages(pageNumbers);
            if (pages.Length >= _workspace.Pages.Count) throw new ArgumentException("A document keeps at least one page.", nameof(pageNumbers));
            return _workspace.DeleteAsync(pages, cancellationToken);
        }, cancellationToken);

    /// <summary>Duplicates pages as one undoable change; each copy follows its page.</summary>
    internal Task DuplicatePagesAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.DuplicateAsync(RequirePages(pageNumbers), cancellationToken), cancellationToken);

    /// <summary>Puts the pages in a new order as one undoable change: <paramref name="order"/> lists every current page once.</summary>
    internal Task ReorderPagesAsync(IReadOnlyList<int> order, CancellationToken cancellationToken) =>
        RunAsync(() => {
            int count = _workspace.Pages.Count;
            if (order.Count != count || order.Distinct().Count() != count || order.Any(page => page < 1 || page > count))
                throw new ArgumentException("The new order must list every page exactly once.", nameof(order));
            return _workspace.ReorderAsync(order, cancellationToken);
        }, cancellationToken);

    /// <summary>Inserts a blank page of the given size in points before a page, or after the last page with count + 1.</summary>
    internal Task InsertBlankPageAsync(int insertBefore, double width, double height, CancellationToken cancellationToken) {
        if (!double.IsFinite(width) || !double.IsFinite(height) || width < 3D || height < 3D || width > 14400D || height > 14400D)
            throw new ArgumentOutOfRangeException(nameof(width), "A page is between 3 and 14,400 points on each side.");
        return RunAsync(() => {
            if (insertBefore < 1 || insertBefore > _workspace.Pages.Count + 1) throw new ArgumentOutOfRangeException(nameof(insertBefore), "The position is outside the document.");
            return _workspace.InsertBlankAsync(insertBefore, width, height, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>Writes pages to a new PDF at <paramref name="path"/>; the open document is unchanged.</summary>
    internal Task ExtractPagesAsync(IReadOnlyList<int> pageNumbers, string path, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination is required.", nameof(path));
        // The local path keeps the workflow runner, and the conversions it roots, out of the NativeAOT library.
        return RunAsync(() => _workspace.ExtractToLocalFileAsync(RequirePages(pageNumbers), path, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// The form fields with a widget on a page: each field's kind, value, options and flags, and its widgets' boxes in
    /// points of the page as displayed, with the state that turns a check box or radio button on.
    /// </summary>
    internal async Task<StudioNativeFormFields> GetFormFieldsAsync(int pageNumber, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            PdfDocument snapshot = _workspace.CreateDocumentSnapshot();
            var boxes = new Dictionary<int, StudioNativeRectangle>();
            foreach (PdfPageInteractionRegion region in snapshot.Render.Interactions(pageNumber).Regions) {
                if (region.Kind == PdfInteractionKind.FormWidget && region.ObjectNumber is { } number)
                    boxes[number] = new StudioNativeRectangle(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height);
            }
            var fields = new List<StudioNativeFormField>();
            foreach (PdfFormField field in snapshot.Inspect().FormFields) {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrEmpty(field.Name)) continue;
                var widgets = field.Widgets
                    .Where(widget => widget.PageNumber == pageNumber && !widget.IsHidden && !widget.IsNoView && widget.ObjectNumber is { } number && boxes.ContainsKey(number))
                    .Select(widget => {
                        string? on = widget.NormalAppearanceStates.FirstOrDefault(static state => !string.Equals(state, "Off", StringComparison.Ordinal));
                        return new StudioNativeFormWidget(widget.ObjectNumber!.Value, boxes[widget.ObjectNumber!.Value], on,
                            on is not null && string.Equals(widget.AppearanceState, on, StringComparison.Ordinal));
                    }).ToArray();
                if (widgets.Length == 0) continue;
                fields.Add(new StudioNativeFormField(field.Name, KindOf(field), field.Values.ToArray(),
                    field.Options.Select(static option => new StudioNativeFormOption(option.ExportValue, option.DisplayText)).ToArray(),
                    field.IsReadOnly, field.IsRequired, field.IsMultiline, field.IsPassword, field.IsComb, field.MaxLength,
                    field.AllowsMultipleSelection, field.IsEditableChoice, field.AlternateName, widgets));
            }
            return new StudioNativeFormFields(StudioNativeAbi.Version, _workspace.Revision, pageNumber, _workspace.CanFillForms, fields);
        } finally { _gate.Release(); }
    }

    /// <summary>
    /// Sets form field values as one undoable change: text, the selected options of a choice, or the on state of a check
    /// box or radio button ("Off" to clear it). A signed or certified document takes the values as an appended revision.
    /// </summary>
    internal Task FillFormAsync(StudioNativeFormFill request, CancellationToken cancellationToken) {
        IReadOnlyDictionary<string, IReadOnlyList<string>> values = request.Values is { Count: > 0 } given
            ? given : throw new ArgumentException("Give at least one field value.", nameof(request));
        return RunAsync(() => {
            if (!_workspace.CanFillForms) throw new NotSupportedException("This document's form can't be filled.");
            var fields = _workspace.CreateDocumentSnapshot().Inspect().FormFields
                .Where(static field => !string.IsNullOrEmpty(field.Name)).GroupBy(static field => field.Name!).ToDictionary(static group => group.Key, static group => group.First());
            var fill = new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal);
            foreach ((string name, IReadOnlyList<string> fieldValues) in values) {
                if (!fields.TryGetValue(name, out PdfFormField? field)) throw new ArgumentException($"The form has no field “{name}”.", nameof(request));
                if (field.IsReadOnly) throw new ArgumentException($"The field “{name}” is read-only.", nameof(request));
                if (field.IsSignatureField || field.IsPushButton) throw new ArgumentException($"The field “{name}” doesn't take a value.", nameof(request));
                fill[name] = PdfFormFieldValue.FromValues(fieldValues ?? []);
            }
            return _workspace.FillFormFieldsAsync(fill, cancellationToken);
        }, cancellationToken);
    }

    private static string KindOf(PdfFormField field) =>
        field.IsTextField ? "text" : field.IsCheckBox ? "checkbox" : field.IsRadioButton ? "radio" : field.IsPushButton ? "button"
        : field.IsChoiceField ? (field.IsCombo ? "combo" : "list") : field.IsSignatureField ? "signature" : "other";

    /// <summary>
    /// Recognizes each page's text with the host's recognizer and adds it as invisible text over the page, as one undoable
    /// change. The engine renders each page at <see cref="SearchableDpi"/> and passes the PNG and its page number to
    /// <paramref name="recognize"/>, which returns the words on it. Words over text a page already has are left out.
    /// </summary>
    internal async Task<StudioNativeSearchableResult> MakeSearchableAsync(Func<int, byte[], CancellationToken, IReadOnlyList<PdfScanWord>> recognize,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(recognize);
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (!_workspace.ViewInfo.CanExtractText) throw new NotSupportedException("This document doesn't allow its text to be read.");
            var options = new PdfOcrMergeOptions { Dpi = SearchableDpi, MaxPages = _workspace.Pages.Count, ReconstructLayout = false };
            PdfWorkspaceSearchableResult result = await _workspace.MakeSearchableAsync(HostTextRecognizer.Create(recognize), options, cancellationToken)
                .ConfigureAwait(false);
            return new StudioNativeSearchableResult(StudioNativeAbi.Version, result.ModifiedPages, result.Words, _workspace.Revision);
        } finally { _gate.Release(); }
    }

    /// <summary>The number of pages in the current revision.</summary>
    internal int PageCount => _workspace.Pages.Count;

    /// <summary>The resolution pages are recognized at: enough for body text, within the render budget on a phone.</summary>
    internal const double SearchableDpi = 200D;

    /// <summary>
    /// Inserts pages from other PDFs before <see cref="StudioNativeImportRequest.InsertBefore"/>, in the order given,
    /// as one undoable change. A source without pages contributes all of them. A protected source is opened with its
    /// password; when it is missing or wrong the call fails with a <see cref="PdfImportPasswordException"/> naming it.
    /// </summary>
    internal async Task<StudioNativeImportResult> ImportPagesAsync(StudioNativeImportRequest request, CancellationToken cancellationToken) {
        IReadOnlyList<StudioNativeImportSource> sources = request.Sources is { Count: > 0 } list
            ? list : throw new ArgumentException("Choose at least one PDF to import.", nameof(request));
        if (sources.Any(source => string.IsNullOrWhiteSpace(source.Path)))
            throw new ArgumentException("Each PDF to import needs a path.", nameof(request));
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (!_workspace.CanImportPages) throw new NotSupportedException("Pages can't be added to this document safely.");
            PdfImportPreparation preparation = await _workspace.PrepareImportAsync(sources.Select(source => source.Path!).ToArray(),
                cancellationToken, promptPassword: null, passwords: sources.Select(source => source.Password).ToArray()).ConfigureAwait(false)
                ?? throw new OperationCanceledException(cancellationToken);
            var selections = new PdfImportSelection[sources.Count];
            for (int index = 0; index < sources.Count; index++) {
                int count = preparation.Sources[index].PageCount;
                int[] pages = sources[index].Pages is { Count: > 0 } chosen ? chosen.ToArray() : Enumerable.Range(1, count).ToArray();
                if (pages.Any(page => page < 1 || page > count))
                    throw new ArgumentOutOfRangeException(nameof(request), $"“{preparation.Sources[index].Name}” has {count} pages.");
                selections[index] = new PdfImportSelection(index, pages);
            }
            if (request.InsertBefore < 1 || request.InsertBefore > _workspace.Pages.Count + 1)
                throw new ArgumentOutOfRangeException(nameof(request), "The position is outside the document.");
            int imported = await _workspace.ApplyImportAsync(preparation, selections, request.InsertBefore, cancellationToken).ConfigureAwait(false);
            return new StudioNativeImportResult(StudioNativeAbi.Version, imported, _workspace.Revision);
        } finally { _gate.Release(); }
    }

    /// <summary>
    /// Finds text, ignoring case. Each hit's rectangles are in points of the page as displayed: from its top-left corner,
    /// after rotation and user unit, the same space as the page sizes the host lays out.
    /// </summary>
    /// <remarks><paramref name="pageSearched"/> receives each page's hits as soon as the page is searched.</remarks>
    internal async Task<StudioNativeSearchResult> SearchAsync(string query, CancellationToken cancellationToken,
        Action<StudioNativeSearchPage>? pageSearched = null) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            PdfDocumentSession session = CurrentSession();
            if (!session.CanSearch) throw new NotSupportedException("Text search is restricted by this document's permissions.");
            string needle = query.Trim();
            long revision = _workspace.Revision;
            IReadOnlyList<PdfSearchHit> hits = await session.SearchAsync(query, cancellationToken, pageSearched: pageSearched is null ? null
                : (searched, count, pageHits) => pageSearched(new StudioNativeSearchPage(StudioNativeAbi.Version, needle, revision, searched, count,
                    pageHits.Select(ToNative).ToArray()))).ConfigureAwait(false);
            return new StudioNativeSearchResult(StudioNativeAbi.Version, needle, revision, hits.Select(ToNative).ToArray());
        } finally { _gate.Release(); }

        static StudioNativeSearchHit ToNative(PdfSearchHit hit) => new(hit.PageNumber, hit.OccurrenceNumber, hit.Snippet,
            StudioNativeRectangle.From(hit.Bounds), hit.LineBounds.Select(StudioNativeRectangle.From).ToArray());
    }

    // The markup tools a host may add. Text, images and redaction have their own journeys.
    private static readonly HashSet<PdfEditorTool> AnnotationTools = [
        PdfEditorTool.Ink, PdfEditorTool.Note, PdfEditorTool.FreeText, PdfEditorTool.Highlight, PdfEditorTool.Underline,
        PdfEditorTool.StrikeOut, PdfEditorTool.Squiggly, PdfEditorTool.Rectangle, PdfEditorTool.Ellipse, PdfEditorTool.Line,
        PdfEditorTool.Polygon, PdfEditorTool.Link, PdfEditorTool.Stamp
    ];

    // The standard stamp names of ISO 32000-2, 12.5.6.12, which every reader draws.
    private static readonly HashSet<string> StampNames = new(StringComparer.Ordinal) {
        "Approved", "Experimental", "NotApproved", "AsIs", "Expired", "NotForPublicRelease", "Confidential", "Final", "Sold",
        "Departmental", "ForComment", "TopSecret", "Draft", "ForPublicRelease"
    };

    // The line endings of ISO 32000-2, 12.5.6.7, that an arrow uses.
    private static readonly string[] LineEndings = ["None", "OpenArrow", "ClosedArrow"];

    /// <summary>
    /// Adds a markup annotation as one undoable change. Geometry is in points of the page as displayed, from its
    /// top-left corner; the workspace maps it to the page's own space through its box, rotation and user unit.
    /// </summary>
    internal Task AddAnnotationAsync(StudioNativeAnnotationRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        if (!Enum.TryParse(request.Tool, ignoreCase: true, out PdfEditorTool tool) || !AnnotationTools.Contains(tool))
            throw new ArgumentException($"'{request.Tool}' is not a markup tool.", nameof(request));
        var strokes = (request.Strokes ?? []).Select(stroke => (IReadOnlyList<PdfEditorVisualPoint>)stroke
            .Select(point => new PdfEditorVisualPoint(Finite(point.X), Finite(point.Y))).ToArray()).Where(stroke => stroke.Count > 0).ToArray();
        var quads = (request.Quads ?? []).Select(quad => new PdfEditorVisualBounds(Finite(quad.X), Finite(quad.Y),
            Finite(quad.X + quad.Width), Finite(quad.Y + quad.Height))).ToArray();
        double? width = request.StrokeWidth;
        if (width is { } w && (!double.IsFinite(w) || w <= 0D || w > 144D)) throw new ArgumentOutOfRangeException(nameof(request), "A stroke is wider than 0 and at most 144 points.");
        if (request.Opacity is { } o && (!double.IsFinite(o) || o <= 0D || o > 1D)) throw new ArgumentOutOfRangeException(nameof(request), "Opacity is above 0 and at most 1.");
        (double left, double top, double right, double bottom) = request.Bounds is { } bounds
            ? (Finite(bounds.X), Finite(bounds.Y), Finite(bounds.X + bounds.Width), Finite(bounds.Y + bounds.Height))
            : BoundsOf(strokes.SelectMany(static stroke => stroke), (width ?? 1D) / 2D + 1D);
        if (right <= left || bottom <= top) throw new ArgumentException("The markup needs an area or a drawn path.", nameof(request));
        StudioNativeColor color = request.Color ?? new StudioNativeColor(1D, 0.85D, 0D);
        if (new[] { color.Red, color.Green, color.Blue }.Any(static part => !double.IsFinite(part) || part < 0D || part > 1D))
            throw new ArgumentOutOfRangeException(nameof(request), "Colour components are from 0 through 1.");
        var gesture = new PdfEditorGesture(request.Page, left, top, right, bottom, strokes.FirstOrDefault() ?? [],
            strokes.Length > 0 ? strokes : null, quads.Length > 0 ? quads : null);
        string linkUri = string.Empty;
        if (tool == PdfEditorTool.Link) {
            if (request.LinkPage is null == string.IsNullOrWhiteSpace(request.LinkUri))
                throw new ArgumentException("A link goes to a web address or to a page, not both.", nameof(request));
            if (request.LinkUri is { } uri) {
                if (!Uri.TryCreate(uri.Trim(), UriKind.Absolute, out Uri? parsed) || parsed.Scheme is not ("https" or "http" or "mailto"))
                    throw new ArgumentException("A link's address starts with https://, http:// or mailto:.", nameof(request));
                linkUri = parsed.AbsoluteUri;
            }
        }
        string stamp = string.Empty;
        if (tool == PdfEditorTool.Stamp) {
            stamp = request.Stamp ?? "Draft";
            if (!StampNames.Contains(stamp)) throw new ArgumentException($"“{stamp}” isn't a standard stamp.", nameof(request));
        }
        if (request.LineEnd is { } ending && (tool != PdfEditorTool.Line || !LineEndings.Contains(ending)))
            throw new ArgumentException($"A line ends with {string.Join(", ", LineEndings)}; other markup has no line ending.", nameof(request));
        var properties = new PdfEditorProperties(request.Text ?? string.Empty, request.Author ?? string.Empty,
            new PdfColor(color.Red, color.Green, color.Blue), stamp, linkUri, request.FontSize ?? 12D,
            LinkPageNumber: request.LinkPage, StrokeWidth: width, Opacity: request.Opacity, LineEndEnding: request.LineEnd);
        return RunAsync(() => {
            if (request.Page < 1 || request.Page > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(request), "The page does not exist.");
            if (request.LinkPage is { } target && (target < 1 || target > _workspace.Pages.Count))
                throw new ArgumentOutOfRangeException(nameof(request), "The link's page does not exist.");
            if (!_workspace.CanEditAnnotations) throw new NotSupportedException("This document's annotations can't be changed.");
            return _workspace.ApplyEditorGestureAsync(tool, gesture, properties, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>The markup annotations on a page, with rectangles in points of the page as displayed.</summary>
    internal async Task<StudioNativeAnnotationList> GetAnnotationsAsync(int pageNumber, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            var metadata = _workspace.AnnotationMetadata.Where(static annotation => annotation.ObjectNumber.HasValue)
                .GroupBy(static annotation => annotation.ObjectNumber!.Value).ToDictionary(static group => group.Key, static group => group.First());
            PdfDocument snapshot = _workspace.CreateDocumentSnapshot();
            PdfPageInteractionMap map = snapshot.Annotations.GetEditingInteractions(pageNumber);
            var annotations = new List<StudioNativeAnnotation>();
            foreach (PdfPageInteractionRegion region in map.Regions) {
                cancellationToken.ThrowIfCancellationRequested();
                if (region.Kind != PdfInteractionKind.Annotation || region.ObjectNumber is not { } number) continue;
                metadata.TryGetValue(number, out PdfAnnotation? annotation);
                StudioNativeColor? color = annotation?.Color is { Count: 3 } rgb ? new StudioNativeColor(rgb[0], rgb[1], rgb[2]) : null;
                annotations.Add(new StudioNativeAnnotation(number, pageNumber, region.Subtype ?? annotation?.Subtype ?? string.Empty,
                    new StudioNativeRectangle(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height),
                    annotation?.Contents, annotation?.Title, color));
            }
            // Links aren't markup, so the editing map leaves them out; they are listed so they can be found and removed.
            PdfPageLayoutInfo? layout = null;
            foreach (PdfAnnotation link in metadata.Values.Where(annotation => annotation.PageNumber == pageNumber && annotation.Subtype == "Link" && annotation.ObjectNumber.HasValue &&
                         annotation.Width > 0 && annotation.Height > 0)) {
                layout ??= snapshot.GetPageLayouts()[pageNumber - 1];
                PdfSelectionQuad quad = layout.MapUserSpaceRectangleToVisual(link.X1, link.Y1, link.X2, link.Y2);
                annotations.Add(new StudioNativeAnnotation(link.ObjectNumber!.Value, pageNumber, "Link",
                    new StudioNativeRectangle(quad.Left, quad.Top, quad.Width, quad.Height), link.Contents, link.Title, null));
            }
            return new StudioNativeAnnotationList(StudioNativeAbi.Version, _workspace.Revision, pageNumber, _workspace.CanEditAnnotations, annotations);
        } finally { _gate.Release(); }
    }

    /// <summary>Changes an annotation's text, author or colour as one undoable change; what is not given is kept.</summary>
    internal Task UpdateAnnotationAsync(StudioNativeAnnotationUpdate update, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(update);
        return RunAsync(() => {
            PdfAnnotation current = _workspace.AnnotationMetadata.FirstOrDefault(annotation => annotation.ObjectNumber == update.ObjectNumber)
                ?? throw new ArgumentException("The annotation does not exist.", nameof(update));
            PdfColor color = update.Color is { } given ? new PdfColor(given.Red, given.Green, given.Blue)
                : current.Color is { Count: 3 } rgb ? new PdfColor(rgb[0], rgb[1], rgb[2]) : new PdfColor(1D, 0.85D, 0D);
            return _workspace.UpdateAnnotationAsync(update.ObjectNumber, update.Contents ?? current.Contents ?? string.Empty,
                update.Author ?? current.Title ?? string.Empty, color, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>Removes an annotation as one undoable change.</summary>
    internal Task RemoveAnnotationAsync(int objectNumber, CancellationToken cancellationToken) =>
        RunAsync(() => {
            if (!_workspace.AnnotationMetadata.Any(annotation => annotation.ObjectNumber == objectNumber))
                throw new ArgumentException("The annotation does not exist.", nameof(objectNumber));
            return _workspace.RemoveAnnotationAsync(objectNumber, cancellationToken);
        }, cancellationToken);

    private static double Finite(double value) =>
        double.IsFinite(value) ? value : throw new ArgumentOutOfRangeException(nameof(value), "Coordinates must be finite numbers.");

    private static (double Left, double Top, double Right, double Bottom) BoundsOf(IEnumerable<PdfEditorVisualPoint> points, double margin) {
        double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
        foreach (PdfEditorVisualPoint point in points) {
            left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
        }
        return left > right ? (0D, 0D, 0D, 0D) : (left - margin, top - margin, right + margin, bottom + margin);
    }

    /// <summary>
    /// A page's text elements for selection and copying: each element's text, its box, as x, y, width and height in
    /// points of the page as displayed, four numbers per element in the same order, and how it follows the element before
    /// it as a <see cref="PdfTextBreak"/>, so words and lines are found as the reader finds them.
    /// </summary>
    internal async Task<StudioNativePageText> GetPageTextAsync(int pageNumber, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            PdfDocumentSession session = CurrentSession();
            IReadOnlyList<PdfPageInteractionRegion> regions = await Task.Run(() => session.GetTextRegions(pageNumber, cancellationToken), cancellationToken)
                .ConfigureAwait(false);
            var texts = new string[regions.Count];
            var boxes = new double[regions.Count * 4];
            var breaks = new int[regions.Count];
            for (int index = 0; index < regions.Count; index++) {
                breaks[index] = (int)(index == 0 ? PdfTextBreak.Line : PdfPageInteractionMap.GetTextBreak(regions[index - 1], regions[index]));
                PdfSelectionQuad quad = regions[index].Quad;
                texts[index] = regions[index].Text ?? string.Empty;
                boxes[index * 4] = Math.Round(quad.Left, 2);
                boxes[index * 4 + 1] = Math.Round(quad.Top, 2);
                boxes[index * 4 + 2] = Math.Round(quad.Width, 2);
                boxes[index * 4 + 3] = Math.Round(quad.Height, 2);
            }
            return new StudioNativePageText(StudioNativeAbi.Version, _workspace.Revision, pageNumber, texts, boxes, breaks);
        } finally { _gate.Release(); }
    }

    internal Task UndoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.UndoAsync(cancellationToken), cancellationToken);

    internal Task RedoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.RedoAsync(cancellationToken), cancellationToken);

    /// <summary>Saves in place when <paramref name="path"/> is null, otherwise to that location, which becomes the document's location.</summary>
    internal Task SaveAsync(string? path, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.SaveAsync(path, cancellationToken), cancellationToken);

    /// <summary>
    /// Writes the open revision with a password to open it, or without the password protection it has when
    /// <c>UserPassword</c> is empty. The document itself is unchanged, as a protected export is.
    /// </summary>
    internal async Task SaveProtectedCopyAsync(StudioNativeProtectionRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.Path)) throw new ArgumentException("A destination is required.", nameof(request));
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            OfficeWorkflowResult result;
            if (string.IsNullOrEmpty(request.UserPassword)) {
                if (!_workspace.IsEncrypted) throw new ArgumentException("This PDF has no password to remove.", nameof(request));
                result = await _workspace.SaveDecryptedCopyAsync(request.Path, request.CurrentOwnerPassword, cancellationToken).ConfigureAwait(false);
            } else {
                PdfStandardPermissions allowed = PdfStandardPermissions.All;
                if (!request.AllowPrinting) allowed &= ~(PdfStandardPermissions.Print | PdfStandardPermissions.HighQualityPrint);
                if (!request.AllowCopying) allowed &= ~PdfStandardPermissions.CopyContents;
                if (!request.AllowChanges) {
                    allowed &= ~(PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations |
                        PdfStandardPermissions.FillForms | PdfStandardPermissions.AssembleDocument);
                }
                var encryption = new PdfStandardEncryptionOptions(request.UserPassword) {
                    OwnerPassword = string.IsNullOrEmpty(request.OwnerPassword) ? null : request.OwnerPassword,
                    AllowedPermissions = allowed
                };
                result = await _workspace.SaveProtectedCopyAsync(request.Path, encryption, request.CurrentOwnerPassword, cancellationToken)
                    .ConfigureAwait(false);
            }
            ThrowUnlessCompleted(result, cancellationToken);
        } finally { _gate.Release(); }
    }

    /// <summary>
    /// Writes the open revision with a certificate signature from a PKCS #12 file, so a reader can verify who signed it
    /// and that it hasn't changed since. The document itself is unchanged.
    /// </summary>
    internal async Task SaveSignedCopyAsync(StudioNativeSigningRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.Path)) throw new ArgumentException("A destination is required.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.CertificatePath)) throw new ArgumentException("A certificate is required.", nameof(request));
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (!_workspace.CanSign) throw new NotSupportedException("This document's security policy prevents signing.");
            using X509Certificate2 certificate = LoadSigningCertificate(request.CertificatePath, request.CertificatePassword);
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            var verifier = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default);
            var options = new PdfExternalSignatureOptions {
                Name = certificate.GetNameInfo(X509NameType.SimpleName, forIssuer: false),
                Reason = string.IsNullOrWhiteSpace(request.Reason) ? null : request.Reason.Trim(),
                Location = string.IsNullOrWhiteSpace(request.Location) ? null : request.Location.Trim()
            };
            OfficeWorkflowResult result = await _workspace.SaveSignedCopyAsync(request.Path, signer, options, verifier, cancellationToken)
                .ConfigureAwait(false);
            ThrowUnlessCompleted(result, cancellationToken);
        } finally { _gate.Release(); }
    }

    private static X509Certificate2 LoadSigningCertificate(string path, string? password) {
        X509Certificate2 certificate;
        try {
            certificate = X509CertificateLoader.LoadPkcs12FromFile(path, password, X509KeyStorageFlags.Exportable);
        } catch (System.Security.Cryptography.CryptographicException exception) {
            throw new StudioNativeCertificatePasswordException(exception);
        }
        DateTime now = DateTime.Now;
        if (!certificate.HasPrivateKey || certificate.NotBefore > now || certificate.NotAfter < now) {
            certificate.Dispose();
            throw new ArgumentException("The certificate has no private key or isn't valid today.", nameof(path));
        }
        return certificate;
    }

    private static void ThrowUnlessCompleted(OfficeWorkflowResult result, CancellationToken cancellationToken) {
        switch (result.Status) {
            case OfficeWorkflowStatus.Completed: return;
            case OfficeWorkflowStatus.Cancelled: throw new OperationCanceledException(cancellationToken);
            default:
                string message = result.Diagnostics
                    .FirstOrDefault(static diagnostic => diagnostic.Severity == OfficeWorkflowDiagnosticSeverity.Error)?.Message ?? result.Summary;
                // A preservation report lists every difference it found; people need the outcome, not the list.
                if (message.StartsWith("PDF rewrite preservation failed", StringComparison.Ordinal)) {
                    message = "The copy wasn't written: OfficeIMO couldn't confirm it would keep everything this document contains, such as its accessibility structure.";
                }
                throw new InvalidOperationException(message);
        }
    }

    /// <summary>Writes the open revision to <paramref name="path"/>; the document keeps its location and saved state.</summary>
    internal Task SaveCopyAsync(string path, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.SaveCopyAsync(path, cancellationToken), cancellationToken);

    // A reader session is a snapshot of one revision, with its rendering caches; edits create a new one on the next render.
    private PdfDocumentSession CurrentSession() {
        if (_session is null || _sessionRevision != _workspace.Revision) {
            _session = PdfDocumentSession.FromWorkspace(_workspace);
            _sessionRevision = _workspace.Revision;
        }
        return _session;
    }

    private int[] RequirePages(IReadOnlyList<int> pageNumbers) {
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        int count = _workspace.Pages.Count;
        if (pageNumbers.Any(page => page < 1 || page > count)) throw new ArgumentOutOfRangeException(nameof(pageNumbers), "The page does not exist.");
        return pageNumbers.Distinct().Order().ToArray();
    }

    private async Task RunAsync(Func<Task> operation, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try { await operation().ConfigureAwait(false); }
        finally { _gate.Release(); }
    }

    private async Task EnterAsync(CancellationToken cancellationToken) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        if (_disposed) {
            _gate.Release();
            throw new ObjectDisposedException(nameof(StudioNativeDocument));
        }
    }

    public void Dispose() {
        if (_disposed) return;
        // Wait for a running call so the workspace is never released underneath it.
        _gate.Wait();
        try {
            _disposed = true;
            // Closing discards unsaved changes, so their recovery copy goes too. A host that ends without closing keeps
            // it for the next launch, and so does a failed deletion, which close cannot report.
            if (_workspace.IsDirty) {
                try { _workspace.DiscardRecoveryAsync().GetAwaiter().GetResult(); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            }
            _workspace.Dispose();
        } finally { _gate.Release(); }
    }
}

internal sealed record StudioNativeDocumentInfo(
    int AbiVersion, string FileName, string Path, long FileSize, long Revision, bool IsDirty, bool CanUndo, bool CanRedo, bool HasRecovery,
    bool CanExtractText, bool CanExtractContent, IReadOnlyList<StudioNativePageInfo> Pages, long MaxRenderPixels, bool CanMutatePages,
    bool CanExtractPages, bool CanEditAnnotations, bool CanPrint, bool CanImportPages, bool CanMakeSearchable, bool CanFillForms,
    bool IsEncrypted, bool CanChangeProtection, bool CanSign);

internal sealed record StudioNativePageInfo(int Number, double Width, double Height, int Rotation);

internal sealed record StudioNativeSearchResult(int AbiVersion, string Query, long Revision, IReadOnlyList<StudioNativeSearchHit> Hits);

/// <summary>One searched page's hits while a search runs: <c>PagesSearched</c> of <c>PageCount</c> are done.</summary>
internal sealed record StudioNativeSearchPage(int AbiVersion, string Query, long Revision, int PagesSearched, int PageCount,
    IReadOnlyList<StudioNativeSearchHit> Hits);

/// <summary>A PKCS #12 file that its password doesn't open.</summary>
internal sealed class StudioNativeCertificatePasswordException(Exception inner)
    : Exception("The certificate's password is incorrect, or the file isn't a certificate with its private key.", inner);

/// <summary>A protected or unprotected copy: an empty <c>UserPassword</c> removes the password the document has.</summary>
internal sealed record StudioNativeProtectionRequest(string? Path, string? UserPassword, string? OwnerPassword, string? CurrentOwnerPassword,
    bool AllowPrinting = true, bool AllowCopying = true, bool AllowChanges = true);

/// <summary>A certificate-signed copy from a PKCS #12 file and its password.</summary>
internal sealed record StudioNativeSigningRequest(string? Path, string? CertificatePath, string? CertificatePassword, string? Reason,
    string? Location);

internal sealed record StudioNativeSearchHit(int Page, int Occurrence, string Text, StudioNativeRectangle Bounds, IReadOnlyList<StudioNativeRectangle> LineBounds);

internal sealed record StudioNativeRectangle(double X, double Y, double Width, double Height) {
    internal static StudioNativeRectangle From(StudioRectangle rectangle) => new(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
}

internal sealed record StudioNativePoint(double X, double Y);

internal sealed record StudioNativeColor(double Red, double Green, double Blue);

internal sealed record StudioNativeAnnotationRequest(string Tool, int Page, StudioNativeRectangle? Bounds = null,
    IReadOnlyList<IReadOnlyList<StudioNativePoint>>? Strokes = null, IReadOnlyList<StudioNativeRectangle>? Quads = null,
    string? Text = null, string? Author = null, StudioNativeColor? Color = null, double? StrokeWidth = null, double? Opacity = null,
    double? FontSize = null, string? LinkUri = null, int? LinkPage = null, string? Stamp = null, string? LineEnd = null);

internal sealed record StudioNativeAnnotation(int ObjectNumber, int Page, string Subtype, StudioNativeRectangle Bounds, string? Contents,
    string? Author, StudioNativeColor? Color);

internal sealed record StudioNativeAnnotationList(int AbiVersion, long Revision, int Page, bool CanEdit, IReadOnlyList<StudioNativeAnnotation> Annotations);

internal sealed record StudioNativeAnnotationUpdate(int ObjectNumber, string? Contents = null, string? Author = null, StudioNativeColor? Color = null);

internal sealed record StudioNativeFormWidget(int ObjectNumber, StudioNativeRectangle Bounds, string? OnState, bool IsOn);

internal sealed record StudioNativeFormOption(string Value, string Text);

internal sealed record StudioNativeFormField(string Name, string Kind, IReadOnlyList<string> Values, IReadOnlyList<StudioNativeFormOption> Options,
    bool ReadOnly, bool Required, bool Multiline, bool Password, bool Comb, int? MaxLength, bool MultipleSelection, bool Editable, string? Label,
    IReadOnlyList<StudioNativeFormWidget> Widgets);

internal sealed record StudioNativeFormFields(int AbiVersion, long Revision, int Page, bool CanFill, IReadOnlyList<StudioNativeFormField> Fields);

internal sealed record StudioNativeFormFill(IReadOnlyDictionary<string, IReadOnlyList<string>>? Values);

internal sealed record StudioNativeSearchableResult(int AbiVersion, int Pages, int Words, long Revision);

internal sealed record StudioNativePageText(int AbiVersion, long Revision, int Page, IReadOnlyList<string> Texts, IReadOnlyList<double> Boxes,
    IReadOnlyList<int> Breaks);
