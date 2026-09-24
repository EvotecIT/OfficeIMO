using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// Fill and sign: reusable typed, drawn or image signatures and initials, and date stamps, placed with one
/// click as ordinary page content. Certificate signatures stay in Protect.
/// </summary>
public sealed partial class MainWindowViewModel {
    private Func<StudioSignatureKind, Task<StudioSignatureDraft?>> _createSignature = static _ => Task.FromResult<StudioSignatureDraft?>(null);
    private bool _signaturesLoaded;
    private byte[]? _pendingPlacementImage;
    private double _pendingPlacementAspect = 3D;
    private string? _pendingPlacementText;

    /// <summary>Opens the signature creation surface; returns null when the user cancels.</summary>
    internal Func<StudioSignatureKind, Task<StudioSignatureDraft?>> CreateSignatureDialog {
        get => _createSignature;
        set => _createSignature = value ?? (static _ => Task.FromResult<StudioSignatureDraft?>(null));
    }

    public ObservableCollection<SavedSignatureViewModel> SavedSignatures { get; } = [];

    public ObservableCollection<SavedSignatureViewModel> SavedInitials { get; } = [];

    public bool HasSavedSignatures => SavedSignatures.Count > 0;

    public bool HasSavedInitials => SavedInitials.Count > 0;

    [ObservableProperty]
    private bool _isPlacingFillSignItem;

    [ObservableProperty]
    private string? _placementHint;

    // Page content gives the most faithful result; documents that only accept annotations (for example
    // forms) still take drawn signatures as ink and typed signatures or dates as text.
    public bool CanFillAndSign => _workspace is not null && !IsWorkspaceBusy && (CanEditPageContent || CanEditAnnotations);

    /// <summary>Loads saved signatures the first time the Fill and sign menu opens.</summary>
    internal void EnsureSignaturesLoaded() {
        if (_signaturesLoaded) return;
        _signaturesLoaded = true;
        foreach (StudioSavedSignature saved in _services.Signatures.List(StudioSignatureKind.Signature)) SavedSignatures.Add(new SavedSignatureViewModel(saved));
        foreach (StudioSavedSignature saved in _services.Signatures.List(StudioSignatureKind.Initials)) SavedInitials.Add(new SavedSignatureViewModel(saved));
        NotifySavedSignatures();
    }

    private void NotifySavedSignatures() {
        OnPropertyChanged(nameof(HasSavedSignatures));
        OnPropertyChanged(nameof(HasSavedInitials));
    }

    [RelayCommand]
    private async Task CreateSignatureAsync(string? kind) {
        if (_workspace is null) return;
        EnsureSignaturesLoaded();
        StudioSignatureKind signatureKind = string.Equals(kind, nameof(StudioSignatureKind.Initials), StringComparison.OrdinalIgnoreCase)
            ? StudioSignatureKind.Initials : StudioSignatureKind.Signature;
        StudioSignatureDraft? draft = await _createSignature(signatureKind).ConfigureAwait(true);
        if (draft is null || _workspace is null) return;
        if (draft.Remember) {
            try {
                var saved = new SavedSignatureViewModel(_services.Signatures.Save(signatureKind, draft.Png, draft.Text, draft.Strokes));
                var target = signatureKind == StudioSignatureKind.Initials ? SavedInitials : SavedSignatures;
                target.Insert(0, saved);
                while (target.Count > StudioSignatureStore.MaximumPerKind) target.RemoveAt(target.Count - 1);
                NotifySavedSignatures();
            } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) {
                ErrorMessage = UiFormat("FillSign.SaveFailed", ex.Message);
            }
        }
        BeginImagePlacement(draft.Png, signatureKind, draft.Text, draft.Strokes);
    }

    [RelayCommand]
    private void PlaceSignature(SavedSignatureViewModel? signature) {
        if (signature is null || _workspace is null) return;
        BeginImagePlacement(signature.Png, signature.Kind, signature.Saved.Text, signature.Saved.Strokes);
    }

    [RelayCommand]
    private void DeleteSavedSignature(SavedSignatureViewModel? signature) {
        if (signature is null) return;
        _services.Signatures.Delete(signature.Saved);
        SavedSignatures.Remove(signature);
        SavedInitials.Remove(signature);
        NotifySavedSignatures();
    }

    [RelayCommand]
    private void PlaceDate() {
        if (_workspace is null) return;
        CancelFillSignPlacement();
        _pendingPlacementText = DateTime.Now.ToString("d", CultureInfo.CurrentCulture);
        BeginPlacement(PdfEditorTool.AddText, UiText("FillSign.PlaceDate"));
    }

    [RelayCommand]
    private void CancelFillSignPlacement() {
        bool wasPlacing = IsPlacingFillSignItem;
        _pendingPlacementImage = null;
        _pendingPlacementText = null;
        _pendingPlacementName = null;
        _pendingPlacementStrokes = null;
        IsPlacingFillSignItem = false;
        PlacementHint = null;
        if (wasPlacing && ActiveEditorTool != PdfEditorTool.Select) SelectEditorTool(nameof(PdfEditorTool.Select));
    }

    private string? _pendingPlacementName;
    private IReadOnlyList<IReadOnlyList<Avalonia.Point>>? _pendingPlacementStrokes;

    private void BeginImagePlacement(byte[] png, StudioSignatureKind kind, string? text = null, IReadOnlyList<IReadOnlyList<Avalonia.Point>>? strokes = null) {
        CancelFillSignPlacement();
        if (!CanEditPageContent && strokes is not { Count: > 0 } && string.IsNullOrWhiteSpace(text)) {
            ErrorMessage = UiText("FillSign.ImageNeedsContent");
            return;
        }
        _pendingPlacementImage = png;
        _pendingPlacementName = text;
        _pendingPlacementStrokes = strokes;
        _pendingPlacementAspect = ReadAspectRatio(png) ?? 3D;
        BeginPlacement(PdfEditorTool.AddImage, UiText(kind == StudioSignatureKind.Initials ? "FillSign.PlaceInitials" : "FillSign.PlaceSignature"));
    }

    private void BeginPlacement(PdfEditorTool tool, string hint) {
        if (DocumentMode is not StudioDocumentMode.Forms and not StudioDocumentMode.Edit and not StudioDocumentMode.Annotate) ShowFormsModeCommand.Execute(null);
        IsPlacingFillSignItem = true;
        PlacementHint = hint;
        SelectEditorTool(tool.ToString());
        UpdateFormAnchor();
    }

    private bool IsFillSignPlacement(PdfEditorTool tool) =>
        IsPlacingFillSignItem && (tool == PdfEditorTool.AddImage && _pendingPlacementImage is not null ||
                                  tool == PdfEditorTool.AddText && _pendingPlacementText is not null);

    // A click places the item at a natural size; a drawn box fits the image inside it without distortion.
    private (PdfEditorTool Tool, PdfEditorGesture Gesture, PdfEditorProperties Properties) PrepareFillSignPlacement(PdfEditorGesture gesture, PdfEditorProperties properties) {
        var ink = OfficeIMO.Pdf.PdfColor.FromRgb(27, 42, 74);
        if (_pendingPlacementText is { } text) {
            if (CanEditPageContent) return (PdfEditorTool.AddText, gesture, properties with { Text = text, FontSize = 11D, Color = ink });
            var box = gesture with { Right = gesture.Left + 110D, Bottom = gesture.Top + 22D };
            return (PdfEditorTool.FreeText, box, properties with { Text = text, FontSize = 11D, Color = ink, Author = string.Empty });
        }
        double width = gesture.Right - gesture.Left;
        double height = gesture.Bottom - gesture.Top;
        bool clicked = Math.Abs(width - 160D) < 0.5D && Math.Abs(height - 100D) < 0.5D;
        double aspect = Math.Clamp(_pendingPlacementAspect, 0.2D, 12D);
        if (clicked) { width = aspect >= 2D ? 170D : 70D; height = width / aspect; }
        else if (width / Math.Max(1D, height) > aspect) width = height * aspect;
        else height = width / aspect;
        var placed = gesture with { Right = gesture.Left + width, Bottom = gesture.Top + height };
        if (CanEditPageContent) return (PdfEditorTool.AddImage, placed, properties with { ImageBytes = _pendingPlacementImage });
        if (_pendingPlacementStrokes is { Count: > 0 } strokes) {
            // Strokes are normalised to the signature image, so they scale into the placed box directly.
            PdfEditorVisualPoint[][] mapped = strokes.Select(stroke => stroke
                .Select(point => new PdfEditorVisualPoint(placed.Left + point.X * width, placed.Top + point.Y * height)).ToArray()).ToArray();
            return (PdfEditorTool.Ink, placed with { Path = mapped.SelectMany(stroke => stroke).ToArray(), Strokes = mapped },
                properties with { Color = ink, Text = string.Empty, Author = string.Empty });
        }
        return (PdfEditorTool.FreeText, placed with { Bottom = placed.Top + Math.Max(20D, height) },
            properties with { Text = _pendingPlacementName ?? string.Empty, FontSize = Math.Clamp(height * 0.55D, 10D, 28D), Color = ink, Author = string.Empty });
    }

    private void CompleteFillSignPlacement() {
        _pendingPlacementImage = null;
        _pendingPlacementName = null;
        _pendingPlacementStrokes = null;
        _pendingPlacementText = null;
        IsPlacingFillSignItem = false;
        PlacementHint = null;
        SelectEditorTool(nameof(PdfEditorTool.Select));
    }

    // PNG and JPEG carry their pixel size in the header; anything else keeps the signature default.
    private static double? ReadAspectRatio(byte[] image) {
        if (image.Length > 24 && image[0] == 0x89 && image[1] == 0x50) {
            int w = (image[16] << 24) | (image[17] << 16) | (image[18] << 8) | image[19];
            int h = (image[20] << 24) | (image[21] << 16) | (image[22] << 8) | image[23];
            return w > 0 && h > 0 ? (double)w / h : null;
        }
        for (int i = 2; i + 9 < image.Length && image[0] == 0xFF && image[1] == 0xD8;) {
            if (image[i] != 0xFF) return null;
            byte marker = image[i + 1];
            int length = (image[i + 2] << 8) | image[i + 3];
            if (marker is >= 0xC0 and <= 0xC3) {
                int h = (image[i + 5] << 8) | image[i + 6];
                int w = (image[i + 7] << 8) | image[i + 8];
                return w > 0 && h > 0 ? (double)w / h : null;
            }
            i += 2 + length;
        }
        return null;
    }
}

/// <summary>
/// A signature or initials ready to place, and whether to keep it for next time. Drawn signatures keep their
/// strokes (normalised to the image) and typed ones their text, so they can also be placed as annotations.
/// </summary>
internal sealed record StudioSignatureDraft(byte[] Png, bool Remember, string? Text = null, IReadOnlyList<IReadOnlyList<Avalonia.Point>>? Strokes = null);

public sealed class SavedSignatureViewModel {
    internal SavedSignatureViewModel(StudioSavedSignature saved) {
        Saved = saved;
        try {
            using var stream = new MemoryStream(saved.Png);
            Preview = new Avalonia.Media.Imaging.Bitmap(stream);
        } catch (Exception) {
            Preview = null;
        }
    }

    internal StudioSavedSignature Saved { get; }
    internal StudioSignatureKind Kind => Saved.Kind;
    internal byte[] Png => Saved.Png;
    public Avalonia.Media.Imaging.Bitmap? Preview { get; }
}
