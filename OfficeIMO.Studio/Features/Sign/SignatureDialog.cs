using Avalonia.Controls;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Sign;

/// <summary>Desktop presentation of the shared signature editor.</summary>
public sealed class SignatureDialog : StudioDialogWindow {
    private readonly SignatureDialogContent _view;
    public SignatureDialog() : this(new SignatureDialogContent()) { }
    internal SignatureDialog(StudioSignatureKind kind, IStudioLocalizer localizer) : this(new SignatureDialogContent(kind, localizer)) { }
    private SignatureDialog(SignatureDialogContent view) : base(view) => _view = view;
    internal TextBox NameBox => _view.NameBox;
    internal SignaturePad DrawingPad => _view.DrawingPad;
    internal void ShowMethod(int index) => _view.ShowMethod(index);
    internal void SetImage(byte[] image) => _view.SetImage(image);
    internal byte[]? CreateImage() => _view.CreateImage();
}
