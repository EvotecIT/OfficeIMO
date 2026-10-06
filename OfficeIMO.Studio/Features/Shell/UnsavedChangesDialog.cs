using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal enum UnsavedChangesDecision {
    Cancel,
    Discard,
    Save
}

internal sealed class UnsavedChangesDialog : StudioDialogWindow {
    internal UnsavedChangesDialog(string documentName, IStudioLocalizer? localizer = null) : base(new UnsavedChangesDialogContent(documentName, localizer)) { }
}
