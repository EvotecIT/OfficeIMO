using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class PdfPasswordDialog : StudioDialogWindow {
    internal PdfPasswordDialog(string documentName, bool invalidPassword, IStudioLocalizer? localizer = null) : base(new PdfPasswordDialogContent(documentName, invalidPassword, localizer)) { }
}
