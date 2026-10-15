using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class ActiveOperationsDialog : StudioDialogWindow {
    internal ActiveOperationsDialog(IEnumerable<MainWindowViewModel> documents, IStudioLocalizer localizer,
        StudioDocumentTabHost? host = null, StudioSessionController? session = null, Workflows.StudioJobHistory? jobs = null)
        : base(new ActiveOperationsDialogContent(documents, localizer, host, session, jobs)) { }
}
