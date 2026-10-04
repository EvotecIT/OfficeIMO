using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class ProviderSaveDialog : StudioDialogWindow {
    internal ProviderSaveDialog(string name, IStudioLocalizer localizer, bool workflowOutput = false, bool folderOutput = false) : base(new ProviderSaveDialogContent(name, localizer, workflowOutput, folderOutput)) { }
}
