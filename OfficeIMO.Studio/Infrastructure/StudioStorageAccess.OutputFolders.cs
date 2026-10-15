using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    /// <summary>Folder outputs use the filesystem owner's atomic creation and conflict handling.</summary>
    internal static string RequireLocalOutputFolder(string location, IStudioLocalizer? localizer = null) =>
        OfficeStorageIdentity.GetLocalPath(location)
        ?? throw new IOException((localizer ?? StudioLocalization.Current).Get("Output.LocalFolderRequired"));
}
