using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Home;

/// <summary>Describes a recently opened document in the Studio shell.</summary>
public sealed partial class RecentDocumentViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;

    public RecentDocumentViewModel(string path, DateTimeOffset openedAt)
        : this(path, openedAt, StudioLocalization.Current) { }

    internal RecentDocumentViewModel(string path, DateTimeOffset openedAt, IStudioLocalizer localizer) {
        Path = OfficeIMO.Internal.OfficeStorageIdentity.Normalize(path);
        OpenedAt = openedAt;
        _localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
    }

    public string Path { get; }

    internal Infrastructure.StudioStorageReference? StorageReference { get; init; }

    public string FileName => StorageReference?.Name ?? OfficeIMO.Internal.OfficeStorageIdentity.GetFileName(Path);

    public string DirectoryName => OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(Path) is { } local
        ? System.IO.Path.GetDirectoryName(local) ?? string.Empty : new Uri(Path).GetLeftPart(UriPartial.Authority);

    public DateTimeOffset OpenedAt { get; }

    public string OpenedLabel {
        get {
            DateTimeOffset local = OpenedAt.ToLocalTime();
            DateTimeOffset now = DateTimeOffset.Now;
            if (local.Date == now.Date) return _localizer.Format("Recent.TodayAt", local.ToString("t", _localizer.Culture));
            if (local.Date == now.Date.AddDays(-1)) return _localizer.Format("Recent.YesterdayAt", local.ToString("t", _localizer.Culture));
            return local.ToString("d", _localizer.Culture);
        }
    }

    public string FileSizeLabel {
        get {
            if (OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(Path) is null) return _localizer.Get("Common.Unavailable");
            try {
                long bytes = new FileInfo(Path).Length;
                string[] units = ["B", "KB", "MB", "GB"];
                double value = bytes;
                int unit = 0;
                while (value >= 1024D && unit < units.Length - 1) {
                    value /= 1024D;
                    unit++;
                }
                return _localizer.Format("Recent.FileSize", value.ToString("0.#", _localizer.Culture), units[unit]);
            } catch (IOException) {
                return _localizer.Get("Common.Unavailable");
            } catch (UnauthorizedAccessException) {
                return _localizer.Get("Common.Unavailable");
            }
        }
    }
}
