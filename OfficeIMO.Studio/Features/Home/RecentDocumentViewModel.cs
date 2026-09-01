using CommunityToolkit.Mvvm.ComponentModel;

namespace OfficeIMO.Studio.Features.Home;

/// <summary>Describes a recently opened document in the Studio shell.</summary>
public sealed partial class RecentDocumentViewModel : ObservableObject {
    public RecentDocumentViewModel(string path, DateTimeOffset openedAt) {
        Path = System.IO.Path.GetFullPath(path);
        OpenedAt = openedAt;
    }

    public string Path { get; }

    public string FileName => System.IO.Path.GetFileName(Path);

    public string DirectoryName => System.IO.Path.GetDirectoryName(Path) ?? string.Empty;

    public DateTimeOffset OpenedAt { get; }

    public string OpenedLabel {
        get {
            DateTimeOffset local = OpenedAt.ToLocalTime();
            DateTimeOffset now = DateTimeOffset.Now;
            if (local.Date == now.Date) return $"Today, {local:HH:mm}";
            if (local.Date == now.Date.AddDays(-1)) return $"Yesterday, {local:HH:mm}";
            return local.ToString("d MMM yyyy", System.Globalization.CultureInfo.CurrentCulture);
        }
    }

    public string FileSizeLabel {
        get {
            try {
                long bytes = new FileInfo(Path).Length;
                string[] units = ["B", "KB", "MB", "GB"];
                double value = bytes;
                int unit = 0;
                while (value >= 1024D && unit < units.Length - 1) {
                    value /= 1024D;
                    unit++;
                }
                return $"{value:0.#} {units[unit]}";
            } catch (IOException) {
                return "Unavailable";
            } catch (UnauthorizedAccessException) {
                return "Unavailable";
            }
        }
    }
}
