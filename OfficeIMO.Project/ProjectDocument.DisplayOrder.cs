namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Assigns native display IDs in current outline/resource order, preserving reserved UID-zero rows.</summary>
    public void NormalizeDisplayIds() {
        EnsureNotDisposed();
        using var update = BeginUpdate();
        int displayId = 0;
        foreach (var task in AllTasks) task.DisplayId = task.Uid == 0 ? 0 : ++displayId;
        displayId = 0;
        foreach (var resource in Resources) resource.DisplayId = resource.Uid == 0 ? 0 : ++displayId;
    }
}
