using OfficeIMO.Internal;
namespace OfficeIMO;

public sealed partial record OfficeBatchReport {
    /// <summary>Serializes bounded observational state. The report never authorizes checkpoint reuse.</summary>
    public string ToJson() => OfficeBatchReportStore.Serialize(this);
    /// <summary>Reads strict bounded JSON without checking local writer activity.</summary>
    public static OfficeBatchReport Parse(string json) => OfficeBatchReportStore.Deserialize(json);
    /// <summary>Reads a saved report, recovering abandoned unfinished records as Interrupted.</summary>
    public static OfficeBatchReport Load(string path) => OfficeBatchReportStore.ReadFile(path);
    /// <summary>Writes a new observational report atomically, preserving any existing destination.</summary>
    public void Save(string path) {
        string fullPath = Path.GetFullPath(path);
        OfficeBatchExecutor.EnsureNoLinks(fullPath);
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(ToJson());
        OfficeIMO.Core.Internal.OfficeFileCommit.WriteAtomically(fullPath, stream => {
            OfficeRuntimeCompatibility.WriteAll(stream, bytes); if (stream is FileStream file) file.Flush(true);
        }, conflictPolicy: OfficeIMO.Core.Internal.OfficeFileCommit.ConflictPolicy.FailIfExists);
    }
    /// <summary>Checks the local report writer lease without treating timestamps as activity evidence.</summary>
    public static bool IsActive(string path) => OfficeBatchReportStore.IsActive(path);
    /// <summary>Removes only an inactive valid recorded report and its matching lease.</summary>
    internal static bool TryDeleteFinished(string path) => OfficeBatchReportStore.TryDeleteFinished(path);
}
