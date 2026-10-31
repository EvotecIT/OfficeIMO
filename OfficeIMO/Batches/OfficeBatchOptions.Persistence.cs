namespace OfficeIMO;

public sealed partial record OfficeBatchOptions {
    /// <summary>Reads a bounded saved definition. Relative paths resolve against the file's directory; secrets remain unresolved.</summary>
    public static OfficeBatchOptions Load(string path) => OfficeBatchOptionsSerializer.ReadFile(path);
    /// <summary>Reads a bounded saved definition using an explicit absolute base directory.</summary>
    public static OfficeBatchOptions Parse(string json, string baseDirectory) => OfficeBatchOptionsSerializer.Deserialize(json, baseDirectory);
    /// <summary>Saves a portable definition atomically. Runtime passwords and advanced settings are rejected.</summary>
    public void Save(string path, bool overwrite = false) => OfficeBatchOptionsSerializer.SaveFile(this, path, overwrite);
    /// <summary>Serializes portable settings using paths relative to an explicit absolute job directory.</summary>
    public string ToJson(string baseDirectory) => OfficeBatchOptionsSerializer.Serialize(this, baseDirectory);
}
