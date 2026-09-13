namespace OfficeIMO.Project;

/// <summary>Shared mapped-record operations over generation-specific storage.</summary>
internal interface IProjectNativeTableEditor : IDisposable {
    bool Contains(int uid);
    bool HasField(uint id);
    IEnumerable<int> Uids { get; }
    void Integer(int uid, uint id, int value);
    void Set(int uid, uint id, byte[]? value);
    void Add(int uid);
    void AddReserved(int index);
    void Delete(int uid);
    void Export(Dictionary<string, byte[]> replacements);
}
