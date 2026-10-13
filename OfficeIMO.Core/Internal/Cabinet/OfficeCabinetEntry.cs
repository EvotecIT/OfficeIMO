namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>A bounded, decoded Cabinet member. Path safety belongs to the extracting consumer.</summary>
internal sealed class OfficeCabinetEntry {
    internal OfficeCabinetEntry(string name, byte[] data) {
        Name = name;
        Data = data;
    }

    internal string Name { get; }
    internal byte[] Data { get; }
}
