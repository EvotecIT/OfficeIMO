namespace OfficeIMO;

/// <summary>The result of removing hidden data from encoded document bytes: what was found and done, and the cleaned bytes.</summary>
public sealed class OfficeHiddenDataResult {
    private readonly byte[] _data;

    internal OfficeHiddenDataResult(OfficeHiddenDataReport report, byte[] data) {
        Report = report;
        _data = data;
    }

    /// <summary>Every finding and what happened to it.</summary>
    public OfficeHiddenDataReport Report { get; }

    /// <summary>The length of the cleaned document in bytes.</summary>
    public long Length => _data.LongLength;

    /// <summary>Returns an owned copy of the cleaned document.</summary>
    public byte[] ToArray() => (byte[])_data.Clone();
}
