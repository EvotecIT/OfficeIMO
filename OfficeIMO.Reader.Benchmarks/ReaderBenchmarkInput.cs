namespace OfficeIMO.Reader.Benchmarks;

internal sealed class ReaderBenchmarkInput {
    public ReaderBenchmarkInput(string name, string sourceName, byte[] bytes) {
        Name = name;
        SourceName = sourceName;
        Bytes = bytes;
    }

    public string Name { get; }
    public string SourceName { get; }
    public byte[] Bytes { get; }
}
