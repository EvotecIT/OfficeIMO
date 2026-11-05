namespace OfficeIMO.CSV.Benchmarks;

public readonly record struct CsvReadObservation(
    int Rows,
    int Cells,
    long Characters,
    ulong Checksum);

internal struct CsvObservationAccumulator {
    private const ulong ChecksumOffset = 14695981039346656037UL;
    private const ulong ChecksumPrime = 1099511628211UL;
    private int _cells;
    private long _characters;
    private ulong _checksum;

    internal void Add(string? value) => Add(value.AsSpan());

    internal void Add(ReadOnlySpan<char> value) {
        if (_cells == 0) {
            _checksum = ChecksumOffset;
        }

        _cells++;
        _characters += value.Length;
        foreach (char character in value) {
            _checksum ^= character;
            _checksum *= ChecksumPrime;
        }

        _checksum ^= (ulong)value.Length;
        _checksum *= ChecksumPrime;
    }

    internal CsvReadObservation ToObservation(int rows) =>
        new(rows, _cells, _characters, _cells == 0 ? ChecksumOffset : _checksum);
}
