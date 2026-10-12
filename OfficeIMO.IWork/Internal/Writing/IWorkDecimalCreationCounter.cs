namespace OfficeIMO.IWork.Internal;

/// <summary>Tracks the two non-tiered decimal levels shared by owned models and native encoding.</summary>
internal sealed class IWorkDecimalCreationCounter {
    private readonly int?[] _numbers = new int?[2];

    internal bool HasParent => _numbers[0].HasValue;
    internal int? Next(int level) => _numbers[level] == int.MaxValue ? null : (_numbers[level] ?? 0) + 1;
    internal void Accept(int level, int number) {
        _numbers[level] = number;
        if (level == 0) _numbers[1] = null;
    }
    internal void Reset() { _numbers[0] = null; _numbers[1] = null; }
}
