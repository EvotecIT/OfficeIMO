namespace OfficeIMO.IWork.Internal;

/// <summary>Tracks the bounded non-tiered decimal levels shared by owned models and native encoding.</summary>
internal sealed class IWorkDecimalCreationCounter {
    private readonly int?[] _numbers = new int?[IWorkDecimalListProfile.MaximumLevel + 1];

    internal bool HasParent(int level) => _numbers[level - 1].HasValue;
    internal int? Next(int level) => _numbers[level] == int.MaxValue ? null : (_numbers[level] ?? 0) + 1;
    internal void Accept(int level, int number) {
        _numbers[level] = number;
        for (int child = level + 1; child < _numbers.Length; child++) _numbers[child] = null;
    }
    internal void Reset() { Array.Clear(_numbers, 0, _numbers.Length); }
}
