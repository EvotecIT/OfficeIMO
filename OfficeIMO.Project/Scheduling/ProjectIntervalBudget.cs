namespace OfficeIMO.Project;

/// <summary>Bounds retained proposals and temporary interval construction across local and external scheduling.</summary>
internal sealed class ProjectIntervalBudget {
    private readonly long _limit;
    private long _reserved;
    internal ProjectIntervalBudget(int limit) { _limit = limit; }
    internal Scope Open() => new Scope(this);

    internal sealed class Scope : IDisposable {
        private readonly ProjectIntervalBudget _owner;
        private long _reserved;
        internal Scope(ProjectIntervalBudget owner) { _owner = owner; }
        internal void Reserve() {
            if (_owner._reserved >= _owner._limit) throw new ProjectIntervalLimitException();
            _owner._reserved++; _reserved++;
        }
        internal void Keep(long retained) {
            if (retained < 0 || retained > _reserved) throw new InvalidOperationException("Invalid retained interval count.");
            _owner._reserved -= _reserved - retained;
            _reserved = 0;
        }
        public void Dispose() { _owner._reserved -= _reserved; _reserved = 0; }
    }
}

internal sealed class ProjectIntervalLimitException : InvalidOperationException {
    internal ProjectIntervalLimitException() : base("The calculated local and external assignment and cost intervals exceed MaxIntervals.") { }
}
