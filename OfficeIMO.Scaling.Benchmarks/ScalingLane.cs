namespace OfficeIMO.Scaling.Benchmarks;

/// <summary>
/// One format operation measured at several sizes. <see cref="Setup"/> builds
/// the input for <c>units</c> items outside the timed region; <see cref="Run"/>
/// performs the operation once; <see cref="Validate"/> checks the last result so
/// a lane can't get faster by doing less work.
/// </summary>
public abstract class ScalingLane {
    public abstract string Name { get; }

    /// <summary>What one unit is (paragraph, row, shape, attachment, ...).</summary>
    public abstract string Unit { get; }

    public int Units { get; private set; }

    public void Setup(int units) {
        Units = units;
        Prepare(units);
    }

    protected virtual void Prepare(int units) {
    }

    public abstract object? Run();

    public virtual void Validate(object? result) {
        if (result == null) {
            throw new InvalidOperationException($"{Name} at {Units} {Unit}s returned no result.");
        }
    }

    protected void Require(bool condition, string detail) {
        if (!condition) {
            throw new InvalidOperationException($"{Name} at {Units} {Unit}s: {detail}");
        }
    }
}
