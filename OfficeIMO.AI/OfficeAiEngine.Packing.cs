namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private sealed record PackedPrefix(int Count, OfficeAiExecutionRequest? Request);

    // Grow geometrically before refining the boundary. Every retained request is the exact
    // measured object, including its transport envelope and request number. This avoids
    // serializing every successively longer prefix of a large evidence or draft collection.
    private PackedPrefix FindFittingPrefix(int count, int maximumCharacters,
        Func<int, OfficeAiExecutionRequest?> create, CancellationToken token) {
        int low = 0, high = count;
        OfficeAiExecutionRequest? best = null;
        bool Fits(int length) {
            token.ThrowIfCancellationRequested();
            OfficeAiExecutionRequest? candidate = create(length);
            if (candidate is null) return false;
            if (!TryMeasureRequest(candidate, token, out int characters))
                throw new InvalidDataException("The executor could not measure this request.");
            if (characters > maximumCharacters) return false;
            best = candidate;
            return true;
        }
        while (low < count) {
            int next = Math.Min(count, Math.Max(1, low * 2));
            if (!Fits(next)) { high = next - 1; break; }
            low = next;
        }
        while (low < high) {
            int middle = low + (high - low + 1) / 2;
            if (Fits(middle)) low = middle;
            else high = middle - 1;
        }
        return new(low, best);
    }
}
