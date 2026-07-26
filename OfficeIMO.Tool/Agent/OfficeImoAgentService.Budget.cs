namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private static void TrimInspect(AgentInspectResult result, int maximumCharacters) {
        var folders = result.Folders.ToList();
        var metadata = result.Metadata.ToList();
        var diagnostics = result.Diagnostics.ToList();
        result.Folders = folders;
        result.Metadata = metadata;
        result.Diagnostics = diagnostics;
        while (AgentJson.Measure(result) > maximumCharacters) {
            result.Truncated = true;
            if (folders.Count > 0) {
                int remove = Math.Max(1, folders.Count / 2);
                folders.RemoveRange(folders.Count - remove, remove);
                continue;
            }
            if (metadata.Count > 0) {
                metadata.RemoveAt(metadata.Count - 1);
                continue;
            }
            if (diagnostics.Count > 0) {
                diagnostics.RemoveAt(diagnostics.Count - 1);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Preview)) {
                result.Preview = Reduce(result.Preview);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Title)) {
                result.Title = Reduce(result.Title);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Author)) {
                result.Author = Reduce(result.Author);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Subject)) {
                result.Subject = Reduce(result.Subject);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Format)) {
                result.Format = Reduce(result.Format);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Path)) {
                result.Path = Reduce(result.Path) ?? string.Empty;
                continue;
            }
            break;
        }
        EnsureWithinBudget(result, maximumCharacters);
    }

    private static void TrimSearch(
        AgentSearchResult result,
        int maximumCharacters,
        int cursor) {
        var hits = result.Results.ToList();
        result.Results = hits;
        bool removedHits = false;
        while (AgentJson.Measure(result) > maximumCharacters) {
            result.Truncated = true;
            AgentSearchHit? verbose = hits.FirstOrDefault(hit =>
                !string.IsNullOrEmpty(hit.Snippet) ||
                !string.IsNullOrEmpty(hit.Title) ||
                !string.IsNullOrEmpty(hit.Sender) ||
                !string.IsNullOrEmpty(hit.FolderId));
            if (verbose != null) {
                verbose.Snippet = null;
                verbose.Title = null;
                verbose.Sender = null;
                verbose.FolderId = null;
                continue;
            }
            if (hits.Count > 1) {
                hits.RemoveAt(hits.Count - 1);
                removedHits = true;
                continue;
            }
            if (!string.IsNullOrEmpty(result.Query)) {
                result.Query = Reduce(result.Query);
                continue;
            }
            break;
        }
        result.Returned = hits.Count;
        if (removedHits) result.NextCursor = cursor + hits.Count;
        EnsureWithinBudget(result, maximumCharacters);
    }

    private static void TrimFetch(
        AgentFetchResult result,
        int maximumCharacters,
        int cursor) {
        var metadata = result.Metadata.ToList();
        var diagnostics = result.Diagnostics.ToList();
        result.Metadata = metadata;
        result.Diagnostics = diagnostics;
        while (AgentJson.Measure(result) > maximumCharacters) {
            result.Truncated = true;
            int excess = AgentJson.Measure(result) - maximumCharacters;
            if (result.Content.Length > 0) {
                int remove = Math.Min(result.Content.Length, Math.Max(excess + 8, result.Content.Length / 8));
                result.Content = result.Content.Substring(0, result.Content.Length - remove);
                continue;
            }
            if (diagnostics.Count > 0) {
                diagnostics.RemoveAt(diagnostics.Count - 1);
                continue;
            }
            if (metadata.Count > 0) {
                metadata.RemoveAt(metadata.Count - 1);
                continue;
            }
            if (!string.IsNullOrEmpty(result.Title)) {
                result.Title = Reduce(result.Title);
                continue;
            }
            break;
        }
        int returnedUntil = cursor + result.Content.Length;
        if (returnedUntil < result.ContentLength) {
            result.Truncated = true;
            result.NextCursor = returnedUntil;
        }
        EnsureWithinBudget(result, maximumCharacters);
    }

    private static string? Reduce(string value) =>
        value.Length <= 1
            ? null
            : AgentJson.Limit(value, value.Length / 2);

    private static void EnsureWithinBudget<T>(T result, int maximumCharacters) {
        int actual = AgentJson.Measure(result);
        if (actual <= maximumCharacters) return;
        throw new AgentUsageException(
            "The output budget is too small for stable identifiers in this result. " +
            "Increase max output characters above " + maximumCharacters + ".");
    }
}
