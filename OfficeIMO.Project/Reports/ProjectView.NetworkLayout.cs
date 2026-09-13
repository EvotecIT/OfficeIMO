namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private int[][] NetworkStages(CancellationToken token) {
        var locations = Rows.Select((row, index) => (row.Uid, index)).ToDictionary(item => item.Uid, item => item.index);
        var successors = Enumerable.Range(0, Rows.Count).Select(_ => new List<int>()).ToArray();
        var predecessors = Enumerable.Range(0, Rows.Count).Select(_ => new List<int>()).ToArray();
        var indegree = new int[Rows.Count]; var rank = new int[Rows.Count];
        foreach (var link in Links) {
            token.ThrowIfCancellationRequested();
            int from = locations[link.PredecessorUid], to = locations[link.SuccessorUid];
            successors[from].Add(to); predecessors[to].Add(from); indegree[to]++;
        }
        var ready = new SortedSet<int>(Enumerable.Range(0, Rows.Count).Where(index => indegree[index] == 0));
        int visited = 0;
        while (ready.Count > 0) {
            token.ThrowIfCancellationRequested();
            int from = ready.Min; ready.Remove(from); visited++;
            foreach (int to in successors[from]) {
                rank[to] = Math.Max(rank[to], rank[from] + 1);
                if (--indegree[to] == 0) ready.Add(to);
            }
        }
        if (visited != Rows.Count) throw new InvalidOperationException("A dependency network cannot lay out a cycle.");
        var stages = Enumerable.Range(0, Rows.Count).GroupBy(index => rank[index]).OrderBy(group => group.Key).Select(group => group.ToArray()).ToArray();
        var order = new double[Rows.Count];
        for (int stage = 0; stage < stages.Length; stage++) {
            token.ThrowIfCancellationRequested();
            stages[stage] = stages[stage].OrderBy(index => predecessors[index].Count == 0 ? index : predecessors[index].Average(parent => order[parent])).ThenBy(index => index).ToArray();
            for (int lane = 0; lane < stages[stage].Length; lane++) order[stages[stage][lane]] = lane;
        }
        return stages;
    }
}
