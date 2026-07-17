namespace OfficeIMO.Excel {
    internal sealed class ExcelFormulaDependencyAnalysisResult {
        internal ExcelFormulaDependencyAnalysisResult(
            IReadOnlyDictionary<string, int> depths,
            IReadOnlyList<IReadOnlyList<string>> circularReferenceGroups,
            ISet<string> circularReferences,
            int maximumDepth) {
            Depths = depths;
            CircularReferenceGroups = circularReferenceGroups;
            CircularReferences = circularReferences;
            MaximumDepth = maximumDepth;
        }

        internal IReadOnlyDictionary<string, int> Depths { get; }
        internal IReadOnlyList<IReadOnlyList<string>> CircularReferenceGroups { get; }
        internal ISet<string> CircularReferences { get; }
        internal int MaximumDepth { get; }
    }

    internal static class ExcelFormulaDependencyAnalysis {
        internal static ExcelFormulaDependencyAnalysisResult Analyze(
            IReadOnlyDictionary<string, IReadOnlyCollection<string>> dependencies) {
            var reverse = dependencies.Keys.ToDictionary(
                reference => reference,
                _ => new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, IReadOnlyCollection<string>> pair in dependencies) {
                foreach (string dependency in pair.Value) {
                    if (reverse.TryGetValue(dependency, out HashSet<string>? dependents)) {
                        dependents.Add(pair.Key);
                    }
                }
            }

            List<string> finishingOrder = CreateFinishingOrder(dependencies);
            List<IReadOnlyList<string>> circularGroups = FindCircularGroups(dependencies, reverse, finishingOrder);
            var circularReferences = new HashSet<string>(
                circularGroups.SelectMany(group => group),
                StringComparer.OrdinalIgnoreCase);
            Dictionary<string, int> depths = CalculateResolvableDepths(dependencies, reverse);
            int maximumDepth = depths.Count == 0 ? 0 : depths.Values.Max();

            return new ExcelFormulaDependencyAnalysisResult(depths, circularGroups, circularReferences, maximumDepth);
        }

        private static List<string> CreateFinishingOrder(
            IReadOnlyDictionary<string, IReadOnlyCollection<string>> dependencies) {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var order = new List<string>(dependencies.Count);
            foreach (string root in dependencies.Keys) {
                if (visited.Contains(root)) {
                    continue;
                }

                var stack = new Stack<(string Reference, bool Expanded)>();
                stack.Push((root, false));
                while (stack.Count > 0) {
                    (string reference, bool expanded) = stack.Pop();
                    if (expanded) {
                        order.Add(reference);
                        continue;
                    }

                    if (!visited.Add(reference)) {
                        continue;
                    }

                    stack.Push((reference, true));
                    if (!dependencies.TryGetValue(reference, out IReadOnlyCollection<string>? targets)) {
                        continue;
                    }

                    foreach (string target in targets.OrderByDescending(value => value, StringComparer.OrdinalIgnoreCase)) {
                        if (!visited.Contains(target)) {
                            stack.Push((target, false));
                        }
                    }
                }
            }

            return order;
        }

        private static List<IReadOnlyList<string>> FindCircularGroups(
            IReadOnlyDictionary<string, IReadOnlyCollection<string>> dependencies,
            IReadOnlyDictionary<string, HashSet<string>> reverse,
            IReadOnlyList<string> finishingOrder) {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var groups = new List<IReadOnlyList<string>>();
            for (int index = finishingOrder.Count - 1; index >= 0; index--) {
                string root = finishingOrder[index];
                if (!visited.Add(root)) {
                    continue;
                }

                var component = new List<string>();
                var stack = new Stack<string>();
                stack.Push(root);
                while (stack.Count > 0) {
                    string reference = stack.Pop();
                    component.Add(reference);
                    foreach (string dependent in reverse[reference]) {
                        if (visited.Add(dependent)) {
                            stack.Push(dependent);
                        }
                    }
                }

                bool isCircular = component.Count > 1
                    || (component.Count == 1 && dependencies[component[0]].Contains(component[0], StringComparer.OrdinalIgnoreCase));
                if (isCircular) {
                    groups.Add(component.OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList());
                }
            }

            return groups
                .OrderBy(group => group[0], StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static Dictionary<string, int> CalculateResolvableDepths(
            IReadOnlyDictionary<string, IReadOnlyCollection<string>> dependencies,
            IReadOnlyDictionary<string, HashSet<string>> reverse) {
            var remainingDependencies = dependencies.ToDictionary(
                pair => pair.Key,
                pair => pair.Value.Count,
                StringComparer.OrdinalIgnoreCase);
            var depths = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var candidates = dependencies.Keys.ToDictionary(
                reference => reference,
                _ => 1,
                StringComparer.OrdinalIgnoreCase);
            var ready = new Queue<string>(remainingDependencies
                .Where(pair => pair.Value == 0)
                .Select(pair => pair.Key)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));

            while (ready.Count > 0) {
                string reference = ready.Dequeue();
                int depth = candidates[reference];
                depths[reference] = depth;
                foreach (string dependent in reverse[reference]) {
                    candidates[dependent] = Math.Max(candidates[dependent], depth + 1);
                    remainingDependencies[dependent]--;
                    if (remainingDependencies[dependent] == 0) {
                        ready.Enqueue(dependent);
                    }
                }
            }

            return depths;
        }
    }
}
