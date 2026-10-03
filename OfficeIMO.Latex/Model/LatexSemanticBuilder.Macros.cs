using System.Threading;

namespace OfficeIMO.Latex;

internal static partial class LatexSemanticBuilder {
    private static IReadOnlyList<LatexMacroDefinition> BuildMacroDefinitions(IReadOnlyList<LatexCommand> commands, CancellationToken cancellationToken) {
        var candidates = new List<MacroCandidate>();
        foreach (LatexCommand command in commands.Where(static command =>
                     string.Equals(command.Name, "newcommand", StringComparison.Ordinal) ||
                     string.Equals(command.Name, "renewcommand", StringComparison.Ordinal) ||
                     string.Equals(command.Name, "providecommand", StringComparison.Ordinal))) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexArgument? nameArgument = command.GetRequiredArgument(0);
            LatexArgument? bodyArgument = command.GetRequiredArgument(1);
            if (nameArgument == null || bodyArgument == null) continue;
            string name = nameArgument.Content.Trim();
            if (name.StartsWith("\\", StringComparison.Ordinal)) name = name.Substring(1);
            if (!IsSimpleControlWord(name)) continue;
            int parameterCount = 0;
            string? defaultValue = null;
            LatexArgument[] optional = command.Arguments.Where(static argument => argument.IsOptional).ToArray();
            bool isWellFormed = optional.Length == 0 || int.TryParse(optional[0].Content.Trim(), out parameterCount);
            if (optional.Length > 1) defaultValue = optional[1].Content;
            if (defaultValue != null && parameterCount < 1) isWellFormed = false;
            candidates.Add(new MacroCandidate(command, name, parameterCount, defaultValue, bodyArgument.Content, isWellFormed));
        }

        // Inspect each replacement once, then propagate rejected names through reverse edges.
        // A bad redefinition invalidates references to that name, matching the conservative profile.
        var localNames = new HashSet<string>(candidates.Select(static candidate => candidate.Name), StringComparer.Ordinal);
        var dependents = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        var rejectedNames = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Queue<string>();
        var safeCandidates = new bool[candidates.Count];
        for (int index = 0; index < candidates.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            MacroCandidate candidate = candidates[index];
            var dependencies = new HashSet<string>(StringComparer.Ordinal);
            safeCandidates[index] = candidate.IsWellFormed && candidate.ParameterCount >= 0 && candidate.ParameterCount <= 9 &&
                IsSafeMacroBody(candidate.Body, candidate.ParameterCount, localNames, dependencies, cancellationToken) &&
                (candidate.DefaultValue == null || IsSafeMacroBody(candidate.DefaultValue, 0, localNames, dependencies, cancellationToken));
            foreach (string dependency in dependencies) {
                if (!dependents.TryGetValue(dependency, out List<int>? indices)) {
                    indices = new List<int>();
                    dependents.Add(dependency, indices);
                }
                indices.Add(index);
            }
            if (!safeCandidates[index] && rejectedNames.Add(candidate.Name)) pending.Enqueue(candidate.Name);
        }
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string name = pending.Dequeue();
            if (!dependents.TryGetValue(name, out List<int>? indices)) continue;
            foreach (int index in indices) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!safeCandidates[index]) continue;
                safeCandidates[index] = false;
                if (rejectedNames.Add(candidates[index].Name)) pending.Enqueue(candidates[index].Name);
            }
        }
        var definitions = new List<LatexMacroDefinition>(candidates.Count);
        for (int index = 0; index < candidates.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            MacroCandidate candidate = candidates[index];
            definitions.Add(new LatexMacroDefinition(candidate.Command, candidate.Name, candidate.ParameterCount,
                candidate.DefaultValue, candidate.Body, safeCandidates[index]));
        }
        return definitions;
    }

    private static bool IsSimpleControlWord(string value) {
        if (value.Length == 0) return false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (!((current >= 'a' && current <= 'z') || (current >= 'A' && current <= 'Z') || current == '@')) return false;
        }
        return true;
    }

    private static bool IsSafeMacroBody(string body, int parameterCount, IReadOnlyCollection<string> localNames, HashSet<string> dependencies, CancellationToken cancellationToken) {
        for (int index = 0; index < body.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (body[index] != '\\' || index + 1 >= body.Length) continue;
            if (!IsControlWordCharacter(body[index + 1])) {
                index++;
                continue;
            }
            int nameStart = ++index;
            while (index + 1 < body.Length && IsControlWordCharacter(body[index + 1])) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                index++;
            }
            string name = body.Substring(nameStart, index - nameStart + 1);
            if (!IsSafeReplacementCommand(name)) {
                if (!localNames.Contains(name)) return false;
                dependencies.Add(name);
            }
        }
        for (int index = 0; index + 1 < body.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (body[index] != '#') continue;
            if (body[index + 1] == '#') { index++; continue; }
            if (body[index + 1] < '1' || body[index + 1] > '9' || body[index + 1] - '0' > parameterCount) return false;
        }
        return true;
    }

    private static bool IsSafeReplacementCommand(string name) {
        switch (name) {
            case "textbf":
            case "textit":
            case "emph":
            case "texttt":
            case "underline":
            case "textsuperscript":
            case "textsubscript":
            case "mathrm":
            case "mathbf":
            case "mathit":
            case "mathsf":
            case "mathtt":
            case "operatorname":
            case "frac":
            case "sqrt":
            case "ensuremath":
            case "protect":
            case "ref":
            case "pageref":
            case "autoref":
            case "eqref":
            case "cite":
            case "citep":
            case "citet":
            case "url":
            case "href":
            case "hyperref":
            case "footnote":
            case "newline":
            case "linebreak":
                return true;
            default:
                return false;
        }
    }

    private static bool IsControlWordCharacter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || value == '@';

    private sealed class MacroCandidate {
        internal MacroCandidate(
            LatexCommand command,
            string name,
            int parameterCount,
            string? defaultValue,
            string body,
            bool isWellFormed) {
            Command = command;
            Name = name;
            ParameterCount = parameterCount;
            DefaultValue = defaultValue;
            Body = body;
            IsWellFormed = isWellFormed;
        }

        internal LatexCommand Command { get; }
        internal string Name { get; }
        internal int ParameterCount { get; }
        internal string? DefaultValue { get; }
        internal string Body { get; }
        internal bool IsWellFormed { get; }
    }

}
