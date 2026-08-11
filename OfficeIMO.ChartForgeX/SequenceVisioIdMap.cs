using System;
using System.Collections.Generic;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

internal sealed class SequenceVisioIdMap {
    private readonly Dictionary<string, string> _participants = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _messages = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _annotations = new(StringComparer.Ordinal);
    private readonly HashSet<string> _reserved = new(StringComparer.Ordinal);
    private int _activationNumber = 1;

    public SequenceVisioIdMap(
        IEnumerable<VisualArtifactInterchangeNode> participants,
        IEnumerable<VisualArtifactInterchangeEdge> messages,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        bool includeTitle) {
        foreach (VisualArtifactInterchangeNode participant in participants) {
            _participants.Add(participant.Id, Allocate(participant.Id, "participant-", "-lifeline", "-lifeline-end"));
        }
        foreach (VisualArtifactInterchangeEdge message in messages) {
            _messages.Add(message.Id, Allocate(message.Id, "message-", "-from", "-to"));
        }
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations) {
            string[] helpers = annotation.Kind.StartsWith("SequenceBlock:", StringComparison.Ordinal)
                ? new[] { "-label" }
                : Array.Empty<string>();
            _annotations.Add(annotation.Id, Allocate(annotation.Id, "annotation-", helpers));
        }
        if (includeTitle) TitleId = Allocate("cfx-title", "title-");
    }

    public string? TitleId { get; }

    public string Participant(string sourceId) => _participants[sourceId];

    public string Message(string sourceId) => _messages[sourceId];

    public string Annotation(string sourceId) => _annotations[sourceId];

    public string Activation() {
        while (true) {
            string candidate = "activation-" + _activationNumber.ToString(global::System.Globalization.CultureInfo.InvariantCulture);
            _activationNumber++;
            if (_reserved.Add(candidate)) return candidate;
        }
    }

    private string Allocate(string sourceId, string prefix, params string[] helperSuffixes) {
        string candidate = sourceId;
        var suffix = 1;
        while (!CanReserve(candidate, helperSuffixes)) {
            suffix++;
            candidate = prefix + sourceId + (suffix == 2 ? string.Empty : "-" + suffix.ToString(global::System.Globalization.CultureInfo.InvariantCulture));
        }
        Reserve(candidate, helperSuffixes);
        return candidate;
    }

    private bool CanReserve(string candidate, IEnumerable<string> helperSuffixes) {
        if (_reserved.Contains(candidate)) return false;
        foreach (string helperSuffix in helperSuffixes) {
            if (_reserved.Contains(candidate + helperSuffix)) return false;
        }
        return true;
    }

    private void Reserve(string candidate, IEnumerable<string> helperSuffixes) {
        _reserved.Add(candidate);
        foreach (string helperSuffix in helperSuffixes) _reserved.Add(candidate + helperSuffix);
    }
}
