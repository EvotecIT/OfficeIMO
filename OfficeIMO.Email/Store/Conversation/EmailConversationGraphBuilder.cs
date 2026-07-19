using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class EmailConversationGraphBuilder {
    private static readonly ISet<string> SubjectPrefixes = new HashSet<string>(
        new[] { "RE", "FW", "FWD", "AW", "SV", "VS", "WG", "TR" },
        StringComparer.OrdinalIgnoreCase);
    private readonly EmailStoreSession _session;
    private readonly EmailConversationGraphOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly List<EmailConversationOrphanReply> _orphans = new List<EmailConversationOrphanReply>();
    private readonly Dictionary<string, MutableEdge> _edges =
        new Dictionary<string, MutableEdge>(StringComparer.Ordinal);
    private bool _edgeLimitReached;
    private bool _readIncomplete;
    private int _metadataReads;

    internal EmailConversationGraphBuilder(EmailStoreSession session,
        EmailConversationGraphOptions options, CancellationToken cancellationToken) {
        _session = session;
        _options = options;
        _cancellationToken = cancellationToken;
    }

    internal EmailConversationGraph Build() {
        List<EmailStoreItemReference> references = EnumerateReferences(out bool itemLimitReached);
        var states = new List<NodeState>(references.Count);
        foreach (EmailStoreItemReference reference in references) {
            _cancellationToken.ThrowIfCancellationRequested();
            states.Add(ReadNode(reference));
        }

        Dictionary<string, List<NodeState>> byMessageId = Group(states, state => state.Node.MessageId);
        IReadOnlyList<EmailConversationDuplicateMessageId> duplicates = byMessageId
            .Where(pair => pair.Value.Count > 1)
            .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => new EmailConversationDuplicateMessageId(pair.Key,
                OrderNodes(pair.Value.Select(state => state.Node)).ToArray()))
            .ToArray();

        AddInternetParentEdges(states, byMessageId);
        AddConversationIndexParentEdges(states);
        AddStrongRelatedEdges(states);
        if (_options.IncludeSubjectHeuristics) AddHeuristicRelatedEdges(states);

        EmailConversationEdge[] edges = _edges.Values
            .OrderBy(edge => edge.Kind)
            .ThenBy(edge => edge.Source.Node.Reference.Id, StringComparer.Ordinal)
            .ThenBy(edge => edge.Target.Node.Reference.Id, StringComparer.Ordinal)
            .Select(edge => edge.ToImmutable())
            .ToArray();
        EmailConversation[] conversations = BuildConversations(states, edges);

        if (itemLimitReached) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_CONVERSATION_ITEM_LIMIT",
                string.Concat("At least one item was omitted after MaxItems=",
                    _options.MaxItems.ToString(CultureInfo.InvariantCulture), "."),
                EmailStoreDiagnosticSeverity.Warning));
        }
        if (_edgeLimitReached) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_CONVERSATION_EDGE_LIMIT",
                string.Concat("Additional links were omitted after MaxEdges=",
                    _options.MaxEdges.ToString(CultureInfo.InvariantCulture), "."),
                EmailStoreDiagnosticSeverity.Warning));
        }
        bool isComplete = !itemLimitReached && !_edgeLimitReached && !_readIncomplete;
        return new EmailConversationGraph(
            states.Select(state => state.Node).ToArray(), edges, conversations,
            duplicates, _orphans.ToArray(), _diagnostics.ToArray(), states.Count,
            _metadataReads, itemLimitReached, _edgeLimitReached, isComplete);
    }

    private List<EmailStoreItemReference> EnumerateReferences(out bool itemLimitReached) {
        int probeLimit = _options.MaxItems == int.MaxValue ? int.MaxValue : _options.MaxItems + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            _options.FolderId?.Value,
            _options.IncludeDescendants,
            _options.IncludeAssociatedItems,
            _options.IncludeOrphanedItems,
            probeLimit);
        List<EmailStoreItemReference> references = _session
            .EnumerateItems(enumeration, _cancellationToken)
            .Take(probeLimit)
            .ToList();
        itemLimitReached = references.Count > _options.MaxItems;
        if (itemLimitReached) references.RemoveRange(_options.MaxItems, references.Count - _options.MaxItems);
        return references;
    }

    private NodeState ReadNode(EmailStoreItemReference reference) {
        EmailStoreItemSummary fallback;
        try {
            fallback = _session.ReadSummary(reference, _cancellationToken);
        } catch (Exception exception) when (CanContinue(exception)) {
            if (!_options.ContinueOnItemError) throw;
            _readIncomplete = true;
            _diagnostics.Add(ReadDiagnostic(reference, exception, "summary"));
            fallback = new EmailStoreItemSummary(new EmailDocument(), null, null);
        }

        EmailStoreItemSummary summary = fallback;
        try {
            var readOptions = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.Metadata,
                _options.MaxDecodedPropertyBytesPerItem);
            EmailStoreItem item = _session.ReadItem(reference, readOptions, _cancellationToken);
            _metadataReads++;
            summary = EmailStoreItemSummary.FromMetadata(item.Document, fallback);
        } catch (Exception exception) when (CanContinue(exception)) {
            if (!_options.ContinueOnItemError) throw;
            _readIncomplete = true;
            _diagnostics.Add(ReadDiagnostic(reference, exception, "metadata"));
        }

        bool referencesTruncated;
        IReadOnlyList<string> internetReferences = ParseMessageIds(
            summary.InternetReferences, _options.MaxReferencesPerItem, out referencesTruncated);
        if (referencesTruncated) {
            _readIncomplete = true;
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_CONVERSATION_REFERENCE_LIMIT",
                string.Concat("References exceeded MaxReferencesPerItem=",
                    _options.MaxReferencesPerItem.ToString(CultureInfo.InvariantCulture), "."),
                EmailStoreDiagnosticSeverity.Warning,
                reference.Id));
        }
        string? messageId = NormalizeMessageId(summary.MessageId);
        string? inReplyTo = ParseMessageIds(summary.InReplyToId, 1, out _).FirstOrDefault();
        string? normalizedSubject = NormalizeSubject(summary.NormalizedSubject ?? summary.Subject);
        string? topic = CollapseWhitespace(summary.ConversationTopic);
        string? conversationId = Hex(summary.ConversationId);
        string? conversationIndex = ValidConversationIndex(summary.ConversationIndex, reference.Id);
        string? conversationIndexRoot = conversationIndex?.Substring(0, 44);
        string? meetingKey = Hex(summary.MeetingCleanGlobalObjectId) ?? Hex(summary.MeetingGlobalObjectId);
        var node = new EmailConversationNode(reference, summary, messageId, internetReferences,
            inReplyTo, normalizedSubject, topic, conversationId, conversationIndex,
            conversationIndexRoot, meetingKey, summary.TaskGlobalId);
        return new NodeState(node);
    }

    private void AddInternetParentEdges(IReadOnlyList<NodeState> states,
        IReadOnlyDictionary<string, List<NodeState>> byMessageId) {
        foreach (NodeState child in states) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (child.Node.InReplyToId != null) {
                bool inReplyResolved = AddDeclaredParent(child, child.Node.InReplyToId,
                    EmailConversationLinkReason.InReplyTo, byMessageId, recordOrphan: false);
                if (inReplyResolved) continue;
            }
            bool referenceResolved = false;
            for (int index = child.Node.References.Count - 1; index >= 0; index--) {
                string reference = child.Node.References[index];
                if (AddDeclaredParent(child, reference, EmailConversationLinkReason.References,
                    byMessageId, recordOrphan: false)) {
                    referenceResolved = true;
                    break;
                }
            }
            if (!referenceResolved) {
                string? parent = child.Node.InReplyToId ?? child.Node.References.LastOrDefault();
                if (parent == null) continue;
                EmailConversationOrphanReason reason = byMessageId.TryGetValue(parent,
                    out List<NodeState>? candidates) && candidates.Count > 1
                    ? EmailConversationOrphanReason.AmbiguousParent
                    : EmailConversationOrphanReason.MissingParent;
                _orphans.Add(new EmailConversationOrphanReply(child.Node, parent,
                    child.Node.InReplyToId != null
                        ? EmailConversationLinkReason.InReplyTo
                        : EmailConversationLinkReason.References,
                    reason));
            }
        }
    }

    private bool AddDeclaredParent(NodeState child, string parentMessageId,
        EmailConversationLinkReason reason,
        IReadOnlyDictionary<string, List<NodeState>> byMessageId,
        bool recordOrphan) {
        if (!byMessageId.TryGetValue(parentMessageId, out List<NodeState>? candidates) ||
            candidates.Count == 0) {
            if (recordOrphan) _orphans.Add(new EmailConversationOrphanReply(child.Node,
                parentMessageId, reason, EmailConversationOrphanReason.MissingParent));
            return false;
        }
        List<NodeState> distinct = candidates.Where(candidate => candidate != child).ToList();
        if (distinct.Count != 1) {
            if (recordOrphan) _orphans.Add(new EmailConversationOrphanReply(child.Node,
                parentMessageId, reason, EmailConversationOrphanReason.AmbiguousParent));
            return false;
        }
        AddEdge(distinct[0], child, EmailConversationEdgeKind.ParentChild, reason);
        distinct[0].HasStrongEvidence = true;
        child.HasStrongEvidence = true;
        return true;
    }

    private void AddConversationIndexParentEdges(IReadOnlyList<NodeState> states) {
        Dictionary<string, List<NodeState>> byIndex = Group(states,
            state => state.Node.ConversationIndexKey);
        foreach (NodeState child in states.Where(state =>
            state.Node.ConversationIndexKey != null && state.Node.ConversationIndexKey.Length > 44)) {
            string parentKey = child.Node.ConversationIndexKey!.Substring(
                0, child.Node.ConversationIndexKey.Length - 10);
            if (!byIndex.TryGetValue(parentKey, out List<NodeState>? parents)) continue;
            List<NodeState> distinct = parents.Where(parent => parent != child).ToList();
            if (distinct.Count == 1) {
                AddEdge(distinct[0], child, EmailConversationEdgeKind.ParentChild,
                    EmailConversationLinkReason.ConversationIndexParent);
                distinct[0].HasStrongEvidence = true;
                child.HasStrongEvidence = true;
            } else if (distinct.Count > 1) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_CONVERSATION_INDEX_PARENT_AMBIGUOUS",
                    "Multiple items published the exact parent conversation index; no parent was guessed.",
                    EmailStoreDiagnosticSeverity.Warning,
                    child.Node.Reference.Id));
            }
        }
    }

    private void AddStrongRelatedEdges(IReadOnlyList<NodeState> states) {
        AddRelatedGroups(Group(states, state => state.Node.ConversationIdKey),
            EmailConversationLinkReason.ConversationId, authoritative: true);
        AddRelatedGroups(Group(states, state => state.Node.ConversationIndexRootKey),
            EmailConversationLinkReason.ConversationIndexRoot, authoritative: true);
        if (_options.IncludeMeetingAndTaskLinks) {
            AddRelatedGroups(Group(states, state => state.Node.MeetingGlobalObjectIdKey),
                EmailConversationLinkReason.MeetingGlobalObjectId, authoritative: true);
            AddRelatedGroups(Group(states, state => state.Node.TaskGlobalId?.ToString("N")),
                EmailConversationLinkReason.TaskGlobalId, authoritative: true);
        }
    }

    private void AddHeuristicRelatedEdges(IReadOnlyList<NodeState> states) {
        AddRelatedGroups(Group(states.Where(state => !state.HasStrongEvidence),
                state => state.Node.ConversationTopic),
            EmailConversationLinkReason.ConversationTopic, authoritative: false);
        AddRelatedGroups(Group(states.Where(state => !state.HasStrongEvidence),
                state => state.Node.NormalizedSubject),
            EmailConversationLinkReason.NormalizedSubject, authoritative: false);
    }

    private void AddRelatedGroups(IReadOnlyDictionary<string, List<NodeState>> groups,
        EmailConversationLinkReason reason, bool authoritative) {
        foreach (List<NodeState> group in groups.Values.Where(items => items.Count > 1)) {
            NodeState[] ordered = group.OrderBy(StateTimestamp)
                .ThenBy(state => state.Node.Reference.Id, StringComparer.Ordinal)
                .ToArray();
            NodeState anchor = ordered[0];
            if (authoritative) {
                foreach (NodeState state in ordered) state.HasStrongEvidence = true;
            }
            for (int index = 1; index < ordered.Length; index++) {
                AddEdge(anchor, ordered[index], EmailConversationEdgeKind.Related, reason);
            }
        }
    }

    private void AddEdge(NodeState source, NodeState target, EmailConversationEdgeKind kind,
        EmailConversationLinkReason reason) {
        if (source == target) return;
        string key = string.Concat(((int)kind).ToString(CultureInfo.InvariantCulture), "\u001F",
            source.Node.Reference.Id, "\u001F", target.Node.Reference.Id);
        if (_edges.TryGetValue(key, out MutableEdge? existing)) {
            existing.AddReason(reason);
            return;
        }
        if (_edges.Count >= _options.MaxEdges) {
            _edgeLimitReached = true;
            return;
        }
        _edges.Add(key, new MutableEdge(source, target, kind, reason));
    }

    private EmailConversation[] BuildConversations(IReadOnlyList<NodeState> states,
        IReadOnlyList<EmailConversationEdge> edges) {
        var union = new UnionFind(states.Select(state => state.Node.Reference.Id));
        foreach (EmailConversationEdge edge in edges) {
            union.Union(edge.Source.Reference.Id, edge.Target.Reference.Id);
        }
        var edgesByRoot = edges.GroupBy(edge => union.Find(edge.Source.Reference.Id))
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var components = states.GroupBy(state => union.Find(state.Node.Reference.Id))
            .Select(group => {
                EmailConversationNode[] nodes = OrderNodes(group.Select(state => state.Node)).ToArray();
                IReadOnlyList<EmailConversationEdge> componentEdges = edgesByRoot.TryGetValue(group.Key,
                    out EmailConversationEdge[]? found) ? found : Array.Empty<EmailConversationEdge>();
                var childIds = new HashSet<string>(componentEdges
                    .Where(edge => edge.Kind == EmailConversationEdgeKind.ParentChild)
                    .Select(edge => edge.Target.Reference.Id), StringComparer.Ordinal);
                EmailConversationNode[] roots = nodes.Where(node => !childIds.Contains(node.Reference.Id)).ToArray();
                if (roots.Length == 0 && nodes.Length > 0) roots = new[] { nodes[0] };
                return new EmailConversation(string.Concat("conversation:", nodes[0].Reference.Id),
                    nodes, roots, componentEdges);
            })
            .OrderBy(conversation => conversation.Nodes[0].Summary.ReceivedAt ??
                conversation.Nodes[0].Summary.SentAt)
            .ThenBy(conversation => conversation.Id, StringComparer.Ordinal)
            .ToArray();
        return components;
    }

    private string? ValidConversationIndex(byte[]? value, string location) {
        if (value == null || value.Length == 0) return null;
        if (value.Length < 22 || (value.Length - 22) % 5 != 0) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_CONVERSATION_INDEX_INVALID",
                "ConversationIndex must contain a 22-byte root followed by zero or more 5-byte child blocks.",
                EmailStoreDiagnosticSeverity.Warning,
                location));
            return null;
        }
        return Hex(value);
    }

    private static Dictionary<string, List<NodeState>> Group(IEnumerable<NodeState> states,
        Func<NodeState, string?> selector) {
        var groups = new Dictionary<string, List<NodeState>>(StringComparer.OrdinalIgnoreCase);
        foreach (NodeState state in states) {
            string? key = selector(state);
            if (string.IsNullOrWhiteSpace(key)) continue;
            if (!groups.TryGetValue(key!, out List<NodeState>? values)) {
                values = new List<NodeState>();
                groups.Add(key!, values);
            }
            values.Add(state);
        }
        return groups;
    }

    private static IReadOnlyList<string> ParseMessageIds(string? value, int maximum,
        out bool truncated) {
        truncated = false;
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<string>();
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int offset = 0;
        while (offset < value!.Length) {
            int start = value.IndexOf('<', offset);
            if (start < 0) break;
            int end = value.IndexOf('>', start + 1);
            if (end < 0) break;
            string? candidate = NormalizeMessageId(value.Substring(start, end - start + 1));
            if (candidate != null && seen.Add(candidate)) {
                if (result.Count >= maximum) { truncated = true; break; }
                result.Add(candidate);
            }
            offset = end + 1;
        }
        if (result.Count == 0) {
            foreach (string token in value.Split(new[] { ' ', '\t', '\r', '\n', ',', ';' },
                StringSplitOptions.RemoveEmptyEntries)) {
                string? candidate = NormalizeMessageId(token);
                if (candidate == null || !seen.Add(candidate)) continue;
                if (result.Count >= maximum) { truncated = true; break; }
                result.Add(candidate);
            }
        }
        return result;
    }

    private static string? NormalizeMessageId(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string normalized = value!.Trim().Trim('<', '>').Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static string? NormalizeSubject(string? value) {
        string? subject = CollapseWhitespace(value);
        if (subject == null) return null;
        while (true) {
            int separator = subject.IndexOf(':');
            if (separator <= 0 || separator > 5) break;
            string prefix = subject.Substring(0, separator).Trim();
            if (!SubjectPrefixes.Contains(prefix)) break;
            subject = subject.Substring(separator + 1).TrimStart();
        }
        return subject.Length == 0 ? null : CollapseWhitespace(subject);
    }

    private static string? CollapseWhitespace(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var builder = new StringBuilder(value!.Length);
        bool pendingSpace = false;
        foreach (char character in value.Trim()) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = builder.Length > 0;
            } else {
                if (pendingSpace) builder.Append(' ');
                builder.Append(character);
                pendingSpace = false;
            }
        }
        return builder.Length == 0 ? null : builder.ToString();
    }

    private static string? Hex(byte[]? value) => value == null || value.Length == 0
        ? null
        : BitConverter.ToString(value).Replace("-", string.Empty);

    private bool CanContinue(Exception exception) =>
        _options.ContinueOnItemError && (exception is InvalidDataException ||
            exception is NotSupportedException || exception is EmailStoreLimitExceededException);

    private static EmailStoreDiagnostic ReadDiagnostic(EmailStoreItemReference reference,
        Exception exception, string part) => new EmailStoreDiagnostic(
        "EMAIL_STORE_CONVERSATION_ITEM_READ",
        string.Concat("Conversation graph could not read item ", part, ": ", exception.Message),
        EmailStoreDiagnosticSeverity.Error,
        reference.Id);

    private static DateTimeOffset StateTimestamp(NodeState state) =>
        state.Node.Summary.ReceivedAt ?? state.Node.Summary.SentAt ?? DateTimeOffset.MaxValue;

    private static IOrderedEnumerable<EmailConversationNode> OrderNodes(
        IEnumerable<EmailConversationNode> nodes) => nodes
        .OrderBy(node => node.Summary.ReceivedAt ?? node.Summary.SentAt ?? DateTimeOffset.MaxValue)
        .ThenBy(node => node.Reference.Id, StringComparer.Ordinal);

    private sealed class NodeState {
        internal NodeState(EmailConversationNode node) { Node = node; }
        internal EmailConversationNode Node { get; }
        internal bool HasStrongEvidence { get; set; }
    }

    private sealed class MutableEdge {
        private readonly HashSet<EmailConversationLinkReason> _reasons =
            new HashSet<EmailConversationLinkReason>();
        internal MutableEdge(NodeState source, NodeState target, EmailConversationEdgeKind kind,
            EmailConversationLinkReason reason) {
            Source = source;
            Target = target;
            Kind = kind;
            _reasons.Add(reason);
        }
        internal NodeState Source { get; }
        internal NodeState Target { get; }
        internal EmailConversationEdgeKind Kind { get; }
        internal void AddReason(EmailConversationLinkReason reason) => _reasons.Add(reason);
        internal EmailConversationEdge ToImmutable() => new EmailConversationEdge(
            Source.Node, Target.Node, Kind, _reasons.OrderBy(reason => reason).ToArray());
    }

    private sealed class UnionFind {
        private readonly Dictionary<string, string> _parents;
        internal UnionFind(IEnumerable<string> values) {
            _parents = values.ToDictionary(value => value, value => value, StringComparer.Ordinal);
        }
        internal string Find(string value) {
            string parent = _parents[value];
            if (!string.Equals(parent, value, StringComparison.Ordinal)) {
                _parents[value] = Find(parent);
            }
            return _parents[value];
        }
        internal void Union(string left, string right) {
            string leftRoot = Find(left);
            string rightRoot = Find(right);
            if (string.Equals(leftRoot, rightRoot, StringComparison.Ordinal)) return;
            if (StringComparer.Ordinal.Compare(leftRoot, rightRoot) <= 0) _parents[rightRoot] = leftRoot;
            else _parents[leftRoot] = rightRoot;
        }
    }
}
