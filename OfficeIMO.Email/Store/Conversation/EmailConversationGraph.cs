namespace OfficeIMO.Email.Store;

/// <summary>Direction and meaning of a conversation graph edge.</summary>
public enum EmailConversationEdgeKind {
    /// <summary>Source is a parent message and target is a reply/child.</summary>
    ParentChild,
    /// <summary>Source and target are related but no parent ordering was inferred.</summary>
    Related
}

/// <summary>Evidence used to connect two store items.</summary>
public enum EmailConversationLinkReason {
    /// <summary>Internet In-Reply-To matched a unique Message-ID.</summary>
    InReplyTo,
    /// <summary>The nearest available Internet References identifier matched a unique Message-ID.</summary>
    References,
    /// <summary>A valid Outlook conversation index had an exact parent prefix.</summary>
    ConversationIndexParent,
    /// <summary>Binary Outlook conversation identifiers matched.</summary>
    ConversationId,
    /// <summary>Valid Outlook conversation indexes shared their 22-byte root.</summary>
    ConversationIndexRoot,
    /// <summary>Meeting Clean/Global Object IDs matched.</summary>
    MeetingGlobalObjectId,
    /// <summary>Task Global IDs matched.</summary>
    TaskGlobalId,
    /// <summary>Otherwise unconnected items shared an exact Outlook conversation topic.</summary>
    ConversationTopic,
    /// <summary>Otherwise unconnected items shared a normalized subject.</summary>
    NormalizedSubject
}

/// <summary>Reason why a declared Internet parent could not be selected.</summary>
public enum EmailConversationOrphanReason {
    /// <summary>No scanned item published the requested Message-ID.</summary>
    MissingParent,
    /// <summary>Multiple scanned items published the requested Message-ID.</summary>
    AmbiguousParent
}

/// <summary>One immutable item node in an offline conversation graph.</summary>
public sealed class EmailConversationNode {
    internal EmailConversationNode(EmailStoreItemReference reference, EmailStoreItemSummary summary,
        string? messageId, IReadOnlyList<string> references, string? inReplyToId,
        string? normalizedSubject, string? conversationTopic, string? conversationIdKey,
        string? conversationIndexKey, string? conversationIndexRootKey,
        string? meetingGlobalObjectIdKey, Guid? taskGlobalId) {
        Reference = reference;
        Summary = summary;
        MessageId = messageId;
        References = references;
        InReplyToId = inReplyToId;
        NormalizedSubject = normalizedSubject;
        ConversationTopic = conversationTopic;
        ConversationIdKey = conversationIdKey;
        ConversationIndexKey = conversationIndexKey;
        ConversationIndexRootKey = conversationIndexRootKey;
        MeetingGlobalObjectIdKey = meetingGlobalObjectIdKey;
        TaskGlobalId = taskGlobalId;
    }

    /// <summary>Stable store reference.</summary>
    public EmailStoreItemReference Reference { get; }
    /// <summary>Lightweight item summary enriched by the selective metadata read.</summary>
    public EmailStoreItemSummary Summary { get; }
    /// <summary>Normalized Internet Message-ID without angle brackets.</summary>
    public string? MessageId { get; }
    /// <summary>Ordered, bounded Internet References identifiers.</summary>
    public IReadOnlyList<string> References { get; }
    /// <summary>Normalized Internet In-Reply-To identifier.</summary>
    public string? InReplyToId { get; }
    /// <summary>Source or locally normalized subject.</summary>
    public string? NormalizedSubject { get; }
    /// <summary>Outlook conversation topic.</summary>
    public string? ConversationTopic { get; }
    /// <summary>Hexadecimal binary Outlook conversation identifier.</summary>
    public string? ConversationIdKey { get; }
    /// <summary>Hexadecimal valid Outlook conversation index.</summary>
    public string? ConversationIndexKey { get; }
    /// <summary>Hexadecimal 22-byte root of a valid Outlook conversation index.</summary>
    public string? ConversationIndexRootKey { get; }
    /// <summary>Hexadecimal Clean/Global Object ID used for meeting lifecycle correlation.</summary>
    public string? MeetingGlobalObjectIdKey { get; }
    /// <summary>Task lifecycle correlation identifier.</summary>
    public Guid? TaskGlobalId { get; }
}

/// <summary>One merged-evidence graph edge.</summary>
public sealed class EmailConversationEdge {
    internal EmailConversationEdge(EmailConversationNode source, EmailConversationNode target,
        EmailConversationEdgeKind kind, IReadOnlyList<EmailConversationLinkReason> reasons) {
        Source = source;
        Target = target;
        Kind = kind;
        Reasons = reasons;
    }

    /// <summary>Parent for a parent-child edge, or deterministic group anchor for a related edge.</summary>
    public EmailConversationNode Source { get; }
    /// <summary>Child for a parent-child edge, or another related item.</summary>
    public EmailConversationNode Target { get; }
    /// <summary>Edge direction semantics.</summary>
    public EmailConversationEdgeKind Kind { get; }
    /// <summary>All independent evidence that produced this edge.</summary>
    public IReadOnlyList<EmailConversationLinkReason> Reasons { get; }
    /// <summary>True when at least one structured or explicit identity links the items.</summary>
    public bool IsAuthoritative => Reasons.Any(reason =>
        reason != EmailConversationLinkReason.ConversationTopic &&
        reason != EmailConversationLinkReason.NormalizedSubject);
    /// <summary>True when every reason is a topic/subject heuristic.</summary>
    public bool IsHeuristic => !IsAuthoritative;
}

/// <summary>One connected conversation component.</summary>
public sealed class EmailConversation {
    internal EmailConversation(string id, IReadOnlyList<EmailConversationNode> nodes,
        IReadOnlyList<EmailConversationNode> roots, IReadOnlyList<EmailConversationEdge> edges) {
        Id = id;
        Nodes = nodes;
        Roots = roots;
        Edges = edges;
    }

    /// <summary>Stable identifier within this graph snapshot.</summary>
    public string Id { get; }
    /// <summary>All nodes in deterministic chronological/source order.</summary>
    public IReadOnlyList<EmailConversationNode> Nodes { get; }
    /// <summary>Nodes without an in-graph parent; a deterministic fallback root is used for malformed cycles.</summary>
    public IReadOnlyList<EmailConversationNode> Roots { get; }
    /// <summary>Edges whose endpoints both belong to this component.</summary>
    public IReadOnlyList<EmailConversationEdge> Edges { get; }
    /// <summary>Whether the component is connected exclusively by heuristic edges.</summary>
    public bool IsHeuristicOnly => Edges.Count > 0 && Edges.All(edge => edge.IsHeuristic);
}

/// <summary>Duplicate Message-ID evidence retained instead of silently choosing a parent.</summary>
public sealed class EmailConversationDuplicateMessageId {
    internal EmailConversationDuplicateMessageId(string messageId,
        IReadOnlyList<EmailConversationNode> nodes) {
        MessageId = messageId;
        Nodes = nodes;
    }
    /// <summary>Duplicated normalized Message-ID.</summary>
    public string MessageId { get; }
    /// <summary>All nodes that published the identifier.</summary>
    public IReadOnlyList<EmailConversationNode> Nodes { get; }
}

/// <summary>A reply whose declared Internet parent could not be selected safely.</summary>
public sealed class EmailConversationOrphanReply {
    internal EmailConversationOrphanReply(EmailConversationNode child, string parentMessageId,
        EmailConversationLinkReason linkReason, EmailConversationOrphanReason reason) {
        Child = child;
        ParentMessageId = parentMessageId;
        LinkReason = linkReason;
        Reason = reason;
    }
    /// <summary>Reply node.</summary>
    public EmailConversationNode Child { get; }
    /// <summary>Declared normalized parent Message-ID.</summary>
    public string ParentMessageId { get; }
    /// <summary>Header field that supplied the parent identity.</summary>
    public EmailConversationLinkReason LinkReason { get; }
    /// <summary>Why a parent edge was not created.</summary>
    public EmailConversationOrphanReason Reason { get; }
}

/// <summary>Bounded, duplicate-aware, cross-folder offline conversation graph.</summary>
public sealed class EmailConversationGraph {
    private readonly Dictionary<EmailStoreItemId, EmailConversationNode> _nodesById;
    private readonly Dictionary<EmailStoreItemId, EmailConversation> _conversationsByNode;

    internal EmailConversationGraph(IReadOnlyList<EmailConversationNode> nodes,
        IReadOnlyList<EmailConversationEdge> edges, IReadOnlyList<EmailConversation> conversations,
        IReadOnlyList<EmailConversationDuplicateMessageId> duplicateMessageIds,
        IReadOnlyList<EmailConversationOrphanReply> orphanReplies,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, int itemsScanned, int itemMetadataReads,
        bool itemLimitReached, bool edgeLimitReached, bool isComplete) {
        Nodes = nodes;
        Edges = edges;
        Conversations = conversations;
        DuplicateMessageIds = duplicateMessageIds;
        OrphanReplies = orphanReplies;
        Diagnostics = diagnostics;
        ItemsScanned = itemsScanned;
        ItemMetadataReads = itemMetadataReads;
        ItemLimitReached = itemLimitReached;
        EdgeLimitReached = edgeLimitReached;
        IsComplete = isComplete;
        _nodesById = nodes.ToDictionary(node => node.Reference.Key);
        _conversationsByNode = new Dictionary<EmailStoreItemId, EmailConversation>();
        foreach (EmailConversation conversation in conversations) {
            foreach (EmailConversationNode node in conversation.Nodes) {
                _conversationsByNode[node.Reference.Key] = conversation;
            }
        }
    }

    /// <summary>All scanned graph nodes.</summary>
    public IReadOnlyList<EmailConversationNode> Nodes { get; }
    /// <summary>All bounded graph edges.</summary>
    public IReadOnlyList<EmailConversationEdge> Edges { get; }
    /// <summary>Connected conversation components, including isolated items.</summary>
    public IReadOnlyList<EmailConversation> Conversations { get; }
    /// <summary>Duplicate Message-ID groups.</summary>
    public IReadOnlyList<EmailConversationDuplicateMessageId> DuplicateMessageIds { get; }
    /// <summary>Replies with a missing or ambiguous declared Internet parent.</summary>
    public IReadOnlyList<EmailConversationOrphanReply> OrphanReplies { get; }
    /// <summary>Build, limit, malformed-identity, and recoverable-read diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Items retained in the graph.</summary>
    public int ItemsScanned { get; }
    /// <summary>Successful selective metadata reads.</summary>
    public int ItemMetadataReads { get; }
    /// <summary>Whether at least one additional item existed beyond the configured item bound.</summary>
    public bool ItemLimitReached { get; }
    /// <summary>Whether additional edges were omitted by the configured edge bound.</summary>
    public bool EdgeLimitReached { get; }
    /// <summary>Whether item reads and both configured bounds produced a complete graph for the selected scope.</summary>
    public bool IsComplete { get; }

    /// <summary>Gets one node by typed store identifier.</summary>
    public EmailConversationNode GetNode(EmailStoreItemId itemId) =>
        _nodesById.TryGetValue(itemId, out EmailConversationNode? node)
            ? node
            : throw new KeyNotFoundException(string.Concat("Conversation node not found: ", itemId.ToString()));

    /// <summary>Gets the connected component containing an item.</summary>
    public EmailConversation GetConversation(EmailStoreItemId itemId) =>
        _conversationsByNode.TryGetValue(itemId, out EmailConversation? conversation)
            ? conversation
            : throw new KeyNotFoundException(string.Concat("Conversation node not found: ", itemId.ToString()));

    /// <summary>Returns direct children selected by explicit Internet or Outlook conversation-index evidence.</summary>
    public IReadOnlyList<EmailConversationNode> GetChildren(EmailStoreItemId itemId) => Edges
        .Where(edge => edge.Kind == EmailConversationEdgeKind.ParentChild &&
            edge.Source.Reference.Key == itemId)
        .Select(edge => edge.Target)
        .OrderBy(node => node.Summary.ReceivedAt ?? node.Summary.SentAt)
        .ThenBy(node => node.Reference.Id, StringComparer.Ordinal)
        .ToArray();

    /// <summary>Returns direct parents selected by explicit Internet or Outlook conversation-index evidence.</summary>
    public IReadOnlyList<EmailConversationNode> GetParents(EmailStoreItemId itemId) => Edges
        .Where(edge => edge.Kind == EmailConversationEdgeKind.ParentChild &&
            edge.Target.Reference.Key == itemId)
        .Select(edge => edge.Source)
        .OrderBy(node => node.Summary.ReceivedAt ?? node.Summary.SentAt)
        .ThenBy(node => node.Reference.Id, StringComparer.Ordinal)
        .ToArray();
}
