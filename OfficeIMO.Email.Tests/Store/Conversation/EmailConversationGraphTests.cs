using OfficeIMO.Email;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailConversationGraphTests {
    [Fact]
    public void Builds_cross_folder_graph_with_merged_evidence_lifecycle_links_and_explicit_ambiguity() {
        Guid taskId = new Guid("f84bce47-8619-4b99-b60d-bf74e02ad746");
        byte[] conversationId = { 1, 2, 3, 4 };
        byte[] rootIndex = Enumerable.Range(0, 22).Select(value => (byte)value).ToArray();
        byte[] childIndex = rootIndex.Concat(new byte[] { 9, 8, 7, 6, 5 }).ToArray();
        byte[] meetingId = { 8, 7, 6, 5, 4, 3 };
        var backend = new ConversationBackend(new[] {
            Item("root", "inbox", "Project", "root@example.test", 1, document => {
                document.MessageMetadata.ConversationId = conversationId;
                document.MessageMetadata.ConversationIndex = rootIndex;
            }),
            Item("reply", "archive", "RE: Project", "reply@example.test", 2, document => {
                document.MessageMetadata.InReplyToId = "<root@example.test>";
                document.MessageMetadata.InternetReferences = "<root@example.test>";
                document.MessageMetadata.ConversationId = conversationId;
                document.MessageMetadata.ConversationIndex = childIndex;
            }),
            Item("orphan", "archive", "Missing", "orphan@example.test", 3, document =>
                document.MessageMetadata.InReplyToId = "<outside@example.test>"),
            Item("duplicate-a", "inbox", "Duplicate A", "duplicate@example.test", 4),
            Item("duplicate-b", "archive", "Duplicate B", "duplicate@example.test", 5),
            Item("meeting-request", "calendar", "Planning", null, 6, document => {
                document.OutlookItemKind = OutlookItemKind.Appointment;
                document.Appointment = new OutlookAppointment { CleanGlobalObjectId = meetingId };
            }),
            Item("meeting-response", "inbox", "Accepted: Planning", null, 7, document => {
                document.OutlookItemKind = OutlookItemKind.Appointment;
                document.Appointment = new OutlookAppointment { CleanGlobalObjectId = meetingId };
            }),
            Item("task-request", "tasks", "Task", null, 8, document => {
                document.OutlookItemKind = OutlookItemKind.Task;
                document.Task = new OutlookTask { GlobalId = taskId };
            }),
            Item("task-update", "inbox", "Task updated", null, 9, document => {
                document.OutlookItemKind = OutlookItemKind.Task;
                document.Task = new OutlookTask { GlobalId = taskId };
            }),
            Item("heuristic-root", "inbox", "Quarterly notes", null, 10),
            Item("heuristic-reply", "archive", "RE:   Quarterly   notes", null, 11)
        });
        using EmailStoreSession session = CreateSession(backend);

        EmailConversationGraph graph = session.BuildConversationGraph();

        Assert.True(graph.IsComplete);
        Assert.Equal(11, graph.ItemsScanned);
        Assert.Equal(11, graph.ItemMetadataReads);
        EmailConversationEdge parent = Assert.Single(graph.Edges, edge =>
            edge.Kind == EmailConversationEdgeKind.ParentChild &&
            edge.Source.Reference.Id == "root" && edge.Target.Reference.Id == "reply");
        Assert.Equal(new[] {
            EmailConversationLinkReason.InReplyTo,
            EmailConversationLinkReason.ConversationIndexParent
        }, parent.Reasons);
        Assert.True(parent.IsAuthoritative);
        Assert.Equal("reply", Assert.Single(graph.GetChildren(new EmailStoreItemId("root"))).Reference.Id);
        Assert.Equal(2, graph.GetConversation(new EmailStoreItemId("root")).Nodes.Count);

        EmailConversationDuplicateMessageId duplicate = Assert.Single(graph.DuplicateMessageIds);
        Assert.Equal("duplicate@example.test", duplicate.MessageId);
        Assert.Equal(new[] { "duplicate-a", "duplicate-b" },
            duplicate.Nodes.Select(node => node.Reference.Id));
        EmailConversationOrphanReply orphan = Assert.Single(graph.OrphanReplies);
        Assert.Equal("outside@example.test", orphan.ParentMessageId);
        Assert.Equal(EmailConversationOrphanReason.MissingParent, orphan.Reason);

        Assert.Contains(graph.Edges, edge => edge.Reasons.Contains(
            EmailConversationLinkReason.MeetingGlobalObjectId) && edge.IsAuthoritative);
        Assert.Contains(graph.Edges, edge => edge.Reasons.Contains(
            EmailConversationLinkReason.TaskGlobalId) && edge.IsAuthoritative);
        EmailConversationEdge heuristic = Assert.Single(graph.Edges, edge =>
            edge.Reasons.Contains(EmailConversationLinkReason.NormalizedSubject));
        Assert.True(heuristic.IsHeuristic);
        Assert.True(graph.GetConversation(new EmailStoreItemId("heuristic-root")).IsHeuristicOnly);
    }

    [Fact]
    public void In_reply_to_is_the_single_preferred_parent_and_references_are_only_a_fallback() {
        var backend = new ConversationBackend(new[] {
            Item("reply-parent", "inbox", "Preferred", "preferred@example.test", 1),
            Item("reference-parent", "inbox", "Fallback", "fallback@example.test", 2),
            Item("reply", "inbox", "RE: Preferred", "reply@example.test", 3, document => {
                document.MessageMetadata.InReplyToId = "<preferred@example.test>";
                document.MessageMetadata.InternetReferences = "<fallback@example.test>";
            }),
            Item("fallback-reply", "inbox", "RE: Fallback", "fallback-reply@example.test", 4,
                document => {
                    document.MessageMetadata.InReplyToId = "<missing@example.test>";
                    document.MessageMetadata.InternetReferences = "<fallback@example.test>";
                })
        });
        using EmailStoreSession session = CreateSession(backend);

        EmailConversationGraph graph = session.BuildConversationGraph();

        EmailConversationEdge preferred = Assert.Single(graph.Edges, edge =>
            edge.Kind == EmailConversationEdgeKind.ParentChild &&
            edge.Target.Reference.Id == "reply");
        Assert.Equal("reply-parent", preferred.Source.Reference.Id);
        Assert.Equal(new[] { EmailConversationLinkReason.InReplyTo }, preferred.Reasons);
        EmailConversationEdge fallback = Assert.Single(graph.Edges, edge =>
            edge.Kind == EmailConversationEdgeKind.ParentChild &&
            edge.Target.Reference.Id == "fallback-reply");
        Assert.Equal("reference-parent", fallback.Source.Reference.Id);
        Assert.Equal(new[] { EmailConversationLinkReason.References }, fallback.Reasons);
        Assert.DoesNotContain(graph.OrphanReplies, orphan =>
            orphan.Child.Reference.Id == "fallback-reply");
    }

    [Fact]
    public void Bounds_report_incomplete_graphs_without_claiming_more_items_or_edges() {
        var backend = new ConversationBackend(new[] {
            Item("a", "inbox", "Same", null, 1),
            Item("b", "inbox", "Same", null, 2),
            Item("c", "inbox", "Same", null, 3),
            Item("d", "inbox", "Same", null, 4)
        });
        using EmailStoreSession session = CreateSession(backend);

        EmailConversationGraph itemBound = session.BuildConversationGraph(
            new EmailConversationGraphOptions(maxItems: 2));
        EmailConversationGraph edgeBound = session.BuildConversationGraph(
            new EmailConversationGraphOptions(maxEdges: 1));

        Assert.False(itemBound.IsComplete);
        Assert.True(itemBound.ItemLimitReached);
        Assert.Equal(2, itemBound.Nodes.Count);
        Assert.Contains(itemBound.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_CONVERSATION_ITEM_LIMIT");
        Assert.False(edgeBound.IsComplete);
        Assert.True(edgeBound.EdgeLimitReached);
        Assert.Single(edgeBound.Edges);
        Assert.Contains(edgeBound.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_CONVERSATION_EDGE_LIMIT");
    }

    [Fact]
    public void Invalid_conversation_index_is_diagnosed_and_not_used_as_structured_evidence() {
        var backend = new ConversationBackend(new[] {
            Item("broken", "inbox", "Broken", null, 1, document =>
                document.MessageMetadata.ConversationIndex = new byte[] { 1, 2, 3 })
        });
        using EmailStoreSession session = CreateSession(backend);

        EmailConversationGraph graph = session.BuildConversationGraph();

        Assert.True(graph.IsComplete);
        Assert.Null(graph.Nodes[0].ConversationIndexKey);
        Assert.Contains(graph.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_CONVERSATION_INDEX_INVALID");
    }

    private static EmailStoreSession CreateSession(IEmailStoreSessionBackend backend) {
        System.Reflection.ConstructorInfo? constructor = typeof(EmailStoreSession).GetConstructor(
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            binder: null,
            new[] {
                typeof(Stream), typeof(bool), typeof(long), typeof(EmailStoreReaderOptions),
                typeof(IEmailStoreSessionBackend)
            },
            modifiers: null);
        Assert.NotNull(constructor);
        return (EmailStoreSession)constructor!.Invoke(new object[] {
            Stream.Null, true, 0L, EmailStoreReaderOptions.Default, backend
        });
    }

    private static ConversationItem Item(string id, string folderId, string subject,
        string? messageId, int day, Action<EmailDocument>? configure = null) {
        var document = new EmailDocument {
            Subject = subject,
            MessageId = messageId,
            Date = new DateTimeOffset(2026, 1, day, 9, 0, 0, TimeSpan.Zero),
            ReceivedDate = new DateTimeOffset(2026, 1, day, 10, 0, 0, TimeSpan.Zero)
        };
        configure?.Invoke(document);
        return new ConversationItem(id, folderId, document);
    }

    private sealed class ConversationItem {
        internal ConversationItem(string id, string folderId, EmailDocument document) {
            Document = document;
            Summary = new EmailStoreItemSummary(document, false, false);
            Reference = new EmailStoreItemReference(id, folderId, false, false, Summary);
        }
        internal EmailDocument Document { get; }
        internal EmailStoreItemSummary Summary { get; }
        internal EmailStoreItemReference Reference { get; }
    }

    private sealed class ConversationBackend : IEmailStoreSessionBackend {
        private readonly ConversationItem[] _items;
        private readonly EmailStoreFolderInfo[] _folders;
        internal ConversationBackend(IEnumerable<ConversationItem> items) {
            _items = items.ToArray();
            _folders = _items.Select(item => item.Reference.FolderId)
                .Distinct(StringComparer.Ordinal)
                .Select(id => new EmailStoreFolderInfo(id, null, id))
                .ToArray();
        }
        public EmailStoreFormat Format => EmailStoreFormat.Pst;
        public string? DisplayName => "Conversation graph test";
        public long SourceLength => 0;
        public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
        public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => Array.Empty<EmailStoreDiagnostic>();

        public IEnumerable<EmailStoreItemReference> EnumerateItems(
            EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
            IEnumerable<ConversationItem> selected = _items;
            if (options.FolderId != null) selected = selected.Where(item =>
                string.Equals(item.Reference.FolderId, options.FolderId, StringComparison.Ordinal));
            foreach (ConversationItem item in selected.Take(options.MaxItems)) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return item.Reference;
            }
        }

        public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Find(reference).Summary;
        }

        public EmailStoreItem ReadItem(EmailStoreItemReference reference,
            EmailStoreItemReadOptions options, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            ConversationItem item = Find(reference);
            return new EmailStoreItem(reference.Id, reference.FolderId, item.Document,
                loadedParts: options.Parts, format: EmailStoreFormat.Pst, summary: item.Summary);
        }

        public void Dispose() { }

        private ConversationItem Find(EmailStoreItemReference reference) => _items.Single(item =>
            string.Equals(item.Reference.Id, reference.Id, StringComparison.Ordinal));
    }
}
