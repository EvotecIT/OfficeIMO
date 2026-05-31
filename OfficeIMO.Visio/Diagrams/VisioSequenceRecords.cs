using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Simple data record used to import sequence diagram participants into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceParticipantRecord {
        /// <summary>Initializes a sequence participant import record.</summary>
        public VisioSequenceParticipantRecord(string id, string text, VisioSequenceParticipantKind kind = VisioSequenceParticipantKind.Participant) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Participant id cannot be null or whitespace.", nameof(id)) : id;
            Text = text ?? string.Empty;
            Kind = kind;
        }

        /// <summary>Stable participant id.</summary>
        public string Id { get; }

        /// <summary>Visible participant label.</summary>
        public string Text { get; }

        /// <summary>Semantic participant kind.</summary>
        public VisioSequenceParticipantKind Kind { get; }
    }

    /// <summary>
    /// Simple data record used to import sequence diagram messages into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceMessageRecord {
        /// <summary>Initializes a sequence message import record.</summary>
        public VisioSequenceMessageRecord(string id, string fromId, string toId, string label, VisioSequenceMessageKind kind = VisioSequenceMessageKind.Call, bool selfMessage = false) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Message id cannot be null or whitespace.", nameof(id)) : id;
            FromId = string.IsNullOrWhiteSpace(fromId) ? throw new ArgumentException("Source participant id cannot be null or whitespace.", nameof(fromId)) : fromId;
            ToId = string.IsNullOrWhiteSpace(toId) ? throw new ArgumentException("Target participant id cannot be null or whitespace.", nameof(toId)) : toId;
            Label = label ?? string.Empty;
            Kind = kind;
            SelfMessage = selfMessage;
        }

        /// <summary>Stable message id.</summary>
        public string Id { get; }

        /// <summary>Source participant id.</summary>
        public string FromId { get; }

        /// <summary>Target participant id.</summary>
        public string ToId { get; }

        /// <summary>Visible message label.</summary>
        public string Label { get; }

        /// <summary>Semantic message kind.</summary>
        public VisioSequenceMessageKind Kind { get; }

        /// <summary>Whether the message should render as a self-message loop on <see cref="FromId"/>.</summary>
        public bool SelfMessage { get; }
    }

    /// <summary>
    /// Simple data record used to import sequence activation bars into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceActivationRecord {
        /// <summary>Initializes a sequence activation import record.</summary>
        public VisioSequenceActivationRecord(string id, string participantId, int startRowIndex, int endRowIndex) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Activation id cannot be null or whitespace.", nameof(id)) : id;
            ParticipantId = string.IsNullOrWhiteSpace(participantId) ? throw new ArgumentException("Participant id cannot be null or whitespace.", nameof(participantId)) : participantId;
            StartRowIndex = startRowIndex;
            EndRowIndex = endRowIndex;
        }

        /// <summary>Stable activation id.</summary>
        public string Id { get; }

        /// <summary>Participant id owning the activation.</summary>
        public string ParticipantId { get; }

        /// <summary>Zero-based starting message row.</summary>
        public int StartRowIndex { get; }

        /// <summary>Zero-based ending message row.</summary>
        public int EndRowIndex { get; }
    }

    /// <summary>
    /// Simple data record used to import sequence combined fragments into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceFragmentRecord {
        /// <summary>Initializes a sequence fragment import record.</summary>
        public VisioSequenceFragmentRecord(string id, string text, int startRowIndex, int endRowIndex, IEnumerable<string>? participantIds = null) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Fragment id cannot be null or whitespace.", nameof(id)) : id;
            Text = text ?? string.Empty;
            StartRowIndex = startRowIndex;
            EndRowIndex = endRowIndex;
            ParticipantIds = (participantIds ?? Array.Empty<string>()).ToArray();
        }

        /// <summary>Stable fragment id.</summary>
        public string Id { get; }

        /// <summary>Visible fragment label.</summary>
        public string Text { get; }

        /// <summary>Zero-based starting message row.</summary>
        public int StartRowIndex { get; }

        /// <summary>Zero-based ending message row.</summary>
        public int EndRowIndex { get; }

        /// <summary>Participant ids spanned by the fragment. Empty means all participants.</summary>
        public IReadOnlyList<string> ParticipantIds { get; }
    }

    /// <summary>
    /// Simple data record used to import sequence fragment guards and partitions into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceFragmentOperandRecord {
        /// <summary>Initializes a sequence fragment operand import record.</summary>
        public VisioSequenceFragmentOperandRecord(string id, string fragmentId, string text, int rowIndex, bool divider = false) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Fragment operand id cannot be null or whitespace.", nameof(id)) : id;
            FragmentId = string.IsNullOrWhiteSpace(fragmentId) ? throw new ArgumentException("Fragment id cannot be null or whitespace.", nameof(fragmentId)) : fragmentId;
            Text = text ?? string.Empty;
            RowIndex = rowIndex;
            Divider = divider;
        }

        /// <summary>Stable fragment operand id.</summary>
        public string Id { get; }

        /// <summary>Target fragment id.</summary>
        public string FragmentId { get; }

        /// <summary>Visible guard or partition label.</summary>
        public string Text { get; }

        /// <summary>Zero-based message row for the guard or partition.</summary>
        public int RowIndex { get; }

        /// <summary>Whether this operand includes a partition divider before the row.</summary>
        public bool Divider { get; }
    }

    /// <summary>
    /// Simple data record used to import semantic sequence notes into <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioSequenceNoteRecord {
        /// <summary>Initializes a sequence note import record.</summary>
        public VisioSequenceNoteRecord(string id, string participantId, string text, int rowIndex, VisioSide placement = VisioSide.Right) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Note id cannot be null or whitespace.", nameof(id)) : id;
            ParticipantId = string.IsNullOrWhiteSpace(participantId) ? throw new ArgumentException("Participant id cannot be null or whitespace.", nameof(participantId)) : participantId;
            Text = text ?? string.Empty;
            RowIndex = rowIndex;
            Placement = placement;
        }

        /// <summary>Stable note id.</summary>
        public string Id { get; }

        /// <summary>Participant id the note should attach to.</summary>
        public string ParticipantId { get; }

        /// <summary>Visible note text.</summary>
        public string Text { get; }

        /// <summary>Zero-based message row near which the note should be placed.</summary>
        public int RowIndex { get; }

        /// <summary>Requested note placement side.</summary>
        public VisioSide Placement { get; }
    }
}
