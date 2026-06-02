using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for UML-style sequence diagrams with deterministic participant,
    /// lifeline, and message placement.
    /// </summary>
    public sealed partial class VisioSequenceDiagramBuilder {
        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<ParticipantItem> _participants = new();
        private readonly Dictionary<string, ParticipantItem> _participantsById = new(StringComparer.Ordinal);
        private readonly List<MessageItem> _messages = new();
        private readonly List<ActivationItem> _activations = new();
        private readonly List<FragmentItem> _fragments = new();
        private readonly List<FragmentOperandItem> _fragmentOperands = new();
        private readonly List<NoteItem> _notes = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.85;
        private double _rightMargin = 0.85;
        private double _topMargin = 0.7;
        private double _bottomMargin = 0.7;
        private double _participantWidth = 1.45;
        private double _participantHeight = 0.62;
        private double _participantGap = 1.15;
        private double _messageGap = 0.55;
        private double _messageSpacing = 0.62;
        private double _selfMessageWidth = 0.75;
        private double _selfMessageHeight = 0.36;
        private const double SelfMessageLabelGap = 0.18D;
        private const double SelfMessageLabelHeight = 0.3D;
        private const double MessageLabelHeight = 0.28D;
        private const double MessageLabelVerticalOffset = 0.16D;
        private const double MessageLabelLifelinePadding = 0.14D;
        private const double MessageLabelActivationPadding = 0.08D;
        private const double ActivationWidth = 0.16D;
        private const double ActivationMinimumHeight = 0.32D;
        private const double FragmentHorizontalPadding = 0.28D;
        private const double FragmentVerticalPadding = 0.22D;
        private const double FragmentHeaderHeight = 0.3D;
        private const double FragmentOperandLabelHeight = 0.24D;
        private const double FragmentMinimumWidth = 1.2D;
        private const double FragmentMinimumHeight = 0.72D;
        private const double FragmentNestedHorizontalInset = 0.18D;
        private const double FragmentNestedVerticalInset = 0.08D;
        private const double NoteWidth = 1.85D;
        private const double NoteHeight = 0.72D;
        private const double NoteGap = 0.22D;
        private const double NoteCollisionPadding = 0.08D;
        private const double NoteVerticalCandidateStep = 0.26D;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _built;

        internal VisioSequenceDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Sequence Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated sequence diagram page.</summary>
        public VisioSequenceDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioSequenceDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by the automatic layout.</summary>
        public VisioSequenceDiagramBuilder Margins(double left, double top, double right = 0.85, double bottom = 0.7) {
            ValidateNonNegative(left, nameof(left));
            ValidateNonNegative(top, nameof(top));
            ValidateNonNegative(right, nameof(right));
            ValidateNonNegative(bottom, nameof(bottom));
            _leftMargin = left;
            _topMargin = top;
            _rightMargin = right;
            _bottomMargin = bottom;
            return this;
        }

        /// <summary>Sets participant header size.</summary>
        public VisioSequenceDiagramBuilder ParticipantSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _participantWidth = width;
            _participantHeight = height;
            return this;
        }

        /// <summary>Sets participant and message spacing.</summary>
        public VisioSequenceDiagramBuilder Spacing(double participantGap = 1.15, double messageSpacing = 0.62, double messageGap = 0.55) {
            ValidateNonNegative(participantGap, nameof(participantGap));
            ValidatePositive(messageSpacing, nameof(messageSpacing));
            ValidateNonNegative(messageGap, nameof(messageGap));
            _participantGap = participantGap;
            _messageSpacing = messageSpacing;
            _messageGap = messageGap;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a generic participant.</summary>
        public VisioSequenceDiagramBuilder Participant(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Participant);

        /// <summary>Adds an actor participant.</summary>
        public VisioSequenceDiagramBuilder Actor(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Actor);

        /// <summary>Adds a system/service participant.</summary>
        public VisioSequenceDiagramBuilder System(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Participant);

        /// <summary>Adds a boundary/interface participant.</summary>
        public VisioSequenceDiagramBuilder Boundary(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Boundary);

        /// <summary>Adds a control/coordinator participant.</summary>
        public VisioSequenceDiagramBuilder Control(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Control);

        /// <summary>Adds an entity participant.</summary>
        public VisioSequenceDiagramBuilder Entity(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Entity);

        /// <summary>Adds a database participant.</summary>
        public VisioSequenceDiagramBuilder Database(string id, string text) => Participant(id, text, VisioSequenceParticipantKind.Database);

        /// <summary>Adds a participant with an explicit semantic kind.</summary>
        public VisioSequenceDiagramBuilder Participant(string id, string text, VisioSequenceParticipantKind kind) {
            string normalizedId = RequireId(id, nameof(id), "Participant id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            EnsureGeneratedIdsAvailable(normalizedId, nameof(id), normalizedId + "-lifeline", normalizedId + "-lifeline-end");
            if (!Enum.IsDefined(typeof(VisioSequenceParticipantKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            ParticipantItem participant = new(normalizedId, text ?? string.Empty, kind);
            _participants.Add(participant);
            _participantsById.Add(normalizedId, participant);
            return this;
        }

        /// <summary>Adds a synchronous call message.</summary>
        public VisioSequenceDiagramBuilder Call(string fromId, string toId, string label, string? id = null) =>
            Message(fromId, toId, label, VisioSequenceMessageKind.Call, id);

        /// <summary>Adds an asynchronous message.</summary>
        public VisioSequenceDiagramBuilder Async(string fromId, string toId, string label, string? id = null) =>
            Message(fromId, toId, label, VisioSequenceMessageKind.Async, id);

        /// <summary>Adds a return/response message.</summary>
        public VisioSequenceDiagramBuilder Return(string fromId, string toId, string label, string? id = null) =>
            Message(fromId, toId, label, VisioSequenceMessageKind.Return, id);

        /// <summary>Adds a message between two known participants.</summary>
        public VisioSequenceDiagramBuilder Message(string fromId, string toId, string label, VisioSequenceMessageKind kind = VisioSequenceMessageKind.Call, string? id = null) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "Participant id");
            string normalizedToId = RequireId(toId, nameof(toId), "Participant id");
            EnsureKnownParticipant(normalizedFromId, nameof(fromId));
            EnsureKnownParticipant(normalizedToId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioSequenceMessageKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            string messageId = NormalizeMessageId(id);
            _messages.Add(new MessageItem(messageId, normalizedFromId, normalizedToId, label ?? string.Empty, kind, selfMessage: false));
            return this;
        }

        /// <summary>Adds a self-message loop on a known participant.</summary>
        public VisioSequenceDiagramBuilder SelfMessage(string participantId, string label, VisioSequenceMessageKind kind = VisioSequenceMessageKind.Call, string? id = null) {
            string normalizedParticipantId = RequireId(participantId, nameof(participantId), "Participant id");
            EnsureKnownParticipant(normalizedParticipantId, nameof(participantId));
            if (!Enum.IsDefined(typeof(VisioSequenceMessageKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            string messageId = NormalizeMessageId(id);
            _messages.Add(new MessageItem(messageId, normalizedParticipantId, normalizedParticipantId, label ?? string.Empty, kind, selfMessage: true));
            return this;
        }

        /// <summary>Adds an execution/focus activation bar on a known participant over a range of message rows.</summary>
        public VisioSequenceDiagramBuilder Activation(string participantId, int startRowIndex, int endRowIndex, string? id = null) {
            string normalizedParticipantId = RequireId(participantId, nameof(participantId), "Participant id");
            EnsureKnownParticipant(normalizedParticipantId, nameof(participantId));
            if (startRowIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index must be zero or greater.");
            }

            if (endRowIndex < startRowIndex) {
                throw new ArgumentOutOfRangeException(nameof(endRowIndex), "End row index must be greater than or equal to the start row index.");
            }

            string activationId = NormalizeActivationId(id);
            _activations.Add(new ActivationItem(activationId, normalizedParticipantId, startRowIndex, endRowIndex));
            return this;
        }

        /// <summary>Adds a UML combined fragment spanning all known participants over a range of message rows.</summary>
        public VisioSequenceDiagramBuilder Fragment(string text, int startRowIndex, int endRowIndex, string? id = null) =>
            Fragment(text, startRowIndex, endRowIndex, Array.Empty<string>(), id);

        /// <summary>Adds a UML combined fragment spanning selected participants over a range of message rows.</summary>
        public VisioSequenceDiagramBuilder Fragment(string text, int startRowIndex, int endRowIndex, IEnumerable<string> participantIds, string? id = null) {
            AddFragment(text, startRowIndex, endRowIndex, participantIds, id, parentFragmentId: null);
            return this;
        }

        /// <summary>Adds a nested UML combined fragment inside an existing fragment.</summary>
        public VisioSequenceDiagramBuilder NestedFragment(string parentFragmentId, string text, int startRowIndex, int endRowIndex, string? id = null) {
            FragmentItem parent = RequireFragment(parentFragmentId);
            AddFragment(text, startRowIndex, endRowIndex, parent.ParticipantIds, id, parent.Id);
            return this;
        }

        /// <summary>Adds a nested UML combined fragment spanning selected participants inside an existing fragment.</summary>
        public VisioSequenceDiagramBuilder NestedFragment(string parentFragmentId, string text, int startRowIndex, int endRowIndex, IEnumerable<string> participantIds, string? id = null) {
            FragmentItem parent = RequireFragment(parentFragmentId);
            AddFragment(text, startRowIndex, endRowIndex, participantIds, id, parent.Id);
            return this;
        }

        private void AddFragment(string text, int startRowIndex, int endRowIndex, IEnumerable<string> participantIds, string? id, string? parentFragmentId) {
            if (participantIds == null) {
                throw new ArgumentNullException(nameof(participantIds));
            }

            if (startRowIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index must be zero or greater.");
            }

            if (endRowIndex < startRowIndex) {
                throw new ArgumentOutOfRangeException(nameof(endRowIndex), "End row index must be greater than or equal to the start row index.");
            }

            string fragmentId = NormalizeFragmentId(id);
            IReadOnlyList<string> normalizedParticipantIds = GetFragmentParticipantIds(participantIds);
            if (!string.IsNullOrWhiteSpace(parentFragmentId)) {
                FragmentItem parent = RequireFragment(parentFragmentId!);
                if (startRowIndex < parent.StartRowIndex || endRowIndex > parent.EndRowIndex) {
                    throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Nested fragment row range must be inside the parent fragment row range.");
                }

                foreach (string participantId in normalizedParticipantIds) {
                    if (!parent.ParticipantIds.Contains(participantId, StringComparer.Ordinal)) {
                        throw new ArgumentException($"Nested fragment participant id '{participantId}' is outside parent fragment '{parent.Id}'.", nameof(participantIds));
                    }
                }
            }

            _fragments.Add(new FragmentItem(fragmentId, text ?? string.Empty, startRowIndex, endRowIndex, normalizedParticipantIds, parentFragmentId));
        }

        /// <summary>Adds a guard label inside an existing UML combined fragment.</summary>
        public VisioSequenceDiagramBuilder FragmentGuard(string fragmentId, string text, int rowIndex, string? id = null) {
            FragmentItem fragment = RequireFragment(fragmentId);
            if (rowIndex < fragment.StartRowIndex || rowIndex > fragment.EndRowIndex) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Guard row index must be inside the fragment row range.");
            }

            string operandId = NormalizeFragmentOperandId(fragment.Id, id, divider: false);
            _fragmentOperands.Add(new FragmentOperandItem(operandId, fragment.Id, text ?? string.Empty, rowIndex, divider: false));
            return this;
        }

        /// <summary>Adds an operand divider and guard label inside an existing UML combined fragment.</summary>
        public VisioSequenceDiagramBuilder FragmentPartition(string fragmentId, string text, int beforeRowIndex, string? id = null) {
            FragmentItem fragment = RequireFragment(fragmentId);
            if (beforeRowIndex <= fragment.StartRowIndex || beforeRowIndex > fragment.EndRowIndex) {
                throw new ArgumentOutOfRangeException(nameof(beforeRowIndex), "Partition row index must be inside the fragment and after the first fragment row.");
            }

            string operandId = NormalizeFragmentOperandId(fragment.Id, id, divider: true);
            _fragmentOperands.Add(new FragmentOperandItem(operandId, fragment.Id, text ?? string.Empty, beforeRowIndex, divider: true));
            return this;
        }

        /// <summary>Adds a semantic note near a participant at a message row.</summary>
        public VisioSequenceDiagramBuilder Note(string participantId, string text, int rowIndex, VisioSide placement = VisioSide.Right, string? id = null) {
            string normalizedParticipantId = RequireId(participantId, nameof(participantId), "Participant id");
            EnsureKnownParticipant(normalizedParticipantId, nameof(participantId));
            if (rowIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be zero or greater.");
            }

            if (placement != VisioSide.Left && placement != VisioSide.Right) {
                throw new ArgumentOutOfRangeException(nameof(placement), "Sequence notes must be placed to the left or right of a participant.");
            }

            string noteId = NormalizeNoteId(id);
            _notes.Add(new NoteItem(noteId, normalizedParticipantId, text ?? string.Empty, rowIndex, placement));
            return this;
        }

        /// <summary>Imports participant records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Participants(IEnumerable<VisioSequenceParticipantRecord> participants) {
            if (participants == null) {
                throw new ArgumentNullException(nameof(participants));
            }

            foreach (VisioSequenceParticipantRecord participant in participants) {
                Participant(participant.Id, participant.Text, participant.Kind);
            }

            return this;
        }

        /// <summary>Imports message records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Messages(IEnumerable<VisioSequenceMessageRecord> messages) {
            if (messages == null) {
                throw new ArgumentNullException(nameof(messages));
            }

            foreach (VisioSequenceMessageRecord message in messages) {
                if (message.SelfMessage || string.Equals(message.FromId, message.ToId, StringComparison.Ordinal)) {
                    SelfMessage(message.FromId, message.Label, message.Kind, message.Id);
                } else {
                    Message(message.FromId, message.ToId, message.Label, message.Kind, message.Id);
                }
            }

            return this;
        }

        /// <summary>Imports activation records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Activations(IEnumerable<VisioSequenceActivationRecord> activations) {
            if (activations == null) {
                throw new ArgumentNullException(nameof(activations));
            }

            foreach (VisioSequenceActivationRecord activation in activations) {
                Activation(activation.ParticipantId, activation.StartRowIndex, activation.EndRowIndex, activation.Id);
            }

            return this;
        }

        /// <summary>Imports combined-fragment records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Fragments(IEnumerable<VisioSequenceFragmentRecord> fragments) {
            if (fragments == null) {
                throw new ArgumentNullException(nameof(fragments));
            }

            foreach (VisioSequenceFragmentRecord fragment in fragments) {
                if (string.IsNullOrWhiteSpace(fragment.ParentFragmentId)) {
                    Fragment(fragment.Text, fragment.StartRowIndex, fragment.EndRowIndex, fragment.ParticipantIds, fragment.Id);
                } else {
                    NestedFragment(fragment.ParentFragmentId!, fragment.Text, fragment.StartRowIndex, fragment.EndRowIndex, fragment.ParticipantIds, fragment.Id);
                }
            }

            return this;
        }

        /// <summary>Imports fragment guard and partition records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder FragmentOperands(IEnumerable<VisioSequenceFragmentOperandRecord> operands) {
            if (operands == null) {
                throw new ArgumentNullException(nameof(operands));
            }

            foreach (VisioSequenceFragmentOperandRecord operand in operands) {
                if (operand.Divider) {
                    FragmentPartition(operand.FragmentId, operand.Text, operand.RowIndex, operand.Id);
                } else {
                    FragmentGuard(operand.FragmentId, operand.Text, operand.RowIndex, operand.Id);
                }
            }

            return this;
        }

        /// <summary>Imports semantic note records into the sequence diagram.</summary>
        public VisioSequenceDiagramBuilder Notes(IEnumerable<VisioSequenceNoteRecord> notes) {
            if (notes == null) {
                throw new ArgumentNullException(nameof(notes));
            }

            foreach (VisioSequenceNoteRecord note in notes) {
                Note(note.ParticipantId, note.Text, note.RowIndex, note.Placement, note.Id);
            }

            return this;
        }

        /// <summary>Imports sequence participants, messages, activations, fragments, fragment operands, and notes from simple data records.</summary>
        public VisioSequenceDiagramBuilder Import(
            IEnumerable<VisioSequenceParticipantRecord> participants,
            IEnumerable<VisioSequenceMessageRecord> messages,
            IEnumerable<VisioSequenceActivationRecord>? activations = null,
            IEnumerable<VisioSequenceFragmentRecord>? fragments = null,
            IEnumerable<VisioSequenceFragmentOperandRecord>? fragmentOperands = null,
            IEnumerable<VisioSequenceNoteRecord>? notes = null) {
            Participants(participants);
            Messages(messages);
            if (activations != null) {
                Activations(activations);
            }

            if (fragments != null) {
                Fragments(fragments);
            }

            if (fragmentOperands != null) {
                FragmentOperands(fragmentOperands);
            }

            if (notes != null) {
                Notes(notes);
            }

            return this;
        }
    }
}
