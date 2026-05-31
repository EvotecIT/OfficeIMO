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
    public sealed class VisioSequenceDiagramBuilder {
        private const string LifelineLayer = "Sequence Lifelines";
        private const string MessageLayer = "Sequence Messages";
        private const string ActivationLayer = "Sequence Activations";
        private const string FragmentLayer = "Sequence Fragments";
        private const string NoteLayer = "Sequence Notes";
        private const string GuideLayer = "Sequence Guides";

        private sealed class ParticipantItem {
            public ParticipantItem(string id, string text, VisioSequenceParticipantKind kind) {
                Id = id;
                Text = text;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioSequenceParticipantKind Kind { get; }

            public double PinX { get; set; }

            public VisioShape? Header { get; set; }

            public VisioShape? BottomAnchor { get; set; }
        }

        private sealed class MessageItem {
            public MessageItem(string id, string fromId, string toId, string label, VisioSequenceMessageKind kind, bool selfMessage) {
                Id = id;
                FromId = fromId;
                ToId = toId;
                Label = label;
                Kind = kind;
                SelfMessage = selfMessage;
            }

            public string Id { get; }

            public string FromId { get; }

            public string ToId { get; }

            public string Label { get; }

            public VisioSequenceMessageKind Kind { get; }

            public bool SelfMessage { get; }
        }

        private sealed class NoteItem {
            public NoteItem(string id, string participantId, string text, int rowIndex, VisioSide placement) {
                Id = id;
                ParticipantId = participantId;
                Text = text;
                RowIndex = rowIndex;
                Placement = placement;
            }

            public string Id { get; }

            public string ParticipantId { get; }

            public string Text { get; }

            public int RowIndex { get; }

            public VisioSide Placement { get; }
        }

        private sealed class ActivationItem {
            public ActivationItem(string id, string participantId, int startRowIndex, int endRowIndex) {
                Id = id;
                ParticipantId = participantId;
                StartRowIndex = startRowIndex;
                EndRowIndex = endRowIndex;
            }

            public string Id { get; }

            public string ParticipantId { get; }

            public int StartRowIndex { get; }

            public int EndRowIndex { get; }
        }

        private sealed class FragmentItem {
            public FragmentItem(string id, string text, int startRowIndex, int endRowIndex, IReadOnlyList<string> participantIds) {
                Id = id;
                Text = text;
                StartRowIndex = startRowIndex;
                EndRowIndex = endRowIndex;
                ParticipantIds = participantIds;
            }

            public string Id { get; }

            public string Text { get; }

            public int StartRowIndex { get; }

            public int EndRowIndex { get; }

            public IReadOnlyList<string> ParticipantIds { get; }
        }

        private readonly struct LayoutBounds {
            public LayoutBounds(double left, double top, double right, double bottom) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public double Left { get; }

            public double Top { get; }

            public double Right { get; }

            public double Bottom { get; }

            public static LayoutBounds FromCenter(double x, double y, double width, double height) {
                double halfWidth = width / 2D;
                double halfHeight = height / 2D;
                return new LayoutBounds(x - halfWidth, y + halfHeight, x + halfWidth, y - halfHeight);
            }

            public LayoutBounds Inflate(double padding) =>
                new(Left - padding, Top + padding, Right + padding, Bottom - padding);
        }

        private readonly struct NotePlacement {
            public NotePlacement(double x, double y, VisioSide resolvedPlacement, LayoutBounds bounds) {
                X = x;
                Y = y;
                ResolvedPlacement = resolvedPlacement;
                Bounds = bounds;
            }

            public double X { get; }

            public double Y { get; }

            public VisioSide ResolvedPlacement { get; }

            public LayoutBounds Bounds { get; }
        }

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<ParticipantItem> _participants = new();
        private readonly Dictionary<string, ParticipantItem> _participantsById = new(StringComparer.Ordinal);
        private readonly List<MessageItem> _messages = new();
        private readonly List<ActivationItem> _activations = new();
        private readonly List<FragmentItem> _fragments = new();
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
        private const double ActivationWidth = 0.16D;
        private const double ActivationMinimumHeight = 0.32D;
        private const double FragmentHorizontalPadding = 0.28D;
        private const double FragmentVerticalPadding = 0.22D;
        private const double FragmentHeaderHeight = 0.3D;
        private const double FragmentMinimumWidth = 1.2D;
        private const double FragmentMinimumHeight = 0.72D;
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
            _fragments.Add(new FragmentItem(fragmentId, text ?? string.Empty, startRowIndex, endRowIndex, normalizedParticipantIds));
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

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This sequence diagram builder has already produced a page.");
            }

            _built = true;
            if (_participants.Count == 0) {
                throw new InvalidOperationException("A sequence diagram requires at least one participant.");
            }

            double titleBand = string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;
            double requiredWidth = _leftMargin + _rightMargin + (_participants.Count * _participantWidth) + ((_participants.Count - 1) * _participantGap);
            double messageRows = Math.Max(1, GetRequiredRowCount());
            double requiredHeight = _topMargin + titleBand + _participantHeight + _messageGap + (messageRows * _messageSpacing) + _bottomMargin;
            double pageWidth = Math.Max(_pageWidth, requiredWidth);
            double pageHeight = Math.Max(_pageHeight, requiredHeight);
            double headerY = pageHeight - _topMargin - titleBand - (_participantHeight / 2D);
            double firstMessageY = headerY - (_participantHeight / 2D) - _messageGap;
            double lifelineBottomY = Math.Max(_bottomMargin, firstMessageY - (messageRows * _messageSpacing));

            bool previousMastersByDefault = _document.UseMastersByDefault;
            _document.UseMastersByDefault = false;
            try {
                VisioPage page = _document.AddPage(_pageName, pageWidth, pageHeight, _unit);
                page.Grid(visible: false, snap: true);
                page.AddLayer(LifelineLayer);
                page.AddLayer(MessageLayer);
                if (_activations.Count > 0) {
                    page.AddLayer(ActivationLayer);
                }

                if (_fragments.Count > 0) {
                    page.AddLayer(FragmentLayer);
                }

                if (_notes.Count > 0) {
                    page.AddLayer(NoteLayer);
                }

                page.AddLayer(GuideLayer).Print = false;

                AddTitle(page, pageWidth, pageHeight);
                PlaceParticipants(page, pageWidth, headerY, lifelineBottomY);
                AddActivations(page, firstMessageY);
                AddMessages(page, firstMessageY);
                AddFragments(page, firstMessageY);
                AddNotes(page, firstMessageY);
                _document.RequestRecalcOnOpen();
                return page;
            } finally {
                _document.UseMastersByDefault = previousMastersByDefault;
            }
        }

        private int GetRequiredRowCount() {
            int rows = _messages.Count;
            foreach (ActivationItem activation in _activations) {
                rows = Math.Max(rows, activation.EndRowIndex + 1);
            }

            foreach (NoteItem note in _notes) {
                rows = Math.Max(rows, note.RowIndex + 1);
            }

            foreach (FragmentItem fragment in _fragments) {
                rows = Math.Max(rows, fragment.EndRowIndex + 1);
            }

            return rows;
        }

        private void AddTitle(VisioPage page, double pageWidth, double pageHeight) {
            if (string.IsNullOrWhiteSpace(_titleText)) {
                return;
            }

            double y = pageHeight - _topMargin - (_titleHeight / 2D);
            VisioShape title = page.AddTextBox(_titleId, pageWidth / 2D, y, Math.Max(1D, pageWidth - _leftMargin - _rightMargin), _titleHeight, _titleText, _unit);
            title.TextStyle = VisioDiagramTitleStyles.Create(_theme);
        }

        private void PlaceParticipants(VisioPage page, double pageWidth, double headerY, double lifelineBottomY) {
            double totalWidth = (_participants.Count * _participantWidth) + ((_participants.Count - 1) * _participantGap);
            double startX = Math.Max(_leftMargin + (_participantWidth / 2D), (pageWidth - totalWidth) / 2D + (_participantWidth / 2D));

            for (int i = 0; i < _participants.Count; i++) {
                ParticipantItem participant = _participants[i];
                double x = startX + i * (_participantWidth + _participantGap);
                participant.PinX = x;
                participant.Header = CreateParticipantHeader(page, participant, x, headerY);
                participant.BottomAnchor = CreateAnchor(page, participant.Id + "-lifeline-end", x, lifelineBottomY);

                VisioConnector lifeline = page.AddConnector(participant.Header, participant.BottomAnchor, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);
                lifeline.LineColor = _theme.Connector.LineColor;
                lifeline.LineWeight = Math.Max(0.006D, _theme.Connector.LineWeight * 0.75D);
                lifeline.LinePattern = 2;
                lifeline.EndArrow = EndArrow.None;
                page.AddToLayer(LifelineLayer, lifeline);
            }
        }

        private VisioShape CreateParticipantHeader(VisioPage page, ParticipantItem participant, double x, double y) {
            VisioShape shape = page.AddStencilShape(VisioStencils.Sequence, GetParticipantStencilId(participant.Kind), participant.Id, x, y, _participantWidth, _participantHeight, participant.Text);
            GetParticipantStyle(participant.Kind).ApplyTo(shape);
            shape.SetUserCell(VisioSemanticUserCells.Kind, "SequenceParticipant", "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell("OfficeIMO.SequenceParticipantKind", participant.Kind.ToString(), "STR", prompt: "OfficeIMO sequence participant kind");
            page.AddToLayer(LifelineLayer, shape);
            return shape;
        }

        private void AddMessages(VisioPage page, double firstMessageY) {
            for (int i = 0; i < _messages.Count; i++) {
                MessageItem message = _messages[i];
                double y = firstMessageY - (i * _messageSpacing);
                if (message.SelfMessage) {
                    AddSelfMessage(page, message, y);
                } else {
                    AddParticipantMessage(page, message, y);
                }
            }
        }

        private void AddActivations(VisioPage page, double firstMessageY) {
            foreach (ActivationItem activation in _activations) {
                ParticipantItem participant = _participantsById[activation.ParticipantId];
                double topY = firstMessageY - (activation.StartRowIndex * _messageSpacing) + (_messageSpacing * 0.32D);
                double bottomY = firstMessageY - (activation.EndRowIndex * _messageSpacing) - (_messageSpacing * 0.32D);
                double height = Math.Max(ActivationMinimumHeight, topY - bottomY);
                double y = (topY + bottomY) / 2D;
                VisioShape shape = page.AddStencilShape(VisioStencils.Sequence, "seq.activation", activation.Id, participant.PinX, y, ActivationWidth, height, string.Empty);
                _theme.Marker.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.SequenceActivationKind, "STR", prompt: "OfficeIMO semantic kind");
                shape.SetUserCell("OfficeIMO.SequenceParticipantId", activation.ParticipantId, "STR", prompt: "OfficeIMO sequence activation participant");
                shape.SetUserCell("OfficeIMO.SequenceStartRowIndex", activation.StartRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence activation start row");
                shape.SetUserCell("OfficeIMO.SequenceEndRowIndex", activation.EndRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence activation end row");
                page.AddToLayer(ActivationLayer, shape);
            }
        }

        private void AddFragments(VisioPage page, double firstMessageY) {
            foreach (FragmentItem fragment in _fragments) {
                IReadOnlyList<ParticipantItem> participants = fragment.ParticipantIds
                    .Select(id => _participantsById[id])
                    .ToArray();
                double left = participants.Min(participant => participant.PinX) - (_participantWidth / 2D) - FragmentHorizontalPadding;
                double right = participants.Max(participant => participant.PinX) + (_participantWidth / 2D) + FragmentHorizontalPadding;
                left = Math.Max(_leftMargin, left);
                right = Math.Min(page.Width - _rightMargin, right);

                double topY = firstMessageY - (fragment.StartRowIndex * _messageSpacing) + (_messageSpacing * 0.46D) + FragmentVerticalPadding;
                double bottomY = firstMessageY - (fragment.EndRowIndex * _messageSpacing) - (_messageSpacing * 0.46D) - FragmentVerticalPadding;
                double width = Math.Max(FragmentMinimumWidth, right - left);
                double height = Math.Max(FragmentMinimumHeight, topY - bottomY);
                double x = left + (width / 2D);
                double y = (topY + bottomY) / 2D;

                VisioShape frame = page.AddStencilShape(VisioStencils.Sequence, "seq.fragment", fragment.Id, x, y, width, height, string.Empty);
                frame.FillColor = OfficeColor.Transparent;
                frame.FillPattern = 0;
                frame.LineColor = _theme.ControlConnector.LineColor;
                frame.LinePattern = 2;
                frame.LineWeight = Math.Max(0.012D, _theme.ControlConnector.LineWeight);
                frame.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.SequenceFragmentKind, "STR", prompt: "OfficeIMO semantic kind");
                frame.SetUserCell("OfficeIMO.SequenceStartRowIndex", fragment.StartRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment start row");
                frame.SetUserCell("OfficeIMO.SequenceEndRowIndex", fragment.EndRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment end row");
                frame.SetUserCell("OfficeIMO.SequenceParticipantIds", string.Join(";", fragment.ParticipantIds), "STR", prompt: "OfficeIMO sequence fragment participants");
                page.AddToLayer(FragmentLayer, frame);

                double labelWidth = Math.Max(0.75D, Math.Min(width - 0.1D, 0.48D + (fragment.Text.Trim().Length * 0.07D)));
                double labelLeft = left + 0.04D;
                double firstParticipantX = participants.Min(participant => participant.PinX);
                if (labelLeft < firstParticipantX + ActivationWidth && labelLeft + labelWidth > firstParticipantX - ActivationWidth) {
                    labelLeft = firstParticipantX + (ActivationWidth / 2D) + 0.08D;
                    labelLeft = Math.Min(labelLeft, left + width - labelWidth - 0.04D);
                }

                double labelX = labelLeft + (labelWidth / 2D);
                double labelY = topY - 0.04D - (FragmentHeaderHeight / 2D);
                VisioShape label = page.AddTextBox(GetFragmentLabelId(fragment.Id), labelX, labelY, labelWidth, FragmentHeaderHeight, fragment.Text, _unit);
                label.FillColor = _theme.Container.FillColor;
                label.FillPattern = 1;
                label.LineColor = OfficeColor.Transparent;
                label.LinePattern = 0;
                label.TextStyle = CreateFragmentLabelTextStyle();
                label.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");
                label.SetUserCell("OfficeIMO.SequenceFragmentId", fragment.Id, "STR", prompt: "OfficeIMO sequence fragment label target");
                page.AddToLayer(FragmentLayer, label);
            }
        }

        private void AddParticipantMessage(VisioPage page, MessageItem message, double y) {
            ParticipantItem from = _participantsById[message.FromId];
            ParticipantItem to = _participantsById[message.ToId];
            bool leftToRight = from.PinX <= to.PinX;
            VisioShape fromAnchor = CreateAnchor(page, message.Id + "-from", from.PinX, y);
            VisioShape toAnchor = CreateAnchor(page, message.Id + "-to", to.PinX, y);
            VisioConnector connector = page.AddConnector(
                fromAnchor,
                toAnchor,
                ConnectorKind.Straight,
                leftToRight ? VisioSide.Right : VisioSide.Left,
                leftToRight ? VisioSide.Left : VisioSide.Right);

            connector.Id = message.Id;
            ApplyMessageStyle(connector, message.Kind);
            connector.Label = message.Label;
            double labelX = (from.PinX + to.PinX) / 2D;
            double labelWidth = Math.Max(1.2D, Math.Min(2.6D, Math.Abs(to.PinX - from.PinX) - 0.2D));
            connector.PlaceLabelAt(labelX, y + 0.16D, labelWidth, 0.28D);
            page.AddToLayer(MessageLayer, connector);
        }

        private void AddSelfMessage(VisioPage page, MessageItem message, double y) {
            ParticipantItem participant = _participantsById[message.FromId];
            double lowerY = y - Math.Min(_selfMessageHeight, _messageSpacing * 0.55D);
            VisioShape fromAnchor = CreateAnchor(page, message.Id + "-from", participant.PinX, y);
            VisioShape toAnchor = CreateAnchor(page, message.Id + "-to", participant.PinX, lowerY);
            ResolveSelfMessageLabelPlacement(page, participant, message.Label, out double direction, out double labelWidth, out double labelHeight);
            VisioSide connectorSide = direction > 0D ? VisioSide.Right : VisioSide.Left;
            double loopX = participant.PinX + (direction * _selfMessageWidth);
            VisioConnector connector = page.AddConnector(fromAnchor, toAnchor, ConnectorKind.RightAngle, connectorSide, connectorSide);
            connector.Id = message.Id;
            ApplyMessageStyle(connector, message.Kind);
            connector.Label = message.Label;
            connector.RouteThrough(
                VisioConnectorWaypoint.At(loopX, y),
                VisioConnectorWaypoint.At(loopX, lowerY));
            double labelCenterX = loopX + (direction * (SelfMessageLabelGap + (labelWidth / 2D)));
            double labelCenterY = y - (labelHeight / 2D);
            connector.PlaceLabelAt(labelCenterX, labelCenterY, labelWidth, labelHeight);
            page.AddToLayer(MessageLayer, connector);
        }

        private void AddNotes(VisioPage page, double firstMessageY) {
            List<LayoutBounds> reservedBounds = BuildNoteReservedBounds(page);
            foreach (NoteItem note in _notes) {
                NotePlacement placement = ResolveNotePlacement(page, note, firstMessageY, reservedBounds);

                VisioShape shape = page.AddStencilShape(VisioStencils.Sequence, "seq.note", note.Id, placement.X, placement.Y, NoteWidth, NoteHeight, note.Text);
                _theme.Container.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, "SequenceNote", "STR", prompt: "OfficeIMO semantic kind");
                shape.SetUserCell("OfficeIMO.SequenceParticipantId", note.ParticipantId, "STR", prompt: "OfficeIMO sequence note target participant");
                shape.SetUserCell("OfficeIMO.SequenceRowIndex", note.RowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence note row");
                shape.SetUserCell("OfficeIMO.SequenceRequestedPlacement", note.Placement.ToString(), "STR", prompt: "OfficeIMO requested sequence note placement");
                shape.SetUserCell("OfficeIMO.SequenceResolvedPlacement", placement.ResolvedPlacement.ToString(), "STR", prompt: "OfficeIMO resolved sequence note placement");
                page.AddToLayer(NoteLayer, shape);
                reservedBounds.Add(placement.Bounds.Inflate(NoteCollisionPadding));
            }
        }

        private NotePlacement ResolveNotePlacement(VisioPage page, NoteItem note, double firstMessageY, IReadOnlyList<LayoutBounds> reservedBounds) {
            ParticipantItem participant = _participantsById[note.ParticipantId];
            double baseY = firstMessageY - (note.RowIndex * _messageSpacing) - (_messageSpacing / 2D);
            double minY = _bottomMargin + (NoteHeight / 2D);
            double maxY = page.Height - _topMargin - _participantHeight - NoteGap - (NoteHeight / 2D);
            VisioSide[] sides = note.Placement == VisioSide.Right
                ? new[] { VisioSide.Right, VisioSide.Left }
                : new[] { VisioSide.Left, VisioSide.Right };

            NotePlacement bestPlacement = default;
            double bestScore = double.PositiveInfinity;
            bool hasBest = false;
            foreach (VisioSide side in sides) {
                double x = GetNoteX(page, participant, side);
                for (int offsetIndex = 0; offsetIndex <= 12; offsetIndex++) {
                    foreach (double offset in GetNoteCandidateOffsets(offsetIndex)) {
                        double y = Math.Max(minY, Math.Min(maxY, baseY + offset));
                        LayoutBounds bounds = LayoutBounds.FromCenter(x, y, NoteWidth, NoteHeight);
                        double score = ScoreNotePlacement(bounds.Inflate(NoteCollisionPadding), reservedBounds) +
                                       (Math.Abs(offset) * 0.9D) +
                                       (side == note.Placement ? 0D : 1.8D);

                        if (score < bestScore) {
                            bestScore = score;
                            bestPlacement = new NotePlacement(x, y, side, bounds);
                            hasBest = true;
                        }

                        if (score <= 0.0001D) {
                            return bestPlacement;
                        }
                    }
                }
            }

            return hasBest
                ? bestPlacement
                : new NotePlacement(GetNoteX(page, participant, note.Placement), Math.Max(minY, Math.Min(maxY, baseY)), note.Placement, LayoutBounds.FromCenter(GetNoteX(page, participant, note.Placement), Math.Max(minY, Math.Min(maxY, baseY)), NoteWidth, NoteHeight));
        }

        private static IEnumerable<double> GetNoteCandidateOffsets(int offsetIndex) {
            if (offsetIndex == 0) {
                yield return 0D;
            } else {
                double offset = offsetIndex * NoteVerticalCandidateStep;
                yield return offset;
                yield return -offset;
            }
        }

        private double GetNoteX(VisioPage page, ParticipantItem participant, VisioSide placement) {
            double direction = placement == VisioSide.Right ? 1D : -1D;
            double x = participant.PinX + (direction * ((_participantWidth / 2D) + NoteGap + (NoteWidth / 2D)));
            return Math.Max(_leftMargin + (NoteWidth / 2D), Math.Min(page.Width - _rightMargin - (NoteWidth / 2D), x));
        }

        private static double ScoreNotePlacement(LayoutBounds candidate, IReadOnlyList<LayoutBounds> reservedBounds) {
            double score = 0D;
            foreach (LayoutBounds reserved in reservedBounds) {
                score += GetOverlapArea(candidate, reserved) * 25D;
            }

            return score;
        }

        private static double GetOverlapArea(LayoutBounds first, LayoutBounds second) {
            double width = Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left);
            double height = Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom);
            return width > 0D && height > 0D ? width * height : 0D;
        }

        private List<LayoutBounds> BuildNoteReservedBounds(VisioPage page) {
            List<LayoutBounds> reservedBounds = new();
            foreach (VisioShape shape in page.Shapes) {
                if (IsIgnoredNoteObstacle(shape)) {
                    continue;
                }

                reservedBounds.Add(LayoutBounds.FromCenter(shape.PinX, shape.PinY, shape.Width, shape.Height).Inflate(NoteCollisionPadding));
            }

            foreach (VisioConnector connector in page.Connectors) {
                AddConnectorSegmentReservedBounds(connector, reservedBounds);
                VisioConnectorLabelPlacement? label = connector.LabelPlacement;
                if (label != null && label.PinX.HasValue && label.PinY.HasValue) {
                    reservedBounds.Add(LayoutBounds.FromCenter(label.PinX.Value, label.PinY.Value, label.Width, label.Height).Inflate(NoteCollisionPadding));
                }
            }

            return reservedBounds;
        }

        private static void AddConnectorSegmentReservedBounds(VisioConnector connector, IList<LayoutBounds> reservedBounds) {
            List<(double X, double Y)> points = new() {
                (connector.From.PinX, connector.From.PinY)
            };
            foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                points.Add((waypoint.X, waypoint.Y));
            }

            if (connector.Waypoints.Count == 0 && connector.Kind == ConnectorKind.RightAngle) {
                points.Add((connector.From.PinX, connector.To.PinY));
            }

            points.Add((connector.To.PinX, connector.To.PinY));
            const double thickness = 0.08D;
            for (int i = 0; i < points.Count - 1; i++) {
                double left = Math.Min(points[i].X, points[i + 1].X) - (thickness / 2D);
                double right = Math.Max(points[i].X, points[i + 1].X) + (thickness / 2D);
                double bottom = Math.Min(points[i].Y, points[i + 1].Y) - (thickness / 2D);
                double top = Math.Max(points[i].Y, points[i + 1].Y) + (thickness / 2D);
                reservedBounds.Add(new LayoutBounds(left, top, right, bottom));
            }
        }

        private static bool IsIgnoredNoteObstacle(VisioShape shape) {
            if (shape.Width <= 0.06D && shape.Height <= 0.06D) {
                return true;
            }

            string? kind = shape.GetUserCellValue(VisioSemanticUserCells.Kind);
            return string.Equals(kind, VisioSemanticUserCells.SequenceFragmentKind, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(kind, VisioSemanticUserCells.BackgroundSurfaceKind, StringComparison.OrdinalIgnoreCase);
        }

        private void ResolveSelfMessageLabelPlacement(VisioPage page, ParticipantItem participant, string label, out double direction, out double labelWidth, out double labelHeight) {
            double desiredWidth = EstimateSelfMessageLabelWidth(label);
            double rightAvailable = GetSelfMessageLabelAvailableWidth(page, participant, rightSide: true);
            double leftAvailable = GetSelfMessageLabelAvailableWidth(page, participant, rightSide: false);
            bool useRight = rightAvailable >= Math.Min(desiredWidth, 1.1D) || rightAvailable >= leftAvailable;
            double available = Math.Max(0D, useRight ? rightAvailable : leftAvailable);

            direction = useRight ? 1D : -1D;
            labelWidth = Math.Max(0.9D, Math.Min(desiredWidth, available));
            labelHeight = desiredWidth > labelWidth + 0.05D ? 0.46D : SelfMessageLabelHeight;
        }

        private double GetSelfMessageLabelAvailableWidth(VisioPage page, ParticipantItem participant, bool rightSide) {
            double direction = rightSide ? 1D : -1D;
            double loopX = participant.PinX + (direction * _selfMessageWidth);
            double pageLimit = rightSide
                ? page.Width - _rightMargin - loopX - SelfMessageLabelGap
                : loopX - _leftMargin - SelfMessageLabelGap;
            double nearestParticipantLimit = double.PositiveInfinity;

            foreach (ParticipantItem other in _participants) {
                if (ReferenceEquals(other, participant)) {
                    continue;
                }

                if (rightSide && other.PinX > participant.PinX) {
                    nearestParticipantLimit = Math.Min(nearestParticipantLimit, other.PinX - loopX - SelfMessageLabelGap - 0.18D);
                } else if (!rightSide && other.PinX < participant.PinX) {
                    nearestParticipantLimit = Math.Min(nearestParticipantLimit, loopX - other.PinX - SelfMessageLabelGap - 0.18D);
                }
            }

            return Math.Max(0D, Math.Min(pageLimit, nearestParticipantLimit));
        }

        private static double EstimateSelfMessageLabelWidth(string label) {
            if (string.IsNullOrWhiteSpace(label)) {
                return 1.2D;
            }

            double estimatedWidth = 0.55D + (label.Trim().Length * 0.075D);
            return Math.Max(1.2D, Math.Min(2.4D, estimatedWidth));
        }

        private VisioTextStyle CreateFragmentLabelTextStyle() {
            VisioTextStyle style = _theme.Container.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily ??= "Aptos";
            style.Color ??= _theme.ControlConnector.TextStyle?.Color ?? OfficeColor.Black;
            style.Size ??= 9D;
            style.Bold = true;
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Left;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            style.LeftMargin = 0.06D;
            style.RightMargin = 0.04D;
            style.TopMargin = 0.02D;
            style.BottomMargin = 0.02D;
            style.BackgroundColor = _theme.Container.FillColor;
            style.BackgroundTransparency = 0D;
            return style;
        }

        private VisioShape CreateAnchor(VisioPage page, string id, double x, double y) {
            VisioShape anchor = new(id, x, y, 0.04D, 0.04D, string.Empty) {
                NameU = "Circle",
                FillPattern = 0,
                LinePattern = 0,
                FillColor = OfficeColor.Transparent,
                LineColor = OfficeColor.Transparent
            };
            anchor.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");
            page.Shapes.Add(anchor);
            page.AddToLayer(GuideLayer, anchor);
            return anchor;
        }

        private void ApplyMessageStyle(VisioConnector connector, VisioSequenceMessageKind kind) {
            VisioConnectorStyle style = kind switch {
                VisioSequenceMessageKind.Async => _theme.DataConnector,
                VisioSequenceMessageKind.Return => _theme.ControlConnector,
                VisioSequenceMessageKind.Event => _theme.Marker.TextStyle != null
                    ? new VisioConnectorStyle(_theme.Marker.LineColor, _theme.Marker.LineWeight, 1, EndArrow.Arrow) { TextStyle = _theme.Marker.TextStyle.Clone() }
                    : _theme.Connector,
                _ => _theme.Connector
            };

            connector.LineColor = style.LineColor;
            connector.LineWeight = style.LineWeight;
            connector.LinePattern = kind == VisioSequenceMessageKind.Return ? 2 : style.LinePattern;
            connector.EndArrow = kind == VisioSequenceMessageKind.Call ? EndArrow.Triangle : EndArrow.Arrow;
            connector.BeginArrow = EndArrow.None;
            if (style.TextStyle != null) {
                connector.TextStyle = style.TextStyle.Clone();
            }
        }

        private VisioShapeStyle GetParticipantStyle(VisioSequenceParticipantKind kind) {
            return kind switch {
                VisioSequenceParticipantKind.Actor => _theme.Success,
                VisioSequenceParticipantKind.Boundary => _theme.Marker,
                VisioSequenceParticipantKind.Control => _theme.Emphasis,
                VisioSequenceParticipantKind.Entity => _theme.Decision,
                VisioSequenceParticipantKind.Database => _theme.Container,
                _ => _theme.Primary
            };
        }

        private static string GetParticipantStencilId(VisioSequenceParticipantKind kind) {
            return kind switch {
                VisioSequenceParticipantKind.Actor => "seq.actor",
                VisioSequenceParticipantKind.Boundary => "seq.boundary",
                VisioSequenceParticipantKind.Control => "seq.control",
                VisioSequenceParticipantKind.Entity => "seq.entity",
                VisioSequenceParticipantKind.Database => "seq.database",
                _ => "seq.participant"
            };
        }

        private string NormalizeMessageId(string? id) {
            string normalizedId = string.IsNullOrWhiteSpace(id) ? "message-" + (_messages.Count + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture) : id!.Trim();
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            return normalizedId;
        }

        private string NormalizeNoteId(string? id) {
            string normalizedId = string.IsNullOrWhiteSpace(id) ? "note-" + (_notes.Count + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture) : id!.Trim();
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            return normalizedId;
        }

        private string NormalizeFragmentId(string? id) {
            string normalizedId = string.IsNullOrWhiteSpace(id) ? "fragment-" + (_fragments.Count + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture) : id!.Trim();
            string labelId = GetFragmentLabelId(normalizedId);
            if (IsIdInUse(normalizedId) || IsIdInUse(labelId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            return normalizedId;
        }

        private string NormalizeActivationId(string? id) {
            string normalizedId = string.IsNullOrWhiteSpace(id) ? "activation-" + (_activations.Count + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture) : id!.Trim();
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            return normalizedId;
        }

        private IReadOnlyList<string> GetFragmentParticipantIds(IEnumerable<string> participantIds) {
            string[] ids = participantIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (ids.Length == 0) {
                ids = _participants.Select(participant => participant.Id).ToArray();
            }

            if (ids.Length == 0) {
                throw new InvalidOperationException("A sequence fragment requires at least one participant.");
            }

            foreach (string id in ids) {
                EnsureKnownParticipant(id, nameof(participantIds));
            }

            return ids;
        }

        private void EnsureKnownParticipant(string id, string parameterName) {
            string normalizedId = RequireId(id, parameterName, "Participant id");
            if (!_participantsById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown sequence participant id '{normalizedId}'.", parameterName);
            }
        }

        private bool IsIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(id, _titleId, StringComparison.Ordinal)) {
                return true;
            }

            if (_participantsById.ContainsKey(id)) {
                return true;
            }

            foreach (MessageItem message in _messages) {
                if (string.Equals(message.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (ActivationItem activation in _activations) {
                if (string.Equals(activation.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (FragmentItem fragment in _fragments) {
                if (string.Equals(fragment.Id, id, StringComparison.Ordinal) || string.Equals(GetFragmentLabelId(fragment.Id), id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (NoteItem note in _notes) {
                if (string.Equals(note.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private static string GetFragmentLabelId(string fragmentId) => fragmentId + "-label";

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
        }

        private static void ValidatePositive(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite positive number.");
            }
        }

        private static void ValidateNonNegative(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite non-negative number.");
            }
        }
    }
}
