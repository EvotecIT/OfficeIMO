using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for UML-style sequence diagrams with deterministic participant,
    /// lifeline, and message placement.
    /// </summary>
    public sealed class VisioSequenceDiagramBuilder {
        private const string LifelineLayer = "Sequence Lifelines";
        private const string MessageLayer = "Sequence Messages";
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

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<ParticipantItem> _participants = new();
        private readonly Dictionary<string, ParticipantItem> _participantsById = new(StringComparer.Ordinal);
        private readonly List<MessageItem> _messages = new();
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
            double messageRows = Math.Max(1, _messages.Count);
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
                page.AddLayer(GuideLayer).Print = false;

                AddTitle(page, pageWidth, pageHeight);
                PlaceParticipants(page, pageWidth, headerY, lifelineBottomY);
                AddMessages(page, firstMessageY);
                _document.RequestRecalcOnOpen();
                return page;
            } finally {
                _document.UseMastersByDefault = previousMastersByDefault;
            }
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
            string masterNameU = GetParticipantMaster(participant.Kind);
            VisioShape shape = new(participant.Id, x, y, _participantWidth, _participantHeight, participant.Text) {
                NameU = masterNameU,
            };
            GetParticipantStyle(participant.Kind).ApplyTo(shape);
            shape.SetUserCell(VisioSemanticUserCells.Kind, "SequenceParticipant", "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell("OfficeIMO.SequenceParticipantKind", participant.Kind.ToString(), "STR", prompt: "OfficeIMO sequence participant kind");
            page.Shapes.Add(shape);
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

        private static string GetParticipantMaster(VisioSequenceParticipantKind kind) {
            return kind switch {
                VisioSequenceParticipantKind.Actor => "Circle",
                VisioSequenceParticipantKind.Database => "Data",
                _ => "Rectangle"
            };
        }

        private string NormalizeMessageId(string? id) {
            string normalizedId = string.IsNullOrWhiteSpace(id) ? "message-" + (_messages.Count + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture) : id!.Trim();
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            return normalizedId;
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

            return false;
        }

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
