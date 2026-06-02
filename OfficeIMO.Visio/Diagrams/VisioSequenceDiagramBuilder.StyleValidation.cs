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

        private VisioTextStyle CreateFragmentOperandTextStyle() {
            VisioTextStyle style = CreateFragmentLabelTextStyle();
            style.Size = Math.Min(style.Size ?? 8D, 8D);
            style.Bold = false;
            return style;
        }

        private VisioShape CreateAnchor(VisioPage page, string id, double x, double y) {
            VisioShape anchor = new(id, x.ToInches(_unit), y.ToInches(_unit), 0.04D.ToInches(_unit), 0.04D.ToInches(_unit), string.Empty) {
                NameU = "Circle",
                FillPattern = 0,
                LinePattern = 0,
                FillColor = OfficeColor.Transparent,
                LineColor = OfficeColor.Transparent
            };
            VisioSemanticUserCells.MarkGeneratedAdornment(anchor);
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

            EnsureGeneratedIdsAvailable(normalizedId, nameof(id), normalizedId + "-from", normalizedId + "-to");
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

        private string NormalizeFragmentOperandId(string fragmentId, string? id, bool divider) {
            int existingCount = _fragmentOperands.Count(item => string.Equals(item.FragmentId, fragmentId, StringComparison.Ordinal));
            string normalizedId = string.IsNullOrWhiteSpace(id)
                ? fragmentId + "-operand-" + (existingCount + 1).ToString(global::System.Globalization.CultureInfo.InvariantCulture)
                : id!.Trim();
            string labelId = GetFragmentOperandLabelId(normalizedId);
            if (IsIdInUse(normalizedId) || IsIdInUse(labelId)) {
                throw new ArgumentException($"A sequence diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            if (divider) {
                EnsureGeneratedIdsAvailable(normalizedId, nameof(id), normalizedId + "-from", normalizedId + "-to");
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

        private FragmentItem RequireFragment(string fragmentId) {
            string normalizedId = RequireId(fragmentId, nameof(fragmentId), "Fragment id");
            FragmentItem? fragment = _fragments.FirstOrDefault(item => string.Equals(item.Id, normalizedId, StringComparison.Ordinal));
            if (fragment == null) {
                throw new ArgumentException($"Unknown sequence fragment id '{normalizedId}'.", nameof(fragmentId));
            }

            return fragment;
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

            foreach (ParticipantItem participant in _participants) {
                if (string.Equals(participant.Id + "-lifeline", id, StringComparison.Ordinal) ||
                    string.Equals(participant.Id + "-lifeline-end", id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (MessageItem message in _messages) {
                if (string.Equals(message.Id, id, StringComparison.Ordinal) ||
                    string.Equals(message.Id + "-from", id, StringComparison.Ordinal) ||
                    string.Equals(message.Id + "-to", id, StringComparison.Ordinal)) {
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

            foreach (FragmentOperandItem operand in _fragmentOperands) {
                if (string.Equals(operand.Id, id, StringComparison.Ordinal) ||
                    string.Equals(GetFragmentOperandLabelId(operand.Id), id, StringComparison.Ordinal) ||
                    string.Equals(operand.Id + "-from", id, StringComparison.Ordinal) ||
                    string.Equals(operand.Id + "-to", id, StringComparison.Ordinal)) {
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

        private void EnsureGeneratedIdsAvailable(string ownerId, string parameterName, params string[] generatedIds) {
            foreach (string generatedId in generatedIds) {
                if (IsIdInUse(generatedId)) {
                    throw new ArgumentException($"A sequence diagram item with id '{ownerId}' would create generated item id '{generatedId}' that already exists.", parameterName);
                }
            }
        }

        private string CreateGeneratedConnectorId(VisioPage page, string baseId) {
            string candidate = baseId;
            int index = 2;
            while (page.Connectors.Any(connector => string.Equals(connector.Id, candidate, StringComparison.Ordinal))) {
                candidate = baseId + "-" + index.ToString(global::System.Globalization.CultureInfo.InvariantCulture);
                index++;
            }

            return candidate;
        }

        private static string GetFragmentLabelId(string fragmentId) => fragmentId + "-label";

        private static string GetFragmentOperandLabelId(string operandId) => operandId + "-label";

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
