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
                EnsureNotesFitPage(page);
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
            VisioSemanticUserCells.MarkGeneratedAdornment(title);
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

                string lifelineId = CreateGeneratedConnectorId(page, participant.Id + "-lifeline");
                VisioConnector lifeline = page.AddConnector(lifelineId, participant.Header, participant.BottomAnchor, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);
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
    }
}
