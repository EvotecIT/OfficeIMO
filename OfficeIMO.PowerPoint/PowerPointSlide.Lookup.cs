using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Retrieves a shape by its name.
        /// </summary>
        public PowerPointShape? GetShape(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return _shapes.FirstOrDefault(s => s.Name == name);
        }

        /// <summary>
        ///     Retrieves a shape by name, optionally using case-insensitive comparison.
        /// </summary>
        public PowerPointShape? GetShapeByName(string name, bool ignoreCase = false) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return _shapes.FirstOrDefault(shape => string.Equals(shape.Name, name, comparison));
        }

        /// <summary>
        ///     Attempts to retrieve a shape by name.
        /// </summary>
        public bool TryGetShapeByName(string name, out PowerPointShape? shape, bool ignoreCase = false) {
            shape = GetShapeByName(name, ignoreCase);
            return shape != null;
        }

        /// <summary>
        ///     Retrieves a typed shape by name, optionally using case-insensitive comparison.
        /// </summary>
        public T? GetShapeByName<T>(string name, bool ignoreCase = false) where T : PowerPointShape {
            return GetShapeByName(name, ignoreCase) as T;
        }

        /// <summary>
        ///     Attempts to retrieve a typed shape by name.
        /// </summary>
        public bool TryGetShapeByName<T>(string name, out T? shape, bool ignoreCase = false) where T : PowerPointShape {
            shape = GetShapeByName<T>(name, ignoreCase);
            return shape != null;
        }

        /// <summary>
        ///     Retrieves a shape by its non-visual drawing identifier.
        /// </summary>
        public PowerPointShape? GetShapeById(uint id) {
            return _shapes.FirstOrDefault(shape => shape.Id == id);
        }

        /// <summary>
        ///     Retrieves a typed shape by its non-visual drawing identifier.
        /// </summary>
        public T? GetShapeById<T>(uint id) where T : PowerPointShape {
            return GetShapeById(id) as T;
        }

        /// <summary>
        ///     Retrieves a textbox by its name.
        /// </summary>
        public PowerPointTextBox? GetTextBox(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return TextBoxes.FirstOrDefault(tb => tb.Name == name);
        }

        /// <summary>
        ///     Retrieves a picture by its name.
        /// </summary>
        public PowerPointPicture? GetPicture(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Pictures.FirstOrDefault(p => p.Name == name);
        }

        /// <summary>
        ///     Retrieves a table by its name.
        /// </summary>
        public PowerPointTable? GetTable(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Tables.FirstOrDefault(t => t.Name == name);
        }

        /// <summary>
        ///     Replaces text across all textboxes on the slide.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue, bool includeTables = true, bool includeNotes = false) {
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            string replacement = newValue ?? string.Empty;
            int count = 0;

            foreach (PowerPointTextBox textBox in TextBoxes) {
                count += textBox.ReplaceText(oldValue, replacement);
            }

            if (includeTables) {
                foreach (PowerPointTable table in Tables) {
                    for (int r = 0; r < table.Rows; r++) {
                        for (int c = 0; c < table.Columns; c++) {
                            count += table.GetCell(r, c).ReplaceText(oldValue, replacement);
                        }
                    }
                }
            }

            if (includeNotes && _slidePart.NotesSlidePart != null) {
                string notesText = Notes.Text ?? string.Empty;
                int occurrences = CountOccurrences(notesText, oldValue);
                if (occurrences > 0) {
                    Notes.Text = notesText.Replace(oldValue, replacement);
                    count += occurrences;
                }
            }

            return count;
        }

        /// <summary>
        ///     Retrieves a chart by its name.
        /// </summary>
        public PowerPointChart? GetChart(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Charts.FirstOrDefault(c => c.Name == name);
        }

        /// <summary>
        ///     Retrieves an embedded OLE object by its shape name.
        /// </summary>
        public PowerPointOleObject? GetOleObject(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }
            return OleObjects.FirstOrDefault(ole => ole.Name == name);
        }

        /// <summary>
        ///     Removes the specified shape from the slide.
        /// </summary>
        public void RemoveShape(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            EnsureShapeOnSlide(shape);
            var knownRelationshipIds = new HashSet<string>(
                _slidePart.Parts.Select(pair => pair.RelationshipId)
                    .Concat(_slidePart.DataPartReferenceRelationships
                        .Select(relationship => relationship.Id))
                    .Concat(_slidePart.ExternalRelationships.Select(
                        relationship => relationship.Id))
                    .Concat(_slidePart.HyperlinkRelationships.Select(
                        relationship => relationship.Id)),
                StringComparer.Ordinal);
            string[] relationshipIds = new[] { shape.Element }
                .Concat(shape.Element.Descendants())
                .SelectMany(element => element.GetAttributes())
                .Select(attribute => attribute.Value)
                .Where(value => value != null
                    && knownRelationshipIds.Contains(value))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            var removedShapeIds = new HashSet<uint>(shape.Element
                .Descendants<NonVisualDrawingProperties>()
                .Select(properties => properties.Id?.Value)
                .Where(id => id.HasValue)
                .Select(id => id!.Value));
            if (shape.Id.HasValue) removedShapeIds.Add(shape.Id.Value);
            PowerPointClassicAnimation[] remainingAnimations =
                ReadClassicAnimations().Where(animation =>
                    !removedShapeIds.Contains(animation.ShapeId)).ToArray();
            if (remainingAnimations.Length != ClassicAnimations.Count) {
                SetClassicAnimations(remainingAnimations);
            }
            shape.Element.Remove();
            _shapes.Remove(shape);
            foreach (string relationshipId in relationshipIds) {
                RemoveSlideRelationshipIfUnused(relationshipId);
            }
        }

        private void RemoveSlideRelationshipIfUnused(
            string relationshipId) {
            if (SlideRoot.GetAttributes().Any(attribute => string.Equals(
                    attribute.Value, relationshipId,
                    StringComparison.Ordinal))
                || SlideRoot.Descendants().Any(element =>
                    element.GetAttributes().Any(attribute => string.Equals(
                        attribute.Value, relationshipId,
                        StringComparison.Ordinal)))) {
                return;
            }
            DataPartReferenceRelationship? dataRelationship = _slidePart
                .DataPartReferenceRelationships.FirstOrDefault(
                    relationship => string.Equals(relationship.Id,
                        relationshipId, StringComparison.Ordinal));
            if (dataRelationship != null) {
                DataPart dataPart = dataRelationship.DataPart;
                _slidePart.DeleteReferenceRelationship(dataRelationship);
                if (!dataPart.GetDataPartReferenceRelationships().Any()
                    && _slidePart.OpenXmlPackage is PresentationDocument
                        document) {
                    document.DeletePart(dataPart);
                }
                return;
            }
            ExternalRelationship? external = _slidePart
                .ExternalRelationships.FirstOrDefault(relationship =>
                    string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
            if (external != null) {
                _slidePart.DeleteReferenceRelationship(external);
                return;
            }
            HyperlinkRelationship? hyperlink = _slidePart
                .HyperlinkRelationships.FirstOrDefault(relationship =>
                    string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
            if (hyperlink != null) {
                _slidePart.DeleteReferenceRelationship(hyperlink);
                return;
            }
            if (_slidePart.TryGetPartById(relationshipId,
                    out _)) {
                _slidePart.DeletePart(relationshipId);
            }
        }

        private static int CountOccurrences(string value, string oldValue) {
            int count = 0;
            int index = 0;
            while (true) {
                index = value.IndexOf(oldValue, index, StringComparison.Ordinal);
                if (index < 0) {
                    break;
                }
                count++;
                index += oldValue.Length;
            }
            return count;
        }
    }
}
