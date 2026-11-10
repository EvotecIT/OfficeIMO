using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Finds metadata, personal information and hidden data, the way Office's Document Inspector does, without changing the presentation.</summary>
        public OfficeHiddenDataReport InspectHiddenData() {
            ThrowIfDisposed();
            return CollectHiddenData(null);
        }

        /// <summary>Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item. Save the presentation to keep the result.</summary>
        /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps hidden slides and off-slide content.</param>
        public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) {
            ThrowIfDisposed();
            return CollectHiddenData(options ?? new OfficeHiddenDataOptions());
        }

        private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
            PresentationDocument package = _document ?? throw new ObjectDisposedException(nameof(PowerPointDocument));
            PresentationPart presentationPart = package.PresentationPart ?? throw new InvalidOperationException("PresentationPart is missing.");
            var collector = new OfficeHiddenDataCollector("PowerPoint", removal);

            CollectPowerPointComments(collector);
            CollectPowerPointHiddenContent(collector);
            CollectPowerPointSpeakerNotes(collector);
            CollectPowerPointIdentities(collector);
            OfficeOpenXmlHiddenData.InspectDocumentProperties(package, collector);
            OfficeOpenXmlHiddenData.InspectCustomProperties(package, collector, null);
            OfficeOpenXmlHiddenData.InspectCustomXml(presentationPart, presentationPart.CustomXmlParts, collector);
            OfficeOpenXmlHiddenData.InspectThumbnail(package, collector);
            OfficeOpenXmlHiddenData.InspectImageMetadata(package, collector);
            OfficeOpenXmlHiddenData.InspectPrinterSettings(package, collector);
            return collector.Build();
        }

        private void CollectPowerPointComments(OfficeHiddenDataCollector collector) {
            foreach (PowerPointSlide slide in _slides.ToList()) {
                string location = SlideLocation(slide);
                foreach (PowerPointClassicComment comment in GetClassicComments(slide)) {
                    collector.Add(OfficeHiddenDataKind.Comments, location, "Comment by " + comment.Author.Name, comment.Text, comment.Remove);
                }
                foreach (PowerPointModernComment comment in GetModernComments(slide)) {
                    collector.Add(OfficeHiddenDataKind.Comments, location, "Comment by " + comment.Author.Name, comment.Text, comment.Remove);
                }
            }
        }

        private void CollectPowerPointHiddenContent(OfficeHiddenDataCollector collector) {
            long slideWidth = SlideSize.Width.ToEmus();
            long slideHeight = SlideSize.Height.ToEmus();
            foreach (PowerPointSlide slide in _slides.ToList()) {
                string location = SlideLocation(slide);
                if (slide.Hidden) {
                    collector.Add(OfficeHiddenDataKind.HiddenContent, location, "Hidden slide", null, () => RemoveSlide(_slides.IndexOf(slide)));
                    if (collector.Removes(OfficeHiddenDataKind.HiddenContent)) continue;
                }

                foreach (PowerPointShape shape in slide.Shapes.ToList()) {
                    string shapeLocation = location + ", shape '" + shape.Name + "'";
                    if (shape.Hidden) {
                        collector.Add(OfficeHiddenDataKind.HiddenContent, shapeLocation, "Hidden shape", null, shape.Remove);
                        continue;
                    }
                    long left = shape.Left.ToEmus();
                    long top = shape.Top.ToEmus();
                    bool offSlide = left + shape.Width.ToEmus() <= 0 || top + shape.Height.ToEmus() <= 0 || left >= slideWidth || top >= slideHeight;
                    if (offSlide) {
                        collector.Add(OfficeHiddenDataKind.HiddenContent, shapeLocation, "Shape placed off the slide", null, shape.Remove);
                    }
                }
            }
        }

        private void CollectPowerPointSpeakerNotes(OfficeHiddenDataCollector collector) {
            foreach (PowerPointSlide slide in _slides) {
                SlidePart slidePart = slide.SlidePart;
                NotesSlidePart? notes = slidePart.NotesSlidePart;
                if (notes == null) continue;
                slide.Notes.TryGetExistingText(out string text);
                collector.Add(OfficeHiddenDataKind.SpeakerNotes, SlideLocation(slide), "Speaker notes", text,
                    () => slidePart.DeletePart(notes));
            }
        }

        private void CollectPowerPointIdentities(OfficeHiddenDataCollector collector) {
            var authored = new List<(string Name, string Location, Action<PowerPointCommentAuthor> SetAuthor)>();
            foreach (PowerPointSlide slide in _slides) {
                string location = SlideLocation(slide);
                foreach (PowerPointClassicComment comment in GetClassicComments(slide)) authored.Add((comment.Author.Name, location, comment.SetAuthor));
                foreach (PowerPointModernComment comment in GetModernComments(slide)) {
                    authored.Add((comment.Author.Name, location, comment.SetAuthor));
                    foreach (PowerPointModernCommentReply reply in comment.Replies) authored.Add((reply.Author.Name, location, reply.SetAuthor));
                }
            }

            foreach (IGrouping<string, (string Name, string Location, Action<PowerPointCommentAuthor> SetAuthor)> group in authored.GroupBy(item => item.Name, StringComparer.Ordinal)) {
                if (collector.Removal != null && group.Key == collector.Removal.AnonymousAuthor) continue;
                string locations = string.Join(", ", group.Select(item => item.Location).Distinct());
                collector.Add(OfficeHiddenDataKind.PersonalIdentities, locations,
                    "Comment author on " + group.Count().ToString(CultureInfo.InvariantCulture) + " comments", group.Key, () => {
                        var anonymous = new PowerPointCommentAuthor(collector.Removal!.AnonymousAuthor, collector.Removal.AnonymousInitials);
                        foreach ((_, _, Action<PowerPointCommentAuthor> setAuthor) in group) setAuthor(anonymous);
                    }, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
            }
        }

        private string SlideLocation(PowerPointSlide slide) =>
            "Slide " + (_slides.IndexOf(slide) + 1).ToString(CultureInfo.InvariantCulture);
    }
}
