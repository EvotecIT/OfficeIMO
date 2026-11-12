using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

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
            OfficeOpenXmlHiddenData.InspectImagesAndProvenance(package, collector);
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
            double slideWidth = SlideSize.Width.ToPoints();
            double slideHeight = SlideSize.Height.ToPoints();
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
                    // Placeholders without their own transform take their bounds from the layout; a shape whose
                    // bounds can't be resolved isn't known to be off the slide, so it stays out of the report.
                    if (!shape.TryGetExportBoundsPoints(out double left, out double top, out double width, out double height)) continue;
                    bool offSlide = left + width < 0D || top + height < 0D || left > slideWidth || top > slideHeight;
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

            CollectUnreferencedCommentAuthors(collector);
        }

        // Author lists can name people no comment refers to, for example after comments were deleted by another tool.
        private void CollectUnreferencedCommentAuthors(OfficeHiddenDataCollector collector) {
            PresentationPart presentationPart = _document!.PresentationPart!;
            var classicIds = new HashSet<uint>(_slides.SelectMany(slide =>
                    slide.SlidePart.SlideCommentsPart?.CommentList?.Elements<P.Comment>() ?? Enumerable.Empty<P.Comment>())
                .Where(comment => comment.AuthorId?.Value != null)
                .Select(comment => comment.AuthorId!.Value));
            CommentAuthorsPart? classicPart = presentationPart.CommentAuthorsPart;
            foreach (P.CommentAuthor author in classicPart?.CommentAuthorList?.Elements<P.CommentAuthor>().ToList() ?? new List<P.CommentAuthor>()) {
                uint? id = author.Id?.Value;
                if (id.HasValue && classicIds.Contains(id.Value)) continue;
                collector.Add(OfficeHiddenDataKind.PersonalIdentities, classicPart!.Uri.ToString(), "Comment author no comment refers to", author.Name?.Value,
                    () => { if (id.HasValue) RemoveClassicCommentAuthorIfUnused(id); else author.Remove(); });
            }

            var modernIds = new HashSet<string>(_slides.SelectMany(slide => slide.SlidePart.Parts
                    .Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>())
                .SelectMany(part => part.CommentList?.Descendants().Where(element => element is P188.Comment or P188.CommentReply) ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>())
                .Select(element => element is P188.Comment comment ? comment.AuthorId?.Value : ((P188.CommentReply)element).AuthorId?.Value)
                .Where(id => id != null)!, StringComparer.OrdinalIgnoreCase);
            foreach (PowerPointAuthorsPart part in presentationPart.Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointAuthorsPart>().ToList()) {
                foreach (P188.Author author in part.AuthorList?.Elements<P188.Author>().ToList() ?? new List<P188.Author>()) {
                    string? id = author.Id?.Value;
                    if (id != null && modernIds.Contains(id)) continue;
                    collector.Add(OfficeHiddenDataKind.PersonalIdentities, part.Uri.ToString(), "Comment author no comment refers to", author.Name?.Value,
                        () => { if (id != null) RemoveModernCommentAuthorIfUnused(id); else author.Remove(); });
                }
            }
        }

        private string SlideLocation(PowerPointSlide slide) =>
            "Slide " + (_slides.IndexOf(slide) + 1).ToString(CultureInfo.InvariantCulture);
    }
}
