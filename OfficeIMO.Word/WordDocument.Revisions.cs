using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Handles revisions within the document.
    /// </summary>
    public partial class WordDocument {

        /// <summary>
        /// Given author name, accept all revisions by given Author
        /// </summary>
        /// <param name="authorName"></param>
        public void AcceptRevisions(string authorName) {
            // Given a document name and an author name, accept revisions. 
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            // Handle the formatting changes.
            List<OpenXmlElement> changes = body.Descendants<ParagraphPropertiesChange>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();

            foreach (OpenXmlElement change in changes) {
                change.Remove();
            }

            // Handle the deletions.
            List<OpenXmlElement> deletions = body.Descendants<Deleted>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();

            deletions.AddRange(body.Descendants<DeletedRun>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());

            deletions.AddRange(body.Descendants<DeletedMathControl>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement deletion in deletions) {
                deletion.Remove();
            }

            // Handle the insertions.
            List<OpenXmlElement> insertions = body.Descendants<Inserted>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();

            insertions.AddRange(body.Descendants<InsertedRun>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());

            insertions.AddRange(body.Descendants<InsertedMathControl>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement insertion in insertions) {
                // Found new content.
                // Promote them to the same level as node, and then delete the node.
                foreach (var run in insertion.Elements<Run>()) {
                    if (run == insertion.FirstChild) {
                        insertion.InsertAfterSelf(new Run(run.OuterXml));
                    } else {
                        var nextSibling = insertion.NextSibling() ?? throw new InvalidOperationException("Insertion has no next sibling.");
                        nextSibling.InsertAfterSelf(new Run(run.OuterXml));
                    }
                }

                insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.Remove();
            }
        }

        /// <summary>
        /// Accept all revisions in the document
        /// </summary>
        public void AcceptRevisions() {
            // Given a document name and an author name, accept revisions.
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            // Handle the formatting changes.
            List<OpenXmlElement> changes = body.Descendants<ParagraphPropertiesChange>().Cast<OpenXmlElement>().ToList();

            foreach (OpenXmlElement change in changes) {
                change.Remove();
            }

            // Handle the deletions.
            List<OpenXmlElement> deletions = body.Descendants<Deleted>().Cast<OpenXmlElement>().ToList();

            deletions.AddRange(body.Descendants<DeletedRun>().Cast<OpenXmlElement>().ToList());

            deletions.AddRange(body.Descendants<DeletedMathControl>().Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement deletion in deletions) {
                deletion.Remove();
            }

            // Handle the insertions.
            List<OpenXmlElement> insertions = body.Descendants<Inserted>().Cast<OpenXmlElement>().ToList();

            insertions.AddRange(body.Descendants<InsertedRun>().Cast<OpenXmlElement>().ToList());

            insertions.AddRange(body.Descendants<InsertedMathControl>().Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement insertion in insertions) {
                // Found new content.
                // Promote them to the same level as node, and then delete the node.
                foreach (var run in insertion.Elements<Run>()) {
                    if (run == insertion.FirstChild) {
                        insertion.InsertAfterSelf(new Run(run.OuterXml));
                    } else {
                        var nextSibling = insertion.NextSibling() ?? throw new InvalidOperationException("Insertion has no next sibling.");
                        nextSibling.InsertAfterSelf(new Run(run.OuterXml));
                    }
                }

                insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.Remove();
            }
        }

        /// <summary>
        /// Converts tracked revisions into visible markup by replacing revision
        /// elements with formatted runs. Inserted text is underlined and colored
        /// blue, while deleted text is displayed with red strikethrough.
        /// </summary>
        public void ConvertRevisionsToMarkup() {
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            // Process insertions
            foreach (var insertion in body.Descendants<InsertedRun>().ToList()) {
                var parent = insertion.Parent ?? throw new InvalidOperationException("Insertion has no parent.");
                OpenXmlElement last = insertion;
                foreach (var run in insertion.Elements<Run>().Select(r => (Run)r.CloneNode(true))) {
                    var rPr = run.RunProperties ?? new RunProperties();
                    rPr.Color = new Color() { Val = "0000FF" };
                    rPr.Underline = new Underline() { Val = UnderlineValues.Single };
                    run.RunProperties = rPr;
                    parent.InsertAfter(run, last);
                    last = run;
                }
                insertion.Remove();
            }

            // Process deletions
            foreach (var deletion in body.Descendants<DeletedRun>().ToList()) {
                var parent = deletion.Parent ?? throw new InvalidOperationException("Deletion has no parent.");
                OpenXmlElement last = deletion;
                foreach (var run in deletion.Elements<Run>().Select(r => (Run)r.CloneNode(true))) {
                    var rPr = run.RunProperties ?? new RunProperties();
                    rPr.Color = new Color() { Val = "FF0000" };
                    rPr.Strike = new Strike();
                    run.RunProperties = rPr;
                    parent.InsertAfter(run, last);
                    last = run;
                }
                deletion.Remove();
            }
        }

        /// <summary>
        /// Reject all revisions by given author
        /// </summary>
        /// <param name="authorName"></param>
        public void RejectRevisions(string authorName) {
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            // formatting changes
            List<OpenXmlElement> changes = body.Descendants<ParagraphPropertiesChange>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();
            foreach (OpenXmlElement change in changes) {
                change.Remove();
            }

            // insertions are removed
            List<OpenXmlElement> insertions = body.Descendants<Inserted>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();
            insertions.AddRange(body.Descendants<InsertedRun>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());
            insertions.AddRange(body.Descendants<InsertedMathControl>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());
            foreach (OpenXmlElement insertion in insertions) {
                insertion.Remove();
            }

            // deletions are promoted
            List<OpenXmlElement> deletions = body.Descendants<Deleted>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList();
            deletions.AddRange(body.Descendants<DeletedRun>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());
            deletions.AddRange(body.Descendants<DeletedMathControl>().Where(c => c.Author?.Value == authorName).Cast<OpenXmlElement>().ToList());
            foreach (OpenXmlElement deletion in deletions) {
                foreach (var run in deletion.Elements<Run>()) {
                    if (run == deletion.FirstChild) {
                        deletion.InsertAfterSelf(new Run(run.OuterXml));
                    } else {
                        var nextSibling = deletion.NextSibling() ?? throw new InvalidOperationException("Deletion has no next sibling.");
                        nextSibling.InsertAfterSelf(new Run(run.OuterXml));
                    }
                }
                deletion.RemoveAttribute("rsidDel", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                deletion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                deletion.Remove();
            }
        }

        /// <summary>
        /// Reject all revisions in the document
        /// </summary>
        public void RejectRevisions() {
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            List<OpenXmlElement> changes = body.Descendants<ParagraphPropertiesChange>().Cast<OpenXmlElement>().ToList();
            foreach (OpenXmlElement change in changes) {
                change.Remove();
            }

            List<OpenXmlElement> insertions = body.Descendants<Inserted>().Cast<OpenXmlElement>().ToList();
            insertions.AddRange(body.Descendants<InsertedRun>().Cast<OpenXmlElement>().ToList());
            insertions.AddRange(body.Descendants<InsertedMathControl>().Cast<OpenXmlElement>().ToList());
            foreach (OpenXmlElement insertion in insertions) {
                insertion.Remove();
            }

            List<OpenXmlElement> deletions = body.Descendants<Deleted>().Cast<OpenXmlElement>().ToList();
            deletions.AddRange(body.Descendants<DeletedRun>().Cast<OpenXmlElement>().ToList());
            deletions.AddRange(body.Descendants<DeletedMathControl>().Cast<OpenXmlElement>().ToList());
            foreach (OpenXmlElement deletion in deletions) {
                foreach (var run in deletion.Elements<Run>()) {
                    if (run == deletion.FirstChild) {
                        deletion.InsertAfterSelf(new Run(run.OuterXml));
                    } else {
                        var nextSibling = deletion.NextSibling() ?? throw new InvalidOperationException("Deletion has no next sibling.");
                        nextSibling.InsertAfterSelf(new Run(run.OuterXml));
                    }
                }
                deletion.RemoveAttribute("rsidDel", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                deletion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                deletion.Remove();
            }
        }
    }
}