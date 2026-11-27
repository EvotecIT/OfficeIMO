using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
        private const string CommentsExtendedContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
        private const string CommentsExtendedRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
        private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
        private const string Word2012Namespace = "http://schemas.microsoft.com/office/word/2012/wordml";
        private LegacyDocDocxPart? _footnotes;
        private LegacyDocDocxPart? _endnotes;
        private LegacyDocDocxPart? _comments;
        private OfficeXmlTree? _commentsExtended;
        private int _lastCommentId;
        private uint _lastCommentParagraphId;
        private uint _lastDrawingId;

        // ---- Footnotes and endnotes -----------------------------------------------------------------------------

        private void AddNoteReference(RunTarget target, LegacyDocNoteProjection notes, int? characterPosition, LegacyDocTextRun legacyRun) {
            notes.CancellationToken.ThrowIfCancellationRequested();
            if (characterPosition == null) return;
            if (notes.TryGetFootnote(characterPosition.Value, out LegacyDocFootnote? footnote)) {
                AddNote(target, footnote!.ParagraphRuns, notes, legacyRun, endnote: false);
            } else if (notes.TryGetEndnote(characterPosition.Value, out LegacyDocEndnote? endnote)) {
                AddNote(target, endnote!.ParagraphRuns, notes, legacyRun, endnote: true);
            }
        }

        private void AddNote(RunTarget target, IReadOnlyList<LegacyDocNoteParagraph> paragraphs, LegacyDocNoteProjection notes, LegacyDocTextRun legacyRun, bool endnote) {
            if (paragraphs.Count == 0) return;
            LegacyDocDocxPart part = endnote
                ? _endnotes ??= GetOrCreateMainRelatedPart(RelationshipTypePrefix + "endnotes", "/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", "endnotes")
                : _footnotes ??= GetOrCreateMainRelatedPart(RelationshipTypePrefix + "footnotes", "/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "footnotes");
            OfficeXmlTree notesTree = part.Tree;
            string noteName = endnote ? "endnote" : "footnote";
            string referenceStyle = endnote ? "EndnoteReference" : "FootnoteReference";
            long highestId = 0;
            List<int> existing = Elements(notesTree, notesTree.Root, noteName);
            if (existing.Count == 0) {
                highestId = 1;
            } else {
                foreach (int note in existing) {
                    if (long.TryParse(Get(notesTree, note, "id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long id)) highestId = Math.Max(highestId, id);
                }
            }

            string noteId = Text(highestId <= 0 ? 1 : highestId + 1);

            // The reference in the story.
            OfficeXmlTree tree = target.Tree;
            int reference = Append(tree, target.Container, Element(tree, "r"));
            int referenceProperties = Append(tree, reference, Element(tree, "rPr"));
            Append(tree, referenceProperties, Element(tree, "rStyle", referenceStyle));
            int referenceElement = Append(tree, reference, Element(tree, endnote ? "endnoteReference" : "footnoteReference"));
            Set(tree, referenceElement, "id", noteId);
            ApplyReferenceFormatting(target, reference, legacyRun);

            // The note.
            int noteElement = Element(notesTree, noteName);
            Set(notesTree, noteElement, "id", noteId);
            int first = Append(notesTree, noteElement, Element(notesTree, "p"));
            int firstProperties = Append(notesTree, first, Element(notesTree, "pPr"));
            Append(notesTree, firstProperties, Element(notesTree, "pStyle", endnote ? "EndnoteText" : "FootnoteText"));
            int mark = Append(notesTree, first, Element(notesTree, "r"));
            int markProperties = Append(notesTree, mark, Element(notesTree, "rPr"));
            Append(notesTree, markProperties, Element(notesTree, "rStyle", referenceStyle));
            Append(notesTree, mark, Element(notesTree, endnote ? "endnoteRef" : "footnoteRef"));
            Append(notesTree, notesTree.Root, noteElement);

            LegacyDocNoteProjection inner = notes.WithoutReferences;
            ApplyParagraphFormatting(part, first, paragraphs[0].Format);
            AddNoteParagraphRuns(part, first, paragraphs[0], inner);
            for (int index = 1; index < paragraphs.Count; index++) {
                notes.CancellationToken.ThrowIfCancellationRequested();
                int paragraph = Append(notesTree, noteElement, Element(notesTree, "p"));
                Append(notesTree, paragraph, Element(notesTree, "pPr"));
                ApplyParagraphFormatting(part, paragraph, paragraphs[index].Format);
                AddNoteParagraphRuns(part, paragraph, paragraphs[index], inner);
            }
        }

        private void AddNoteParagraphRuns(LegacyDocDocxPart part, int paragraph, LegacyDocNoteParagraph source, LegacyDocNoteProjection notes) {
            var bookmarks = LegacyDocBookmarkProjection.Create(source.Bookmarks, source.StartCharacter, source.EndCharacter);
            AddRuns(new RunTarget(part, paragraph, paragraph), source.Runs, 0, notes, bookmarks);
            bookmarks.EmitRemaining(part.Tree, paragraph);
        }

        // ---- Comments -------------------------------------------------------------------------------------------

        private void AddCommentReference(RunTarget target, LegacyDocNoteProjection notes, int? characterPosition, LegacyDocTextRun legacyRun) {
            if (characterPosition == null || !notes.TryGetComment(characterPosition.Value, out LegacyDocComment? comment) || comment!.Paragraphs.Count == 0) return;
            string commentId = AddComment(comment);
            OfficeXmlTree tree = target.Tree;
            int paragraph = target.Paragraph;
            // Revisions and fields wrap their runs; the preceding content element keeps the reconstructed range and
            // marker at their source stream position.
            int anchor = OfficeXmlTree.NoNode;
            foreach (int child in Elements(tree, paragraph)) {
                if (tree.NamespaceOf(child) == W && tree.LocalNameOf(child) is "pPr" or "bookmarkStart" or "bookmarkEnd" or "commentRangeStart" or "commentRangeEnd") continue;
                anchor = child;
            }

            if (anchor == OfficeXmlTree.NoNode) anchor = Append(tree, paragraph, Element(tree, "r"));
            int rangeStart = Element(tree, "commentRangeStart");
            Set(tree, rangeStart, "id", commentId);
            tree.InsertBefore(paragraph, rangeStart, anchor);
            int rangeEnd = Element(tree, "commentRangeEnd");
            Set(tree, rangeEnd, "id", commentId);
            InsertAfter(tree, paragraph, rangeEnd, anchor);
            int referenceElement = Element(tree, "commentReference");
            Set(tree, referenceElement, "id", commentId);
            int reference = CreateRun(tree, referenceElement);
            InsertAfter(tree, paragraph, reference, rangeEnd);
            ApplyRunFormatting(target, reference, legacyRun);
            if (legacyRun.Revision.HasValue) {
                tree.Remove(reference);
                AppendRevisionRun(new RunTarget(target.Part, paragraph, paragraph), legacyRun, reference);
            }
        }

        private string AddComment(LegacyDocComment comment) {
            LegacyDocDocxPart part = Comments();
            OfficeXmlTree tree = part.Tree;
            string id = Text(++_lastCommentId);
            string paragraphId = (++_lastCommentParagraphId).ToString("X8", CultureInfo.InvariantCulture);
            int commentElement = Element(tree, "comment");
            Set(tree, commentElement, "id", id);
            Set(tree, commentElement, "author", comment.Author);
            Set(tree, commentElement, "initials", comment.Initials);
            Set(tree, commentElement, "date", System.Xml.XmlConvert.ToString(DateTime.Now, System.Xml.XmlDateTimeSerializationMode.RoundtripKind));
            Append(tree, tree.Root, commentElement);
            int word2010 = tree.Namespaces.GetId(Word2010Namespace, out _);
            for (int index = 0; index < comment.ParagraphRuns.Count; index++) {
                LegacyDocNoteParagraph source = comment.ParagraphRuns[index];
                int paragraph = Append(tree, commentElement, Element(tree, "p"));
                if (index == 0) tree.SetAttribute(paragraph, word2010, "paraId", paragraphId);
                ApplyParagraphFormatting(part, paragraph, source.Format);
                AddNoteParagraphRuns(part, paragraph, source, LegacyDocNoteProjection.Empty);
            }

            OfficeXmlTree extended = CommentsExtended();
            int word2012 = extended.Namespaces.GetId(Word2012Namespace, out _);
            int commentEx = extended.CreateElement(Word2012Namespace, "commentEx", "w15");
            extended.SetAttribute(commentEx, word2012, "paraId", paragraphId);
            Append(extended, extended.Root, commentEx);
            return id;
        }

        private LegacyDocDocxPart Comments() {
            if (_comments != null) return _comments;
            const string name = "/word/comments.xml";
            OfficeXmlTree tree = _package.CreateXmlPart(name, CommentsContentType, W, "comments");
            DeclareWordNamespaces(tree);
            tree.DeclareNamespace(tree.Root, "w14", Word2010Namespace);
            tree.DeclareNamespace(tree.Root, "mc", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.MarkupCompatibility)!);
            tree.SetAttribute(tree.Root, OfficeXmlNamespaces.MarkupCompatibility, "Ignorable", "w14");
            _package.AddRelationship(MainPartName, RelationshipTypePrefix + "comments", name, id: NextRelationshipId());
            return _comments = new LegacyDocDocxPart(name, tree);
        }

        private OfficeXmlTree CommentsExtended() {
            if (_commentsExtended != null) return _commentsExtended;
            const string name = "/word/commentsExtended.xml";
            OfficeXmlTree tree = _package.CreateXmlPart(name, CommentsExtendedContentType, _package.Namespaces.GetId(Word2012Namespace, out _), "commentsEx");
            tree.DeclareNamespace(tree.Root, "w15", Word2012Namespace);
            _package.AddRelationship(MainPartName, CommentsExtendedRelationshipType, name, id: NextRelationshipId());
            return _commentsExtended = tree;
        }

        // ---- Pictures -------------------------------------------------------------------------------------------

        private void AddPicture(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            LegacyDocPicture picture = legacyRun.Picture ?? throw new InvalidOperationException("The legacy DOC picture run has no picture payload.");
            OfficeXmlTree tree = target.Tree;
            bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, 0));
            int drawing = CreatePictureDrawing(target.Part, picture);
            if (target.ImageRun != OfficeXmlTree.NoNode) {
                Append(tree, target.ImageRun, drawing);
                if (legacyRun.Revision.HasValue) {
                    tree.Remove(drawing);
                    if (!Elements(tree, target.ImageRun).Any(child => !Is(tree, child, "rPr"))) tree.Remove(target.ImageRun);
                    AppendRevisionRun(target, legacyRun, CreateRun(tree, drawing));
                }
            } else {
                int run = CreateRun(tree, drawing);
                if (legacyRun.Revision.HasValue) AppendRevisionRun(target, legacyRun, run);
                else Append(tree, target.Container, run);
            }
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private int CreatePictureDrawing(LegacyDocDocxPart part, LegacyDocPicture picture) {
            const string description = "Imported legacy DOC inline picture";
            ImageCharacteristics characteristics;
            using (var stream = new MemoryStream(picture.ImageBytes, writable: false)) {
                characteristics = Helpers.GetImageCharacteristics(stream, picture.FileName);
            }

            if (characteristics.Width <= 0 || characteristics.Height <= 0) throw new InvalidDataException("The stream does not contain a supported image payload.");
            string contentType = OfficeImageInfo.GetMimeType(characteristics.Type);
            string partName = NextImagePartName(ImageExtension(contentType));
            // Image parts are typed by their extension, as the SDK adds them.
            _package.AddDefaultContentType(partName.Substring(partName.LastIndexOf('.') + 1), contentType);
            _package.CreateBinaryPart(partName, contentType: null, picture.ImageBytes);
            string relationshipId = _package.AddRelationship(part.Name, RelationshipTypePrefix + "image", partName, id: NextRelationshipId());
            long width = (long)Math.Round(picture.WidthPixels * 9525.0);
            long height = (long)Math.Round(picture.HeightPixels * 9525.0);
            uint documentPropertiesId = ++_lastDrawingId;
            uint picturePropertiesId = ++_lastDrawingId;
            string imageName = Path.GetFileNameWithoutExtension(picture.FileName);
            string crop = string.Empty;
            if (picture.Crop.HasCrop) {
                crop = "<a:srcRect l=\"" + Text((int)Math.Round(picture.Crop.Left * 100000)) + "\" t=\"" + Text((int)Math.Round(picture.Crop.Top * 100000))
                    + "\" r=\"" + Text((int)Math.Round(picture.Crop.Right * 100000)) + "\" b=\"" + Text((int)Math.Round(picture.Crop.Bottom * 100000)) + "\"/>";
            }

            string xml = "<w:drawing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                + " xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">"
                + "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" wp14:editId=\"" + NewEditId() + "\">"
                + "<wp:extent cx=\"" + Text(width) + "\" cy=\"" + Text(height) + "\"/>"
                + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
                + "<wp:docPr id=\"" + Text(documentPropertiesId) + "\" name=\"" + Escape(imageName) + "\" descr=\"" + description + "\"/>"
                + "<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic>"
                + "<pic:nvPicPr><pic:cNvPr id=\"" + Text(picturePropertiesId) + "\" name=\"" + Escape(picture.FileName) + "\" descr=\"" + description + "\"/><pic:cNvPicPr/></pic:nvPicPr>"
                + "<pic:blipFill><a:blip r:embed=\"" + relationshipId + "\" cstate=\"print\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi val=\"0\"/></a:ext></a:extLst></a:blip>"
                + crop + "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
                + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + Text(width) + "\" cy=\"" + Text(height) + "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
                + "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";
            return ImportXml(part.Tree, xml);
        }

        private string NextImagePartName(string extension) {
            string first = "/media/image" + extension;
            if (!_package.PartExists(first)) return first;
            for (int index = 2; ; index++) {
                string name = "/media/image" + Text(index) + extension;
                if (!_package.PartExists(name)) return name;
            }
        }

        private static string ImageExtension(string contentType) => contentType switch {
            "image/jpeg" => ".jpeg",
            "image/png" => ".png",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            "image/svg+xml" => ".svg",
            "image/x-icon" => ".ico",
            "image/x-pcx" => ".pcx",
            _ => ".bin"
        };

        private static string NewEditId() => Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();

        private static string Escape(string value) => System.Security.SecurityElement.Escape(value) ?? string.Empty;

        /// <summary>Parses an XML fragment and copies its root into <paramref name="tree"/>, detached.</summary>
        private static int ImportXml(OfficeXmlTree tree, string xml) {
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(xml);
            OfficeXmlTree fragment = OfficeXmlTree.Parse(bytes, bytes.Length, tree.Namespaces);
            return tree.ImportElement(fragment, fragment.Root);
        }

        // ---- Text boxes -----------------------------------------------------------------------------------------

        private void AddTextBoxStories(LegacyDocNoteProjection notes) {
            foreach (LegacyDocTextBoxStory story in _source.TextBoxStories) {
                notes.CancellationToken.ThrowIfCancellationRequested();
                if (story.IsHeaderFooterTextBox) continue;
                OfficeXmlTree tree = _main.Tree;
                int paragraph = CreateBodyParagraph();
                AppendBodyBlock(paragraph);
                int run = Append(tree, paragraph, Element(tree, "r"));
                int alternateContent = ImportXml(tree, TextBoxXml(++_lastDrawingId, story.Text));
                Append(tree, run, alternateContent);
                int content = OfficeXmlTree.NoNode;
                foreach (int element in Descendants(tree, alternateContent)) {
                    if (Is(tree, element, "txbxContent")) {
                        content = element;
                        break;
                    }
                }

                if (content == OfficeXmlTree.NoNode || story.Paragraphs.Count == 0) continue;
                RemoveAll(tree, content, "p");
                foreach (LegacyDocNoteParagraph source in story.Paragraphs) {
                    notes.CancellationToken.ThrowIfCancellationRequested();
                    int storyParagraph = Append(tree, content, Element(tree, "p"));
                    ApplyParagraphFormatting(_main, storyParagraph, source.Format);
                    var bookmarks = LegacyDocBookmarkProjection.Create(source.Bookmarks, source.StartCharacter, source.EndCharacter);
                    AddRuns(new RunTarget(_main, storyParagraph, storyParagraph), source.Runs, 0, notes, bookmarks);
                    bookmarks.EmitRemaining(tree, storyParagraph);
                }
            }
        }

        private static string TextBoxXml(uint id, string text) =>
            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
            + " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
            + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
            + " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\""
            + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
            + " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<mc:Choice Requires=\"wps\"><w:drawing>"
            + "<wp:anchor distT=\"91440\" distB=\"91440\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"3E379294\" wp14:editId=\"39C62DE8\">"
            + "<wp:simplePos x=\"0\" y=\"0\"/>"
            + "<wp:positionH relativeFrom=\"page\"><wp:align>center</wp:align></wp:positionH>"
            + "<wp:positionV relativeFrom=\"page\"><wp:posOffset>182880</wp:posOffset></wp:positionV>"
            + "<wp:extent cx=\"2360930\" cy=\"1404620\"/>"
            + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
            + "<wp:wrapSquare wrapText=\"bothSides\"/>"
            + "<wp:docPr id=\"" + Text(id) + "\" name=\"Text Box 2\"/>"
            + "<wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp>"
            + "<wps:cNvSpPr txBox=\"1\"><a:spLocks noChangeArrowheads=\"1\"/></wps:cNvSpPr>"
            + "<wps:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2360930\" cy=\"1404620\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">" + Escape(text) + "</w:t></w:r></w:p></w:txbxContent></wps:txbx>"
            + "<wps:bodyPr rot=\"0\" vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" anchor=\"t\" anchorCtr=\"0\"><a:spAutoFit/></wps:bodyPr>"
            + "</wps:wsp></a:graphicData></a:graphic>"
            + "<wp14:sizeRelH relativeFrom=\"margin\"><wp14:pctWidth>58500</wp14:pctWidth></wp14:sizeRelH>"
            + "<wp14:sizeRelV relativeFrom=\"margin\"><wp14:pctHeight>20000</wp14:pctHeight></wp14:sizeRelV>"
            + "</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>";
    }
}
