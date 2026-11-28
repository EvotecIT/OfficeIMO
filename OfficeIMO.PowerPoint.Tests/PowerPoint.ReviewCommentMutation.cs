using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.Tests {
    public class PowerPointReviewCommentMutationTests {
        [Fact]
        public void ClassicCommentReadsIgnoreExtensionDecoysAndKeepHeldCommentsEditable() {
            const int p = OfficeXmlNamespaces.PresentationML;
            const int noNamespace = OfficeXmlNamespaces.None;
            var created = new DateTime(2026, 8, 1, 7, 30, 0, DateTimeKind.Utc);
            var reviewer = new PowerPointCommentAuthor("Review author", "RA");
            using var original = PowerPointDocument.Create();
            PowerPointSlide first = original.AddSlide();
            original.AddSlide();
            original.AddClassicComment(first, reviewer, "First review", created: created);
            original.AddClassicComment(first, reviewer, "Review text", global::OfficeIMO.OfficeLength.Points(-12),
                global::OfficeIMO.OfficeLength.Points(43), created);

            byte[] originalBytes = original.ToBytes();
            using var input = OfficeOpcPackage.Open(new MemoryStream(originalBytes, writable: false));
            using var package = OfficeOpcEditablePackage.Open(new MemoryStream(originalBytes, writable: false));
            string main = input.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slide = input.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            string comments = input.GetRelatedPartName(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments")!;
            string authors = input.GetRelatedPartName(main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "commentAuthors")!;
            OfficeXmlTree commentTree = package.GetXml(comments);
            int firstComment = commentTree.FindChild(commentTree.Root, p, "cm");
            int heldComment = commentTree.FindNextSibling(firstComment, p, "cm");
            commentTree.SetAttribute(heldComment, noNamespace, "dt", "2026-08-01T09:30:00+02:00");
            int extension = AddExtension(commentTree);
            int fake = PowerPointDocument.Append(commentTree, extension, p, "cm");
            commentTree.SetAttribute(fake, noNamespace, "idx", "99");

            OfficeXmlTree authorTree = package.GetXml(authors);
            int realAuthor = authorTree.FindChild(authorTree.Root, p, "cmAuthor");
            extension = AddExtension(authorTree, realAuthor);
            fake = PowerPointDocument.Append(authorTree, extension, p, "cmAuthor");
            authorTree.SetAttribute(fake, noNamespace, "id", "0");
            authorTree.SetAttribute(fake, noNamespace, "name", "Extension author");

            using var bytes = new MemoryStream();
            package.Save(bytes);
            bytes.Position = 0;
            using var presentation = PowerPointDocument.Load(bytes);
            PowerPointSlide slideWrapper = presentation.Slides[0];
            IReadOnlyList<PowerPointClassicComment> commentsOnFirstSlide = presentation.GetClassicComments(slideWrapper);
            Assert.Equal(2, commentsOnFirstSlide.Count);
            PowerPointClassicComment held = commentsOnFirstSlide[1];
            Assert.Empty(presentation.GetClassicComments(presentation.Slides[1]));
            Assert.Equal(2U, held.Index);
            Assert.Equal(reviewer.Name, held.Author.Name);
            Assert.Equal(reviewer.Initials, held.Author.Initials);
            Assert.Equal("Review text", held.Text);
            Assert.Equal(created, held.Created);
            Assert.Equal(DateTimeKind.Utc, held.Created!.Value.Kind);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(-12), held.X);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(43), held.Y);
            Assert.False(presentation.IsSdkLoaded);

            presentation.AddClassicComment(presentation.Slides[1], reviewer, "Study note", created: created);
            Assert.True(presentation.IsSdkLoaded);
            held.Text = "Edited text";
            using var saved = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes(), writable: false));
            Assert.Equal(new[] { "First review", "Edited text" }, saved.GetClassicComments(saved.Slides[0]).Select(item => item.Text));
            commentsOnFirstSlide[0].Remove();
            Assert.Equal("Edited text", held.Text);
            held.Remove();
            Assert.Throws<InvalidOperationException>(() => held.Text);

            static int AddExtension(OfficeXmlTree tree, int before = OfficeXmlTree.NoNode) {
                int list = tree.CreateElement(p, "extLst");
                tree.InsertBefore(tree.Root, list, before);
                int extension = PowerPointDocument.Append(tree, list, p, "ext");
                tree.SetAttribute(extension, noNamespace, "uri", "{0A92CB42-753C-4A28-9FA6-E583278CAB90}");
                return extension;
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClassicCommentReadsPreserveSdkValueSemanticsForUnusualMarkup(bool splitText) {
            using var original = PowerPointDocument.Create();
            PowerPointSlide slide = original.AddSlide();
            original.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Original text");
            byte[] input = original.ToBytes();
            using var package = OfficeOpcPackage.Open(new MemoryStream(input, writable: false));
            string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slidePart = package.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            string commentPart = package.GetRelatedPartName(slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments")!;
            using var part = package.ReadPart(commentPart);
            using var stream = new MemoryStream(part.Array, 0, part.Length, writable: false);
            XDocument xml = XDocument.Load(stream);
            XNamespace p = xml.Root!.Name.Namespace;
            XElement comment = xml.Root.Element(p + "cm")!;
            if (splitText) {
                comment.Element(p + "text")!.ReplaceNodes(new XText("First"), new XCData("Middle"), new XText("Last"));
            } else {
                comment.SetAttributeValue("idx", string.Empty);
            }
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(input, writable: false));
            byte[] changedPart = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            edit.ReplacePart(commentPart, changedPart, changedPart.Length);
            using var changed = new MemoryStream();
            edit.Save(changed);
            byte[] source = changed.ToArray();

            using var expected = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            _ = expected.OpenXmlDocument;
            PowerPointClassicComment reference = Assert.Single(expected.GetClassicComments(expected.Slides[0]));
            using var actual = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            PowerPointClassicComment result = Assert.Single(actual.GetClassicComments(actual.Slides[0]));
            if (splitText) {
                Assert.Equal(reference.Text, result.Text);
            } else {
                Exception error = Assert.ThrowsAny<Exception>(() => reference.Index);
                Assert.IsType(error.GetType(), Record.Exception(() => result.Index));
            }
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ClassicCommentReadsPreserveSdkBehaviorForUnusualRelationships(bool interleaved) {
            using var original = PowerPointDocument.Create();
            PowerPointSlide slide = original.AddSlide();
            slide.AddTextBox("Before extraction");
            original.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Review text");
            byte[] input = original.ToBytes();
            using var package = OfficeOpcPackage.Open(new MemoryStream(input, writable: false));
            string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slidePart = package.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(slidePart);
            byte[] source = interleaved
                ? MakeInterleavedCommentRelationships(input, relationshipsPart)
                : MakeNullCommentRelationship(input, package, slidePart, relationshipsPart);

            using var expected = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            _ = expected.OpenXmlDocument;
            string[] reference = expected.GetClassicComments(expected.Slides[0]).Select(comment => comment.Text).ToArray();
            using var actual = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            if (interleaved) {
                Assert.Equal(1, actual.ReplaceAll("Before extraction", "After extraction"));
                Assert.NotEmpty(actual.ExtractMarkdownChunks().ToArray());
                Assert.False(actual.IsSdkLoaded);
            }
            Assert.Equal(reference, actual.GetClassicComments(actual.Slides[0]).Select(comment => comment.Text));
            Assert.True(actual.IsSdkLoaded);
        }

        private static byte[] MakeInterleavedCommentRelationships(byte[] input, string relationshipsPart) {
            using var output = new MemoryStream();
            output.Write(input, 0, input.Length);
            output.Position = 0;
            using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
                string partName = relationshipsPart.TrimStart('/');
                ZipArchiveEntry entry = archive.GetEntry(partName)!;
                using var contents = new MemoryStream();
                using (Stream source = entry.Open()) source.CopyTo(contents);
                entry.Delete();

                byte[] data = contents.ToArray();
                int middle = data.Length / 2;
                using (Stream first = archive.CreateEntry(partName + "/[0].piece").Open())
                    first.Write(data, 0, middle);
                using (Stream last = archive.CreateEntry(partName + "/[1].last.piece").Open())
                    last.Write(data, middle, data.Length - middle);
            }
            return output.ToArray();
        }

        private static byte[] MakeNullCommentRelationship(byte[] input, OfficeOpcPackage package, string slidePart, string relationshipsPart) {
            const string type = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments";
            const string target = "Null";
            string partName = package.GetRelatedPartName(slidePart, type)!;
            using var content = package.ReadPart(partName);
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(input, writable: false));
            edit.CreateBinaryPart(slidePart.Substring(0, slidePart.LastIndexOf('/') + 1) + target,
                package.GetContentType(partName), content.Array.Take(content.Length).ToArray());
            OfficeXmlTree relationships = edit.GetXml(relationshipsPart);
            const int ns = OfficeXmlNamespaces.PackageRelationships;
            int edge = relationships.FindChild(relationships.Root, ns, "Relationship");
            while (edge != OfficeXmlTree.NoNode && relationships.GetAttribute(edge, OfficeXmlNamespaces.None, "Type") != type) {
                edge = relationships.FindNextSibling(edge, ns, "Relationship");
            }
            Assert.NotEqual(OfficeXmlTree.NoNode, edge);
            relationships.SetAttribute(edge, OfficeXmlNamespaces.None, "Target", target);
            using var output = new MemoryStream();
            edit.Save(output);
            return output.ToArray();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClassicCommentFieldEditsPersistAndKeepLookalikesInExtensions(bool omitPosition) {
            const int ns = OfficeXmlNamespaces.PresentationML;
            const int attributes = OfficeXmlNamespaces.None;
            var before = new DateTime(2026, 4, 18, 12, 30, 0, DateTimeKind.Utc);
            var after = before.AddDays(2);
            var x = global::OfficeIMO.OfficeLength.Points(-20);
            var y = global::OfficeIMO.OfficeLength.Points(45);
            using var original = PowerPointDocument.Create();
            PowerPointSlide slide = original.AddSlide();
            original.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Before edit",
                global::OfficeIMO.OfficeLength.Points(12), global::OfficeIMO.OfficeLength.Points(30), before);

            byte[] seed = original.ToBytes();
            using var parts = OfficeOpcPackage.Open(new MemoryStream(seed, writable: false));
            string main = parts.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slidePart = parts.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            string commentPart = parts.GetRelatedPartName(slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments")!;
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(seed, writable: false));
            OfficeXmlTree tree = edit.GetXml(commentPart);
            int commentNode = tree.FindChild(tree.Root, ns, "cm");
            if (omitPosition) tree.Remove(tree.FindChild(commentNode, ns, "pos"));
            int extensions = PowerPointDocument.Append(tree, commentNode, ns, "extLst");
            int extension = PowerPointDocument.Append(tree, extensions, ns, "ext");
            tree.SetAttribute(extension, attributes, "uri", "{D42B21E1-DEF6-4495-A835-4E14F68E42B7}");
            int decoyText = PowerPointDocument.Append(tree, extension, ns, "text");
            int decoyPosition = PowerPointDocument.Append(tree, extension, ns, "pos");
            tree.SetInnerText(decoyText, "Extension text");
            tree.SetAttribute(decoyPosition, attributes, "x", "1234");
            using var modified = new MemoryStream();
            edit.Save(modified);
            byte[] source = modified.ToArray();

            using var reference = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            _ = reference.OpenXmlDocument;
            Apply(Assert.Single(reference.GetClassicComments(reference.Slides[0])));

            using var actual = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            PowerPointClassicComment held = Assert.Single(actual.GetClassicComments(actual.Slides[0]));
            Apply(held);
            Assert.Equal("Revised review\nSecond line.", held.Text);
            Assert.Equal(x, held.X);
            Assert.Equal(omitPosition ? global::OfficeIMO.OfficeLength.Zero : y, held.Y);
            Assert.Equal(after, held.Created);
            byte[] saved = actual.ToBytes();
            Assert.False(actual.IsSdkLoaded);

            XElement expected = ReadClassicCommentPart(reference.ToBytes(), commentPart);
            XElement reopened = ReadClassicCommentPart(saved, commentPart);
            XNamespace p = reopened.Name.Namespace;
            Assert.Equal((string?)expected.Attribute("dt"), (string?)reopened.Attribute("dt"));
            Assert.True(XNode.DeepEquals(expected.Element(p + "pos"), reopened.Element(p + "pos")));
            Assert.Equal(expected.Element(p + "text")!.Value, reopened.Element(p + "text")!.Value);
            XElement retained = reopened.Element(p + "extLst")!.Element(p + "ext")!;
            Assert.Equal("Extension text", retained.Element(p + "text")!.Value);
            Assert.Equal("1234", (string?)retained.Element(p + "pos")!.Attribute("x"));

            void Apply(PowerPointClassicComment item) {
                item.Text = "Revised review\nSecond line.";
                item.X = x;
                if (!omitPosition) item.Y = y;
                item.Created = after;
            }

        }

        [Fact]
        public void ClassicCommentNonUtcAssignmentRetainsTheOriginalDateTimeKindUntilSave() {
            using var original = PowerPointDocument.Create();
            PowerPointSlide slide = original.AddSlide();
            original.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Review");
            using var loaded = PowerPointDocument.Load(new MemoryStream(original.ToBytes(), writable: false));
            PowerPointClassicComment comment = Assert.Single(loaded.GetClassicComments(loaded.Slides[0]));
            var value = new DateTime(2026, 4, 18, 12, 30, 0, DateTimeKind.Unspecified);
            comment.Created = value;
            Assert.Equal(value, comment.Created);
            Assert.Equal(DateTimeKind.Unspecified, comment.Created!.Value.Kind);
        }

        [Theory]
        [InlineData("\r\n")]
        [InlineData("\r")]
        public void ClassicCommentTextWithCarriageReturnsMatchesSdkBeforeAndAfterSave(string separator) {
            (byte[] source, _) = CreateClassicCommentMarkupFixture("plain");
            string replacement = "First" + separator + "Second";
            using var expected = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            _ = expected.OpenXmlDocument;
            PowerPointClassicComment reference = Assert.Single(expected.GetClassicComments(expected.Slides[0]));
            reference.Text = replacement;

            using var actual = PowerPointDocument.Load(new MemoryStream(source, writable: false));
            PowerPointClassicComment result = Assert.Single(actual.GetClassicComments(actual.Slides[0]));
            result.Text = replacement;
            Assert.Equal(reference.Text, result.Text);

            using var expectedSaved = PowerPointDocument.Load(new MemoryStream(expected.ToBytes(), writable: false));
            using var actualSaved = PowerPointDocument.Load(new MemoryStream(actual.ToBytes(), writable: false));
            Assert.Equal(Assert.Single(expectedSaved.GetClassicComments(expectedSaved.Slides[0])).Text,
                Assert.Single(actualSaved.GetClassicComments(actualSaved.Slides[0])).Text);
        }

        [Theory]
        [InlineData("leading CDATA")]
        [InlineData("trailing CDATA")]
        public void ClassicCommentFieldEditsKeepSdkChildOrderWithDirectCData(string markup) {
            (byte[] source, string part) = CreateClassicCommentMarkupFixture(markup);
            XElement reference = SaveWithEdit(openSdk: true);
            XElement actual = SaveWithEdit(openSdk: false);
            Assert.Equal(reference.Nodes().Select(node => node.ToString(SaveOptions.DisableFormatting)),
                actual.Nodes().Select(node => node.ToString(SaveOptions.DisableFormatting)));

            XElement SaveWithEdit(bool openSdk) {
                using var presentation = PowerPointDocument.Load(new MemoryStream(source, writable: false));
                if (openSdk) _ = presentation.OpenXmlDocument;
                PowerPointClassicComment comment = Assert.Single(presentation.GetClassicComments(presentation.Slides[0]));
                if (markup == "leading CDATA") comment.X = global::OfficeIMO.OfficeLength.Points(25);
                else comment.Text = "Revised";
                return ReadClassicCommentPart(presentation.ToBytes(), part);
            }
        }

        [Fact]
        public void ClassicCommentFieldEditsRejectMalformedExtensionReferencesLikeTheSdk() {
            (byte[] source, _) = CreateClassicCommentMarkupFixture("invalid extension entity");
            Exception expected = Assert.ThrowsAny<Exception>(() => Edit(openSdk: true));
            Assert.IsType(expected.GetType(), Record.Exception(() => Edit(openSdk: false)));

            void Edit(bool openSdk) {
                using var presentation = PowerPointDocument.Load(new MemoryStream(source, writable: false));
                if (openSdk) _ = presentation.OpenXmlDocument;
                PowerPointClassicComment comment = Assert.Single(presentation.GetClassicComments(presentation.Slides[0]));
                comment.Created = new DateTime(2026, 4, 20, 9, 0, 0, DateTimeKind.Utc);
            }
        }

        private static (byte[] Bytes, string Part) CreateClassicCommentMarkupFixture(string markup) {
            using var original = PowerPointDocument.Create();
            PowerPointSlide slide = original.AddSlide();
            original.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Before edit");
            byte[] source = original.ToBytes();
            using var package = OfficeOpcPackage.Open(new MemoryStream(source, writable: false));
            string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slidePart = package.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            string partName = package.GetRelatedPartName(slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments")!;
            if (markup == "plain") return (source, partName);
            XElement comment = ReadClassicCommentPart(source, partName);
            XNamespace p = comment.Name.Namespace;
            switch (markup) {
                case "leading CDATA":
                    comment.Element(p + "pos")!.Remove();
                    comment.AddFirst(new XCData("marker"));
                    break;
                case "trailing CDATA":
                    comment.Add(new XCData("marker"));
                    break;
                case "invalid extension entity":
                    XNamespace extension = "urn:officeimo:test:comment-extension";
                    comment.Add(new XElement(p + "extLst", new XElement(p + "ext", new XAttribute("uri", "{CF2A7480-91B6-4A09-8075-22C63EB620D0}"),
                        new XElement(extension + "payload", "InvalidEntityMarker"))));
                    break;
            }

            string xml = comment.Parent!.ToString(SaveOptions.DisableFormatting);
            if (markup == "invalid extension entity") xml = xml.Replace("InvalidEntityMarker", "&undeclared;");
            byte[] replacement = Encoding.UTF8.GetBytes(xml);
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(source, writable: false));
            edit.ReplacePart(partName, replacement, replacement.Length);
            using var output = new MemoryStream();
            edit.Save(output);
            return (output.ToArray(), partName);
        }

        private static XElement ReadClassicCommentPart(byte[] bytes, string partName) {
            using var package = OfficeOpcPackage.Open(new MemoryStream(bytes, writable: false));
            using var part = package.ReadPart(partName);
            using var input = new MemoryStream(part.Array, 0, part.Length, writable: false);
            XElement root = XElement.Load(input);
            return root.Element(root.Name.Namespace + "cm")!;
        }

        [Fact]
        public void ClassicComments_CreateEditReassignRemove_AndRoundTripPpt() {
            byte[] binary;
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                var alice = new PowerPointCommentAuthor("Alice Reviewer", "AR");
                var bob = new PowerPointCommentAuthor("Bob Reviewer", "BR");
                PowerPointClassicComment first = presentation.AddClassicComment(slide,
                    alice, "First review", global::OfficeIMO.OfficeLength.Points(120 / 8D), global::OfficeIMO.OfficeLength.Points(240 / 8D),
                    new DateTime(2026, 8, 1, 9, 0, 0, DateTimeKind.Utc));
                PowerPointClassicComment removed = presentation.AddClassicComment(slide,
                    alice, "Remove me");

                first.Text = "Updated review";
                first.X = global::OfficeIMO.OfficeLength.Points(-25 / 8D);
                first.Y = global::OfficeIMO.OfficeLength.Points(900 / 8D);
                first.SetAuthor(bob);
                removed.Remove();

                PowerPointClassicComment current = Assert.Single(
                    presentation.GetClassicComments(slide));
                Assert.Equal("Bob Reviewer", current.Author.Name);
                Assert.Equal("Updated review", current.Text);
                Assert.Equal(global::OfficeIMO.OfficeLength.Points(-25 / 8D), current.X);
                Assert.Equal(global::OfficeIMO.OfficeLength.Points(900 / 8D), current.Y);
                var preflight = presentation.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite, string.Join(Environment.NewLine,
                    preflight.Findings.Select(finding => finding.Code + ": " + finding.Description)));
                binary = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            }

            using PowerPointDocument roundTrip = PowerPointDocument.Load(
                new MemoryStream(binary, writable: false));
            PowerPointReviewComment projected = Assert.Single(
                roundTrip.InspectReviewComments().Comments);
            Assert.Equal(PowerPointCommentKind.Classic, projected.Kind);
            Assert.Equal("Bob Reviewer", projected.AuthorName);
            Assert.Equal("Updated review", projected.Text);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(-25 / 8D), projected.X);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(900 / 8D), projected.Y);
        }

        [Fact]
        public void ClassicCommentMutation_RecomputesRetainedAuthorLastIndex() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            presentation.AddClassicComment(slide, alice, "First");
            PowerPointClassicComment second = presentation.AddClassicComment(
                slide, alice, "Second");
            PowerPointClassicComment third = presentation.AddClassicComment(
                slide, alice, "Third");

            third.SetAuthor(bob);
            Assert.Equal(2U, GetClassicAuthor(presentation, "Alice").LastIndex!.Value);

            second.Remove();
            Assert.Equal(1U, GetClassicAuthor(presentation, "Alice").LastIndex!.Value);
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.NotEmpty(presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void ClassicCommentMutation_RejectsMissingTimestamp() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointClassicComment comment = presentation.AddClassicComment(
                presentation.AddSlide(),
                new PowerPointCommentAuthor("Reviewer", "R"), "Review");
            DateTime? created = comment.Created;

            Assert.Throws<ArgumentNullException>(() => comment.Created = null);

            Assert.Equal(created, comment.Created);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingCommentedSlideReconcilesClassicAndModernAuthors() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide removed = presentation.AddSlide();
            PowerPointSlide retained = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            PowerPointClassicComment removedClassic = presentation.AddClassicComment(
                removed, alice, "Removed first");
            PowerPointClassicComment surviving = presentation.AddClassicComment(
                retained, alice, "Surviving");
            presentation.AddClassicComment(removed, alice, "Removed last");
            PowerPointModernComment modern = presentation.AddModernComment(
                removed, bob, "Modern removed");
            PowerPointModernCommentReply removedReply = modern.AddReply(
                alice, "Modern reply removed");

            presentation.RemoveSlide(0);

            Assert.Throws<InvalidOperationException>(() => removedClassic.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => modern.Status =
                PowerPointModernCommentStatus.Resolved);
            Assert.Throws<InvalidOperationException>(() => removedReply.Text = "Detached");
            Assert.Equal(surviving.Index,
                GetClassicAuthor(presentation, "Alice").LastIndex!.Value);
            Assert.Empty(presentation.OpenXmlDocument.PresentationPart!.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.NotEmpty(presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void ModernCommentMutation_PreservesAbsentOptionalAuthorIdentity() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddModernComment(slide,
                new PowerPointCommentAuthor("Imported", "I"), "Review");
            P188.Author author = presentation.OpenXmlDocument.PresentationPart!
                .Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList!.Elements<P188.Author>())
                .Single();
            author.UserId = null;
            author.ProviderId = null;
            author.Initials = null;
            string id = author.Id!.Value!;
            PowerPointModernComment comment = Assert.Single(
                presentation.GetModernComments(slide));

            Assert.Equal("I", comment.Author.Initials);

            comment.SetAuthor(comment.Author);

            P188.Author retained = presentation.OpenXmlDocument.PresentationPart!
                .Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList!.Elements<P188.Author>())
                .Single();
            Assert.Equal(id, retained.Id!.Value);
            Assert.Null(retained.Initials);
            Assert.Null(retained.UserId);
            Assert.Null(retained.ProviderId);
        }

        [Fact]
        public void ClassicCommentMutation_PreservesEmptyImportedInitials() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointClassicComment comment = presentation.AddClassicComment(
                presentation.AddSlide(),
                new PowerPointCommentAuthor("Imported", "I"), "Review");
            P.CommentAuthor stored = GetClassicAuthor(presentation, "Imported");
            stored.Initials = string.Empty;
            uint id = stored.Id!.Value;
            uint index = comment.Index;

            PowerPointCommentAuthor imported = comment.Author;
            Assert.Equal("I", imported.Initials);

            comment.SetAuthor(imported);

            P.CommentAuthor retained = Assert.Single(presentation
                .OpenXmlDocument.PresentationPart!.CommentAuthorsPart!
                .CommentAuthorList!.Elements<P.CommentAuthor>());
            Assert.Equal(id, retained.Id!.Value);
            Assert.Equal(string.Empty, retained.Initials!.Value);
            Assert.Equal(index, comment.Index);
        }

        [Fact]
        public void ClassicCommentMutation_ReusesFreeAuthorIdWhenMaximumIsOccupied() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            CommentAuthorsPart authorsPart = presentation.OpenXmlDocument
                .PresentationPart!.AddNewPart<CommentAuthorsPart>();
            authorsPart.CommentAuthorList = new P.CommentAuthorList(
                new P.CommentAuthor {
                    Id = uint.MaxValue,
                    Name = "Maximum",
                    Initials = "M",
                    LastIndex = 0U,
                    ColorIndex = 0U
                });

            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Available", "A"), "Review");

            Assert.Equal(0U, GetClassicAuthor(presentation, "Available").Id!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClassicCommentMutation_ReusesFreeIndexAfterUInt32Maximum() {
            using PowerPointDocument presentation =
            PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            presentation.AddClassicComment(slide, author, "Maximum");
            P.Comment stored = presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().SlideCommentsPart!.CommentList!
                .Elements<P.Comment>().Single();
            stored.Index = uint.MaxValue;
            GetClassicAuthor(presentation, "Reviewer").LastIndex =
                uint.MaxValue;

            presentation.AddClassicComment(slide, author, "Gap");

            Assert.Equal(new[] { 0U, uint.MaxValue }, presentation
                .OpenXmlDocument.PresentationPart!.SlideParts.Single()
                .SlideCommentsPart!.CommentList!.Elements<P.Comment>()
                .Select(comment => comment.Index!.Value)
                .OrderBy(index => index));
            Assert.Equal(uint.MaxValue,
                GetClassicAuthor(presentation, "Reviewer").LastIndex!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClassicCommentMutation_AllocatesWithinBinaryIndexRange() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            presentation.AddClassicComment(slide, author, "Maximum");
            P.Comment maximum = presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().SlideCommentsPart!.CommentList!
                .Elements<P.Comment>().Single();
            maximum.Index = int.MaxValue;
            GetClassicAuthor(presentation, "Reviewer").LastIndex = int.MaxValue;

            presentation.AddClassicComment(slide, author, "Gap");

            Assert.Equal(new[] { 0U, (uint)int.MaxValue }, presentation
                .OpenXmlDocument.PresentationPart!.SlideParts.Single()
                .SlideCommentsPart!.CommentList!.Elements<P.Comment>()
                .Select(comment => comment.Index!.Value)
                .OrderBy(index => index));
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
        }

        [Theory]
        [InlineData("nul\0text")]
        [InlineData("control\u0001text")]
        public void ClassicCommentMutation_RejectsBinaryIncompatibleText(
            string invalidText) {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, author, invalidText));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid");
            Assert.Throws<ArgumentException>(() => comment.Text = invalidText);
            Assert.Equal("Valid", comment.Text);
        }

        [Fact]
        public void ClassicCommentMutation_RejectsTextAboveBinaryLimit() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            string invalidText = new string('x', 32001);

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, author, invalidText));
            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid");
            Assert.Throws<ArgumentException>(() => comment.Text = invalidText);
            Assert.Equal("Valid", comment.Text);
        }

        [Fact]
        public void ClassicCommentMutation_RejectsBinaryIncompatibleAuthors() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCommentAuthor longName = new(new string('n', 53), "N");
            PowerPointCommentAuthor longInitials = new("Name",
                new string('i', 53));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name\0Suffix", "N"));

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, longName, "Review"));
            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, longInitials, "Review"));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, new PowerPointCommentAuthor("Valid", "V"), "Review");
            Assert.Throws<ArgumentException>(() => comment.SetAuthor(longName));
            Assert.Equal("Valid", comment.Author.Name);
        }

        [Fact]
        public void ClassicCommentMutation_RejectsPositionsOutsideBinaryRange() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddClassicComment(slide, author, "Review",
                    global::OfficeIMO.OfficeLength.Points(((long)int.MaxValue + 1L) / 8D), global::OfficeIMO.OfficeLength.Zero));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddClassicComment(slide, author, "Review",
                    global::OfficeIMO.OfficeLength.Zero, global::OfficeIMO.OfficeLength.Points(((long)int.MinValue - 1L) / 8D)));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid", global::OfficeIMO.OfficeLength.Points(10 / 8D), global::OfficeIMO.OfficeLength.Points(20 / 8D));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.X = global::OfficeIMO.OfficeLength.Emus(long.MaxValue));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.Y = global::OfficeIMO.OfficeLength.Emus(long.MinValue));
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(10 / 8D), comment.X);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(20 / 8D), comment.Y);
        }

        [Fact]
        public void ModernCommentMutation_RejectsPositionsOutsideDrawingRange() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, author, "Review",
                    x: global::OfficeIMO.OfficeLength.Emus(long.MaxValue)));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, author, "Review",
                    y: global::OfficeIMO.OfficeLength.Emus(long.MinValue)));
            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.Empty(presentation.GetModernComments(slide));

            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Valid", x: global::OfficeIMO.OfficeLength.Emus(10L), y: global::OfficeIMO.OfficeLength.Emus(20L));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.X = global::OfficeIMO.OfficeLength.Emus(long.MaxValue));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.Y = global::OfficeIMO.OfficeLength.Emus(long.MinValue));
            Assert.Equal(10L, comment.X.ToEmus());
            Assert.Equal(20L, comment.Y.ToEmus());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentCoordinateGettersDoNotMaterializeMissingPosition() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Positionless");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            nativeComment.GetFirstChild<P188.Point2DType>()!.Remove();

            Assert.Equal(0L, comment.X.ToEmus());
            Assert.Equal(0L, comment.Y.ToEmus());
            Assert.Null(nativeComment.GetFirstChild<P188.Point2DType>());

            comment.X = global::OfficeIMO.OfficeLength.Emus(12L);
            Assert.Equal(12L, comment.X.ToEmus());
            Assert.Equal(0L, comment.Y.ToEmus());
            Assert.NotNull(nativeComment.GetFirstChild<P188.Point2DType>());
        }

        [Fact]
        public void ModernReplyMutationRejectsDetachedReplyList() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Review");
            PowerPointModernCommentReply reply = comment.AddReply(author,
                "Reply");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            P188.CommentReplyList replyList = Assert.IsType<P188.CommentReplyList>(
                nativeComment.GetFirstChild<P188.CommentReplyList>());

            replyList.Remove();

            Assert.NotNull(replyList.GetFirstChild<P188.CommentReply>()?.Parent);
            Assert.Throws<InvalidOperationException>(() => reply.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => reply.Status =
                PowerPointModernCommentStatus.Resolved);
            Assert.Throws<InvalidOperationException>(() =>
                reply.SetAuthor(new PowerPointCommentAuthor("Other", "O")));
        }

        [Fact]
        public void ModernComments_CreateEditReplyReassignRemove_AndRoundTripPptx() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                var alice = new PowerPointCommentAuthor("Alice Reviewer", "AR",
                    "alice@example.test", "OfficeIMO");
                var bob = new PowerPointCommentAuthor("Bob Reviewer", "BR",
                    "bob@example.test", "OfficeIMO");
                PowerPointModernComment comment = presentation.AddModernComment(slide,
                    alice, "Modern review", PowerPointModernCommentStatus.Active,
                    global::OfficeIMO.OfficeLength.Emus(123), global::OfficeIMO.OfficeLength.Emus(456), new DateTime(2026, 8, 1, 10, 0, 0, DateTimeKind.Utc));
                PowerPointModernCommentReply reply = comment.AddReply(bob, "Initial reply");
                PowerPointModernCommentReply removed = comment.AddReply(alice, "Remove reply");

                comment.Text = "Updated modern review";
                comment.Status = PowerPointModernCommentStatus.Resolved;
                comment.X = global::OfficeIMO.OfficeLength.Emus(321);
                comment.SetAuthor(bob);
                reply.Text = "Updated reply";
                reply.Status = PowerPointModernCommentStatus.Closed;
                removed.Remove();

                Assert.Equal("Bob Reviewer", comment.Author.Name);
                Assert.Equal("Bob Reviewer", reply.Author.Name);
                Assert.Single(comment.Replies);
                Assert.Contains(presentation.AnalyzeLegacyPptWrite().Findings,
                    finding => finding.Code == "ppt-write-modern-comments");
                PowerPointFeatureFinding finding = presentation.InspectFeatures().Features
                    .Single(item => item.Name == "Comments");
                Assert.Equal(OfficeFeatureSupportLevel.Editable, finding.SupportLevel);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointModernComment current = Assert.Single(
                reopened.GetModernComments(reopened.Slides[0]));
            Assert.Equal("Updated modern review", current.Text);
            Assert.Equal(PowerPointModernCommentStatus.Resolved, current.Status);
            Assert.Equal(321, current.X.ToEmus());
            PowerPointModernCommentReply currentReply = Assert.Single(current.Replies);
            Assert.Equal("Updated reply", currentReply.Text);
            Assert.Equal(PowerPointModernCommentStatus.Closed, currentReply.Status);
            PowerPointReviewReport report = reopened.InspectReviewComments();
            Assert.Equal(2, report.ModernCount);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));

            current.Remove();
            Assert.Empty(reopened.GetModernComments(reopened.Slides[0]));
        }

        [Fact]
        public void ModernCommentMutation_RejectsXmlInvalidText() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R",
                "reviewer@example.test", "OfficeIMO");

            Assert.Throws<ArgumentException>(() =>
                presentation.AddModernComment(slide, author, "Bad\0text"));
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Valid");
            PowerPointModernCommentReply reply = comment.AddReply(author,
                "Valid reply");
            Assert.Throws<ArgumentNullException>(() =>
                comment.Created = null);
            Assert.Throws<ArgumentNullException>(() =>
                reply.Created = null);
            Assert.Throws<ArgumentException>(() => comment.Text = "Bad\u0001text");
            Assert.Throws<ArgumentException>(() => reply.Text = "Bad\u000Btext");
            Assert.Throws<ArgumentException>(() =>
                comment.AddReply(author, "Bad\u000Ctext"));
            Assert.Equal("Valid", comment.Text);
            Assert.Equal("Valid reply", reply.Text);
            Assert.NotNull(comment.Created);
            Assert.NotNull(reply.Created);
            Assert.Single(comment.Replies);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentMutation_RejectsOversizedAndNewlineAmplifiedTextBeforeAllocation() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Alice Reviewer", "AR");
            string tooLong = new string('x', 32001);
            string tooManyParagraphs = string.Join("\n",
                Enumerable.Repeat("x", 1025));

            Assert.Throws<ArgumentException>(() =>
                presentation.AddModernComment(slide, author, tooLong));
            Assert.Throws<ArgumentException>(() =>
                presentation.AddModernComment(slide, author, tooManyParagraphs));

            Assert.Empty(presentation.GetModernComments(slide));
        }

        [Fact]
        public void ModernCommentMutation_RejectsInvalidStatusBeforeAuthorCreation() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A",
                "alice@example.test", "OfficeIMO");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, alice, "Review",
                    (PowerPointModernCommentStatus)999));

            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.Empty(presentation.GetModernComments(slide));
            Assert.DoesNotContain(presentation.AnalyzeLegacyPptWrite().Findings,
                finding => finding.Code == "ppt-write-modern-comments");

            PowerPointModernComment comment = presentation.AddModernComment(
                slide, alice, "Valid review");
            var bob = new PowerPointCommentAuthor("Bob", "B",
                "bob@example.test", "OfficeIMO");
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.AddReply(bob, "Reply",
                    (PowerPointModernCommentStatus)999));
            Assert.Equal(new[] { "Alice" }, GetModernAuthorNames(presentation));
            Assert.Empty(comment.Replies);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CommentAuthor_RejectsXmlInvalidIdentityFields() {
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Bad\u0001name"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "B\u0001"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "N", "bad\u0001user"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "N", "user",
                    "bad\u0001provider"));
        }

        [Fact]
        public void CommentMutation_RejectsForeignSlidesAndEmptyText() {
            using PowerPointDocument first = PowerPointDocument.Create();
            using PowerPointDocument second = PowerPointDocument.Create();
            PowerPointSlide foreign = second.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer");

            Assert.Throws<ArgumentException>(() =>
                first.AddClassicComment(foreign, author, "Review"));
            Assert.Throws<ArgumentException>(() =>
                first.AddModernComment(foreign, author, "Review"));
            Assert.Throws<ArgumentException>(() =>
                first.AddClassicComment(first.AddSlide(), author, " "));
        }

        [Fact]
        public void CommentAuthor_GeneratesUnicodeScalarInitialsThatRoundTrip() {
            const string supplementaryCjk = "\U00020000";
            var emojiAuthor = new PowerPointCommentAuthor("😀 Reviewer");
            var cjkAuthor = new PowerPointCommentAuthor(supplementaryCjk + " Reviewer");

            Assert.Equal("😀R", emojiAuthor.Initials);
            Assert.Equal(supplementaryCjk + "R", cjkAuthor.Initials);

            using var stream = new MemoryStream();
            using (PowerPointDocument presentation =
                   PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                presentation.AddClassicComment(slide, emojiAuthor, "Emoji author");
                presentation.AddModernComment(slide, cjkAuthor, "CJK author");
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            Assert.Equal("😀R", Assert.Single(reopened.GetClassicComments(
                reopened.Slides[0])).Author.Initials);
            Assert.Equal(supplementaryCjk + "R", Assert.Single(
                reopened.GetModernComments(reopened.Slides[0])).Author.Initials);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ModernCommentMutation_RemovesUnusedAuthorMetadata() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, alice, "Review");
            PowerPointModernCommentReply reply = comment.AddReply(alice, "Reply");

            comment.SetAuthor(bob);
            Assert.Equal(new[] { "Alice", "Bob" }, GetModernAuthorNames(presentation));
            reply.SetAuthor(bob);
            Assert.Equal(new[] { "Bob" }, GetModernAuthorNames(presentation));

            comment.Remove();
            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.DoesNotContain(presentation.OpenXmlDocument.PresentationPart!.Parts,
                pair => pair.OpenXmlPart is PowerPointAuthorsPart);
            Assert.Empty(presentation.GetModernComments(slide));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentMutation_PreservesProducerMetadataWhenRemovingLastComment() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Review");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            part.CommentList!.Append(new OpenXmlUnknownElement(
                "vendor", "producerData",
                "urn:officeimo:test:comments"));
            string original = part.CommentList.OuterXml;

            Assert.Throws<NotSupportedException>(() => comment.Remove());

            Assert.Equal(original, part.CommentList.OuterXml);
            Assert.Contains(slide.SlidePart.Parts,
                pair => ReferenceEquals(pair.OpenXmlPart, part));
            Assert.Single(presentation.GetModernComments(slide));
            Assert.Equal(new[] { "Reviewer" },
                GetModernAuthorNames(presentation));
        }

        [Fact]
        public void ModernCommentText_RejectsRichBodyReplacementWithoutMutation() {
            var comment = new P188.Comment();
            var body = new P188.TextBodyType(new A.BodyProperties(),
                new A.ListStyle());
            body.Append(
                new A.Paragraph(new A.Run(new A.Text("First")),
                    new A.Break(), new A.Run(new A.Text("Second"))),
                new A.Paragraph(new A.Run(new A.Text("Third"))));
            comment.Append(body);

            Assert.Equal("First\nSecond\nThird",
                PowerPointDocument.GetModernCommentText(comment));
            string original = body.OuterXml;

            Assert.Throws<NotSupportedException>(() =>
                PowerPointDocument.SetModernCommentText(comment,
                    "Alpha\r\n\r\nBeta"));

            Assert.Equal(original, body.OuterXml);
            Assert.Equal("First\nSecond\nThird",
                PowerPointDocument.GetModernCommentText(comment));
        }

        [Fact]
        public void RichModernCommentBodyIsPreservedAndReportedAsNonEditable() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Plain");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            P188.TextBodyType body = nativeComment
                .GetFirstChild<P188.TextBodyType>()!;
            body.RemoveAllChildren<A.Paragraph>();
            body.Append(new A.Paragraph(
                new A.Run(new A.RunProperties { Bold = true },
                    new A.Text("Rich")),
                new A.Run(new A.RunProperties { Italic = true },
                    new A.Text(" body"))));
            string original = body.OuterXml;

            PowerPointFeatureFinding finding = Assert.Single(presentation
                .InspectFeatures().FindFeatures("Comments"));
            Assert.Equal(OfficeFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Throws<NotSupportedException>(() =>
                comment.Text = "Replacement");
            Assert.Equal(original, body.OuterXml);
        }

        private static string[] GetModernAuthorNames(
            PowerPointDocument presentation) => presentation.OpenXmlDocument
                .PresentationPart!.Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList?.Elements<P188.Author>()
                    ?? Enumerable.Empty<P188.Author>())
                .Select(author => author.Name?.Value ?? string.Empty)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

        private static P.CommentAuthor GetClassicAuthor(
            PowerPointDocument presentation, string name) => presentation
                .OpenXmlDocument.PresentationPart!.CommentAuthorsPart!
                .CommentAuthorList!.Elements<P.CommentAuthor>()
                .Single(author => author.Name?.Value == name);
    }
}
