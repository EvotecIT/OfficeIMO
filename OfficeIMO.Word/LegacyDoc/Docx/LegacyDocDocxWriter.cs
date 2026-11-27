using System.Threading;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    /// <summary>
    /// Writes the WordprocessingML package for an imported legacy DOC model through OfficeIMO's own Open XML layer.
    /// The package starts from the engine's default document package, so it carries the same styles, theme and settings
    /// as a new document, and the engine loads it like any other DOCX.
    /// </summary>
    internal sealed partial class LegacyDocDocxWriter {
        internal const string MainPartName = "/word/document.xml";
        private const string StylesPartName = "/word/styles.xml";
        private const string SettingsPartName = "/word/settings.xml";
        private const string RelationshipTypePrefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;

        private readonly OfficeOpcEditablePackage _package;
        private readonly LegacyDocDocument _source;
        private readonly LegacyDocStyleSheet _styleSheet;
        private readonly CancellationToken _cancellationToken;
        private readonly LegacyDocDocxPart _main;
        private readonly OfficeXmlTree _styles;
        private readonly int _body;

        private LegacyDocDocxWriter(OfficeOpcEditablePackage package, LegacyDocDocument source, CancellationToken cancellationToken) {
            _package = package;
            _source = source;
            _styleSheet = source.StyleSheet;
            _cancellationToken = cancellationToken;
            _main = new LegacyDocDocxPart(MainPartName, package.GetXml(MainPartName));
            _styles = package.GetXml(StylesPartName);
            _body = _main.Tree.FindChild(_main.Tree.Root, W, "body");
            if (_body == OfficeXmlTree.NoNode) throw new InvalidDataException("The default Word package has no document body.");
        }

        /// <summary>
        /// Writes the DOCX package for <paramref name="source"/>, starting from <paramref name="defaultPackage"/>, the
        /// engine's default package for a plain document.
        /// </summary>
        internal static byte[] Write(LegacyDocDocument source, byte[] defaultPackage, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            using var input = new MemoryStream(defaultPackage, writable: false);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            var writer = new LegacyDocDocxWriter(package, source, cancellationToken);
            writer.WriteDocument();
            using var output = new MemoryStream(defaultPackage.Length * 2);
            package.Save(output);
            return output.ToArray();
        }

        private void WriteDocument() {
            AddDocumentProperties();
            ApplyParagraphDefaults();
            AddNumberingDefinitions(_source.Numbering);
            AddParagraphStyleDefinitions();
            ResetStyleToggleIndex();

            var sections = new List<(int SectionProperties, LegacyDocSectionFormat Format)> {
                (FinalSectionProperties(), _source.SectionFormat)
            };
            LegacyDocNoteProjection notes = LegacyDocNoteProjection.Create(_source.Footnotes, _source.Endnotes, _source.Comments, _styleSheet, _cancellationToken);

            if (_source.BodyBlocks.Count == 0) {
                AppendBodyBlock(CreateBodyParagraph());
            } else {
                int pendingBoundaryParagraph = OfficeXmlTree.NoNode;
                foreach (LegacyDocBodyBlock block in _source.BodyBlocks) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (block is LegacyDocParagraphBlock paragraphBlock) {
                        int paragraph = AddBodyParagraph(paragraphBlock, notes);
                        pendingBoundaryParagraph = paragraphBlock.EndsWithSectionMark ? paragraph : OfficeXmlTree.NoNode;
                    } else if (block is LegacyDocSectionBreakBlock sectionBreakBlock) {
                        AddSection(sections, sectionBreakBlock.Format, pendingBoundaryParagraph);
                        pendingBoundaryParagraph = OfficeXmlTree.NoNode;
                    } else if (block is LegacyDocTableBlock tableBlock) {
                        pendingBoundaryParagraph = OfficeXmlTree.NoNode;
                        int table = AddTable(_main, tableBlock, notes, projectNestedTables: true);
                        if (table != OfficeXmlTree.NoNode) {
                            AppendBodyBlock(table);
                            AddTableBlockBookmarks(_main.Tree, table, tableBlock);
                        }
                    }
                }
            }

            foreach ((int sectionProperties, LegacyDocSectionFormat format) in sections) {
                _cancellationToken.ThrowIfCancellationRequested();
                ApplySectionFormatting(sectionProperties, format);
            }

            List<int> sectionList = sections.Select(section => section.SectionProperties).ToList();
            ApplyDocumentOptions(sectionList);
            AddHeaderFooterStories(sectionList);
            AddTextBoxStories(notes);
            _cancellationToken.ThrowIfCancellationRequested();
        }

        // ---- Body ---------------------------------------------------------------------------------------------

        /// <summary>The body-level section properties, which describe the last section.</summary>
        private int FinalSectionProperties() {
            OfficeXmlTree tree = _main.Tree;
            int sectionProperties = tree.FindChild(_body, W, "sectPr");
            if (sectionProperties == OfficeXmlTree.NoNode) sectionProperties = Append(tree, _body, Element(tree, "sectPr"));
            return sectionProperties;
        }

        /// <summary>Adds a body block before the final section properties.</summary>
        private void AppendBodyBlock(int block) {
            OfficeXmlTree tree = _main.Tree;
            int sectionProperties = tree.FindChild(_body, W, "sectPr");
            tree.InsertBefore(_body, block, sectionProperties);
        }

        private int CreateBodyParagraph() {
            OfficeXmlTree tree = _main.Tree;
            int paragraph = Element(tree, "p");
            Append(tree, paragraph, Element(tree, "pPr"));
            return paragraph;
        }

        private int AddBodyParagraph(LegacyDocParagraphBlock paragraphBlock, LegacyDocNoteProjection notes) {
            int paragraph = CreateBodyParagraph();
            AppendBodyBlock(paragraph);
            ApplyParagraphFormatting(_main, paragraph, paragraphBlock.Format);
            var bookmarks = LegacyDocBookmarkProjection.Create(paragraphBlock.Bookmarks, paragraphBlock.StartCharacter, paragraphBlock.EndCharacter);
            var target = new RunTarget(_main, paragraph, paragraph);
            if (paragraphBlock.Runs.Count > 0) AddRuns(target, paragraphBlock.Runs, 0, notes, bookmarks);
            bookmarks.EmitRemaining(_main.Tree, paragraph);
            return paragraph;
        }

        // ---- Package -------------------------------------------------------------------------------------------

        /// <summary>The XML part related to the main part by <paramref name="relationshipType"/>, created when missing.</summary>
        private LegacyDocDocxPart GetOrCreateMainRelatedPart(string relationshipType, string partName, string contentType, string rootName) {
            foreach (LegacyDocRelationship relationship in Relationships(MainPartName)) {
                if (relationship.Type == relationshipType && !relationship.External) {
                    string existing = OfficeOpcPartNames.ResolveRelationshipTarget(OfficeOpcPartNames.GetRelationshipsPartName(MainPartName), relationship.Target)
                        ?? throw new InvalidDataException("The default Word package has an unreadable relationship target.");
                    return new LegacyDocDocxPart(existing, _package.GetXml(existing));
                }
            }

            OfficeXmlTree tree = _package.CreateXmlPart(partName, contentType, W, rootName);
            DeclareWordNamespaces(tree);
            _package.AddRelationship(MainPartName, relationshipType, RelativeTarget(partName));
            return new LegacyDocDocxPart(partName, tree);
        }

        private static string RelativeTarget(string partName) => partName.StartsWith("/word/", StringComparison.Ordinal) ? partName.Substring(6) : partName;

        private readonly struct LegacyDocRelationship {
            internal LegacyDocRelationship(string id, string type, string target, bool external) {
                Id = id;
                Type = type;
                Target = target;
                External = external;
            }

            internal string Id { get; }
            internal string Type { get; }
            internal string Target { get; }
            internal bool External { get; }
        }

        private IEnumerable<LegacyDocRelationship> Relationships(string sourcePartName) {
            OfficeXmlTree relationships = _package.GetRelationships(sourcePartName);
            var result = new List<LegacyDocRelationship>();
            for (int child = relationships.FirstChildOf(relationships.Root); child != OfficeXmlTree.NoNode; child = relationships.NextSiblingOf(child)) {
                if (relationships.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                string? id = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Id");
                string? type = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Type");
                string? target = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Target");
                string? mode = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "TargetMode");
                if (id == null || type == null || target == null) continue;
                result.Add(new LegacyDocRelationship(id, type, target, string.Equals(mode, "External", StringComparison.Ordinal)));
            }
            return result;
        }

        /// <summary>Declares the namespaces Word parts use on the root of a new part.</summary>
        private static void DeclareWordNamespaces(OfficeXmlTree tree) {
            tree.DeclareNamespace(tree.Root, "w", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.WordprocessingML)!);
            tree.DeclareNamespace(tree.Root, "r", OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.Relationships)!);
        }

        private string NextPartName(string prefix, string extension) {
            for (int index = 1; ; index++) {
                string name = prefix + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + extension;
                if (!_package.PartExists(name)) return name;
            }
        }
    }

    /// <summary>An XML part of the package being written: its name, for relationships, and its tree.</summary>
    internal sealed class LegacyDocDocxPart {
        internal LegacyDocDocxPart(string name, OfficeXmlTree tree) {
            Name = name;
            Tree = tree;
        }

        internal string Name { get; }

        internal OfficeXmlTree Tree { get; }
    }
}
