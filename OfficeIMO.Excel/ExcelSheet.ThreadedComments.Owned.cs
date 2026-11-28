using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string ThreadedNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
        private const string ThreadedRelationship = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
        private const string PersonRelationship = "http://schemas.microsoft.com/office/2017/10/relationships/person";
        private const string ThreadedContentType = "application/vnd.ms-excel.threadedcomments+xml";
        private const string PersonContentType = "application/vnd.ms-excel.person+xml";

        private static IReadOnlyList<ExcelThreadedCommentInfo> ReadOwnedThreadedComments(OfficeOpcPackage package, ExcelDeferredSheet sheet) {
            string? workbook = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
            return ReadThreadedComments(ReadThreadedRoots(package, sheet.PartName, ThreadedRelationship),
                ReadThreadedPeople(workbook == null ? Enumerable.Empty<OfficeXmlNode>() : ReadThreadedRoots(package, workbook, PersonRelationship)), sheet.Name);
        }

        internal static Dictionary<string, string> ReadThreadedPeople(ExcelPartRef workbook) =>
            ReadThreadedPeople(ThreadedPartRoots(workbook, PersonRelationship));

        internal static IReadOnlyList<ExcelThreadedCommentInfo> ReadThreadedComments(ExcelPartRef worksheet, IReadOnlyDictionary<string, string> people, string sheetName) =>
            ReadThreadedComments(ThreadedPartRoots(worksheet, ThreadedRelationship), people, sheetName);

        private static IEnumerable<OfficeXmlNode> ThreadedPartRoots(ExcelPartRef owner, string relationshipType) =>
            owner.Parts.Where(part => part.Part.RelationshipType == relationshipType).Select(part => part.Part.Root);

        private static IEnumerable<OfficeXmlNode> ReadThreadedRoots(OfficeOpcPackage package, string owner, string relationshipType) {
            foreach (OfficeOpcRelationship relation in package.GetRelationships(owner)) {
                if (relation.Type != relationshipType || relation.TargetPartName is not string target || !package.PartExists(target)) continue;
                using OfficePooledBytes bytes = package.ReadPart(target);
                OfficeXmlTree tree = OfficeXmlTree.Parse(bytes.Array, bytes.Length, package.Namespaces);
                yield return new OfficeXmlNode(tree, tree.Root);
            }
        }

        private static Dictionary<string, string> ReadThreadedPeople(IEnumerable<OfficeXmlNode> roots) {
            var people = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeXmlNode root in roots) {
                foreach (OfficeXmlNode person in ThreadedChildren(root, "person")) {
                    string? id = PresentThreadedValue(person.Attribute("id"));
                    string? name = PresentThreadedValue(person.Attribute("displayName"));
                    if (id != null && name != null) people[id] = name;
                }
            }
            return people;
        }

        private static IReadOnlyList<ExcelThreadedCommentInfo> ReadThreadedComments(IEnumerable<OfficeXmlNode> roots,
            IReadOnlyDictionary<string, string> people, string sheetName) {
            var comments = new Dictionary<string, List<ExcelThreadedCommentInfo>>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeXmlNode root in roots) {
                foreach (OfficeXmlNode comment in ThreadedChildren(root, "threadedComment")) {
                    string? reference = PresentThreadedValue(comment.Attribute("ref"));
                    if (reference == null) continue;
                    string? personId = PresentThreadedValue(comment.Attribute("personId"));
                    string? author = personId != null && people.TryGetValue(personId, out string? name) ? name : null;
                    string? date = comment.Attribute("dT");
                    if (!comments.TryGetValue(reference, out List<ExcelThreadedCommentInfo>? bucket)) {
                        bucket = new List<ExcelThreadedCommentInfo>();
                        comments.Add(reference, bucket);
                    }
                    bucket.Add(new ExcelThreadedCommentInfo {
                        SheetName = sheetName,
                        CellReference = reference,
                        Id = PresentThreadedValue(comment.Attribute("id")),
                        ParentId = PresentThreadedValue(comment.Attribute("parentId")),
                        PersonId = personId,
                        Author = author,
                        Text = comment.Child(ThreadedNamespace, "text").InnerText.Replace("\r\n", "\n").Replace('\r', '\n'),
                        Date = date == null ? (DateTime?)null : XmlConvert.ToDateTime(date, XmlDateTimeSerializationMode.RoundtripKind),
                        Done = comment.BooleanAttribute("done") == true
                    });
                }
            }
            return comments.Values.SelectMany(bucket => bucket).ToArray();
        }

        private static bool AddOwnedThreadedComments(OfficeOpcEditablePackage package, ExcelDeferredSheet sheet,
            List<(string Address, string Text, string Author, string Id, string? ParentId, DateTime Date, bool Done)> prepared,
            List<ExcelThreadedCommentResult> results, string parameterName) {
            string workbook = GetOwnedThreadedWorkbook(package);
            var identities = new Dictionary<string, (string Worksheet, string? Address, bool IsReply)>(StringComparer.OrdinalIgnoreCase);
            const string worksheetRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "worksheet";
            foreach (var worksheet in OwnedThreadedParts(package, workbook, worksheetRelationship)) {
                foreach (var part in OwnedThreadedParts(package, worksheet.Target, ThreadedRelationship)) {
                    foreach (OfficeXmlNode comment in ThreadedChildren(OwnedThreadedRoot(package, part.Target), "threadedComment")) {
                        string? id = PresentThreadedValue(comment.Attribute("id"));
                        if (id != null && !identities.ContainsKey(id)) {
                            identities.Add(id, (worksheet.Target, comment.Attribute("ref"), PresentThreadedValue(comment.Attribute("parentId")) != null));
                        }
                    }
                }
            }
            foreach (var item in prepared) {
                if (identities.ContainsKey(item.Id)) throw new InvalidOperationException($"A threaded comment with id '{item.Id}' already exists in the workbook.");
                if (item.ParentId != null) {
                    if (!identities.TryGetValue(item.ParentId, out var parent))
                        throw new ArgumentException($"Parent threaded comment '{item.ParentId}' does not exist.", parameterName);
                    if (!string.Equals(parent.Worksheet, sheet.PartName, StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(parent.Address, item.Address, StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException("A threaded reply must use the same worksheet and cell as its parent comment.", parameterName);
                    if (parent.IsReply) throw new ArgumentException("A threaded reply must reference the root comment rather than another reply.", parameterName);
                }
                identities.Add(item.Id, (sheet.PartName, item.Address, item.ParentId != null));
            }

            var people = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeXmlNode person in OwnedThreadedPeople(package, workbook)) {
                if (person.Attribute("displayName") is string name && PresentThreadedValue(person.Attribute("id")) is string id
                    && !people.ContainsKey(name)) people.Add(name, id);
            }
            OfficeXmlNode comments = EnsureOwnedThreadedRoot(package, sheet.PartName, ThreadedRelationship, ThreadedContentType,
                "/xl/threadedcomments/threadedcomment", "ThreadedComments");
            foreach (var item in prepared) {
                string person = EnsureOwnedThreadedPerson(package, workbook, item.Author, people);
                OfficeXmlNode comment = AppendThreadedElement(comments, "threadedComment");
                comment.SetAttribute("ref", item.Address);
                comment.SetAttribute("dT", XmlConvert.ToString(item.Date, XmlDateTimeSerializationMode.Utc));
                comment.SetAttribute("personId", person);
                comment.SetAttribute("id", item.Id);
                comment.SetAttribute("parentId", item.ParentId);
                if (item.Done) comment.SetBooleanAttribute("done", true);
                AppendThreadedElement(comment, "text").SetInnerText(item.Text);
                results.Add(new ExcelThreadedCommentResult(sheet.Name, item.Address, item.Id, person, item.Author, item.ParentId != null, item.Done));
            }
            return true;
        }

        private static ExcelThreadedCommentResult UpdateOwnedThreadedComment(OfficeOpcEditablePackage package, ExcelDeferredSheet sheet,
            string id, ExcelThreadedCommentUpdateOptions options) {
            string workbook = GetOwnedThreadedWorkbook(package);
            OfficeXmlNode comment = OwnedThreadedParts(package, sheet.PartName, ThreadedRelationship)
                .SelectMany(part => ThreadedChildren(OwnedThreadedRoot(package, part.Target), "threadedComment"))
                .FirstOrDefault(item => string.Equals(item.Attribute("id"), id, StringComparison.OrdinalIgnoreCase));
            if (comment.IsNone) throw new KeyNotFoundException($"Threaded comment '{id}' was not found on worksheet '{sheet.Name}'.");

            string personId = comment.Attribute("personId") ?? string.Empty;
            string author;
            if (options.Author != null) {
                author = options.Author.Trim();
                personId = EnsureOwnedThreadedPerson(package, workbook, author);
                comment.SetAttribute("personId", personId);
            } else {
                author = OwnedThreadedPeople(package, workbook)
                    .FirstOrDefault(candidate => string.Equals(candidate.Attribute("id"), personId, StringComparison.OrdinalIgnoreCase))
                    .Attribute("displayName") ?? string.Empty;
            }
            if (options.Text != null) {
                foreach (OfficeXmlNode previous in ThreadedChildren(comment, "text").ToList()) previous.Remove();
                AppendThreadedElement(comment, "text").SetInnerText(options.Text);
            }
            if (options.Date.HasValue) {
                comment.SetAttribute("dT", XmlConvert.ToString(NormalizeThreadedTimestamp(options.Date), XmlDateTimeSerializationMode.Utc));
            }
            if (options.Done.HasValue) comment.SetBooleanAttribute("done", options.Done);
            return new ExcelThreadedCommentResult(sheet.Name, comment.Attribute("ref") ?? string.Empty, id, personId, author,
                PresentThreadedValue(comment.Attribute("parentId")) != null, comment.BooleanAttribute("done") == true);
        }

        private static int RemoveOwnedThreadedComments(OfficeOpcEditablePackage package, ExcelDeferredSheet sheet, string? id,
            string? reference, bool removeReplies) {
            var parts = OwnedThreadedParts(package, sheet.PartName, ThreadedRelationship).ToArray();
            var comments = parts.SelectMany(part => ThreadedChildren(OwnedThreadedRoot(package, part.Target), "threadedComment")).ToArray();
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (id != null) {
                if (!comments.Any(comment => string.Equals(comment.Attribute("id"), id, StringComparison.OrdinalIgnoreCase))) return 0;
                ids.Add(id);
                var children = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (OfficeXmlNode comment in comments) {
                    string? parent = PresentThreadedValue(comment.Attribute("parentId"));
                    string? reply = PresentThreadedValue(comment.Attribute("id"));
                    if (parent == null || reply == null) continue;
                    if (!children.TryGetValue(parent, out List<string>? replies)) {
                        replies = new List<string>();
                        children.Add(parent, replies);
                    }
                    replies.Add(reply);
                }
                var pending = new Queue<string>();
                pending.Enqueue(id);
                while (pending.Count != 0) {
                    if (!children.TryGetValue(pending.Dequeue(), out List<string>? replies)) continue;
                    foreach (string reply in replies) {
                        if (ids.Add(reply)) pending.Enqueue(reply);
                    }
                }
                if (!removeReplies && ids.Count > 1) {
                    throw new InvalidOperationException("The threaded comment has replies. Remove the complete thread or remove the replies first.");
                }
            }

            int removed = 0;
            foreach (var part in parts) {
                OfficeXmlNode root = OwnedThreadedRoot(package, part.Target);
                foreach (OfficeXmlNode comment in ThreadedChildren(root, "threadedComment").ToList()) {
                    bool matches = id != null
                        ? comment.Attribute("id") is string commentId && ids.Contains(commentId)
                        : string.Equals(comment.Attribute("ref"), reference, StringComparison.OrdinalIgnoreCase);
                    if (!matches) continue;
                    comment.Remove();
                    removed++;
                }
                if (ThreadedChildren(root, "threadedComment").Any()) continue;
                OfficeXmlNode relationships = part.Relationship.Parent;
                part.Relationship.Remove();
                if (relationships.FirstChild.IsNone) package.DeletePart(OfficeOpcPartNames.GetRelationshipsPartName(sheet.PartName));
                if (!HasOwnedPartReference(package, part.Target)) package.DeletePart(part.Target);
            }
            return removed;
        }

        private static string? PresentThreadedValue(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

        private static string GetOwnedThreadedWorkbook(OfficeOpcEditablePackage package) =>
            OwnedThreadedParts(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType).FirstOrDefault().Target
                ?? throw new InvalidOperationException("Workbook part is missing.");

        private static IEnumerable<OfficeXmlNode> ThreadedChildren(OfficeXmlNode root, string name) {
            foreach (OfficeXmlNode child in root.Children()) {
                if (child.Is(ThreadedNamespace, name)) yield return child;
            }
        }

        private static IEnumerable<(string Target, OfficeXmlNode Relationship)> OwnedThreadedParts(
            OfficeOpcEditablePackage package, string? owner, string type) {
            string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (!package.PartExists(relationships)) yield break;
            OfficeXmlNode root = OwnedThreadedRoot(package, relationships);
            foreach (OfficeXmlNode relationship in root.Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (relationship.Attribute("Type") != type || relationship.Attribute("TargetMode") == "External") continue;
                string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationships, relationship.Attribute("Target"));
                if (target != null && package.PartExists(target)) yield return (target, relationship);
            }
        }

        private static OfficeXmlNode OwnedThreadedRoot(OfficeOpcEditablePackage package, string name) {
            OfficeXmlTree tree = package.GetXml(name);
            return new OfficeXmlNode(tree, tree.Root);
        }

        private static OfficeXmlNode AppendThreadedElement(OfficeXmlNode parent, string name) =>
            parent.Append(parent.CreateElement(ThreadedNamespace, name, "xltc"));

        private static OfficeXmlNode EnsureOwnedThreadedRoot(OfficeOpcEditablePackage package, string owner, string type,
            string contentType, string path, string elementName) {
            string? existing = OwnedThreadedParts(package, owner, type).FirstOrDefault().Target;
            if (existing != null) return OwnedThreadedRoot(package, existing);
            string target = FindUnusedObjectPartName(package, path, ".xml");
            byte[] markup = Encoding.UTF8.GetBytes("<xltc:" + elementName + " xmlns:xltc=\"" + ThreadedNamespace + "\"/>");
            package.CreateBinaryPart(target, contentType, markup);
            package.AddRelationship(owner, type, OfficeOpcPartNames.GetRelativeTarget(owner, target));
            return OwnedThreadedRoot(package, target);
        }

        private static IEnumerable<OfficeXmlNode> OwnedThreadedPeople(OfficeOpcEditablePackage package, string workbook) =>
            OwnedThreadedParts(package, workbook, PersonRelationship)
                .SelectMany(part => ThreadedChildren(OwnedThreadedRoot(package, part.Target), "person"));

        private static string EnsureOwnedThreadedPerson(OfficeOpcEditablePackage package, string workbook, string author,
            Dictionary<string, string>? people = null) {
            if (people != null && people.TryGetValue(author, out string? known)) return known;
            if (people == null) {
                foreach (OfficeXmlNode candidate in OwnedThreadedPeople(package, workbook)) {
                    if (string.Equals(candidate.Attribute("displayName"), author, StringComparison.OrdinalIgnoreCase)
                        && PresentThreadedValue(candidate.Attribute("id")) is string id) return id;
                }
            }

            OfficeXmlNode root = EnsureOwnedThreadedRoot(package, workbook, PersonRelationship, PersonContentType, "/xl/persons/person", "personList");
            OfficeXmlNode person = AppendThreadedElement(root, "person");
            string personId = BracedGuid();
            person.SetAttribute("displayName", author);
            person.SetAttribute("id", personId);
            people?.Add(author, personId);
            return personId;
        }
    }
}
