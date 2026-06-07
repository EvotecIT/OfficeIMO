namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    private IReadOnlyList<string>? _attachmentNames;
    private IReadOnlyList<string>? _attachmentFileNames;
    private IReadOnlyList<string>? _attachmentSources;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>>? _attachmentsByName;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>>? _attachmentsByFileName;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>>? _attachmentsBySource;
    private IReadOnlyDictionary<PdfAssociatedFileRelationship, IReadOnlyList<PdfAttachmentInfo>>? _attachmentsByRelationship;

    /// <summary>Embedded and associated file attachment metadata discovered from the document catalog.</summary>
    public IReadOnlyList<PdfAttachmentInfo> Attachments { get; }

    /// <summary>Number of embedded and associated file attachments discovered from the document catalog.</summary>
    public int AttachmentCount => Attachments.Count;

    /// <summary>True when at least one embedded or associated file attachment was discovered.</summary>
    public bool HasAttachments => AttachmentCount > 0;

    /// <summary>Attachment name-tree keys or associated-file fallback names in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentNames {
        get {
            if (_attachmentNames is not null) {
                return _attachmentNames;
            }

            var names = new List<string>(Attachments.Count);
            for (int i = 0; i < Attachments.Count; i++) {
                names.Add(Attachments[i].Name);
            }

            _attachmentNames = names.AsReadOnly();
            return _attachmentNames;
        }
    }

    /// <summary>Distinct attachment file names in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentFileNames {
        get {
            if (_attachmentFileNames is not null) {
                return _attachmentFileNames;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var fileNames = new List<string>();
            for (int i = 0; i < Attachments.Count; i++) {
                string fileName = Attachments[i].FileName;
                if (seen.Add(fileName)) {
                    fileNames.Add(fileName);
                }
            }

            _attachmentFileNames = fileNames.AsReadOnly();
            return _attachmentFileNames;
        }
    }

    /// <summary>Distinct catalog attachment sources in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentSources {
        get {
            if (_attachmentSources is not null) {
                return _attachmentSources;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var sources = new List<string>();
            for (int i = 0; i < Attachments.Count; i++) {
                string source = Attachments[i].Source;
                if (seen.Add(source)) {
                    sources.Add(source);
                }
            }

            _attachmentSources = sources.AsReadOnly();
            return _attachmentSources;
        }
    }

    /// <summary>Attachments grouped by name-tree key or associated-file fallback name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>> AttachmentsByName {
        get {
            if (_attachmentsByName is not null) {
                return _attachmentsByName;
            }

            var grouped = new Dictionary<string, List<PdfAttachmentInfo>>(StringComparer.Ordinal);
            for (int i = 0; i < Attachments.Count; i++) {
                AddAttachment(grouped, Attachments[i].Name, Attachments[i]);
            }

            _attachmentsByName = ToReadOnlyLookup(grouped);
            return _attachmentsByName;
        }
    }

    /// <summary>Attachments grouped by file specification file name.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>> AttachmentsByFileName {
        get {
            if (_attachmentsByFileName is not null) {
                return _attachmentsByFileName;
            }

            var grouped = new Dictionary<string, List<PdfAttachmentInfo>>(StringComparer.Ordinal);
            for (int i = 0; i < Attachments.Count; i++) {
                AddAttachment(grouped, Attachments[i].FileName, Attachments[i]);
            }

            _attachmentsByFileName = ToReadOnlyLookup(grouped);
            return _attachmentsByFileName;
        }
    }

    /// <summary>Attachments grouped by catalog source.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfAttachmentInfo>> AttachmentsBySource {
        get {
            if (_attachmentsBySource is not null) {
                return _attachmentsBySource;
            }

            var grouped = new Dictionary<string, List<PdfAttachmentInfo>>(StringComparer.Ordinal);
            for (int i = 0; i < Attachments.Count; i++) {
                AddAttachment(grouped, Attachments[i].Source, Attachments[i]);
            }

            _attachmentsBySource = ToReadOnlyLookup(grouped);
            return _attachmentsBySource;
        }
    }

    /// <summary>Attachments grouped by associated-file relationship.</summary>
    public IReadOnlyDictionary<PdfAssociatedFileRelationship, IReadOnlyList<PdfAttachmentInfo>> AttachmentsByRelationship {
        get {
            if (_attachmentsByRelationship is not null) {
                return _attachmentsByRelationship;
            }

            var grouped = new Dictionary<PdfAssociatedFileRelationship, List<PdfAttachmentInfo>>();
            for (int i = 0; i < Attachments.Count; i++) {
                PdfAttachmentInfo attachment = Attachments[i];
                if (!grouped.TryGetValue(attachment.Relationship, out List<PdfAttachmentInfo>? attachments)) {
                    attachments = new List<PdfAttachmentInfo>();
                    grouped.Add(attachment.Relationship, attachments);
                }

                attachments.Add(attachment);
            }

            _attachmentsByRelationship = ToReadOnlyLookup(grouped);
            return _attachmentsByRelationship;
        }
    }

    /// <summary>Returns attachments with a matching name-tree key or associated-file fallback name.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsByName(string name) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        return AttachmentsByName.TryGetValue(name, out IReadOnlyList<PdfAttachmentInfo>? attachments)
            ? attachments
            : Array.Empty<PdfAttachmentInfo>();
    }

    /// <summary>Returns attachments with a matching file specification file name.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsByFileName(string fileName) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName));
        return AttachmentsByFileName.TryGetValue(fileName, out IReadOnlyList<PdfAttachmentInfo>? attachments)
            ? attachments
            : Array.Empty<PdfAttachmentInfo>();
    }

    /// <summary>Returns attachments from a matching catalog source.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsBySource(string source) {
        Guard.NotNullOrWhiteSpace(source, nameof(source));
        return AttachmentsBySource.TryGetValue(source, out IReadOnlyList<PdfAttachmentInfo>? attachments)
            ? attachments
            : Array.Empty<PdfAttachmentInfo>();
    }

    /// <summary>Returns attachments with a matching associated-file relationship.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsByRelationship(PdfAssociatedFileRelationship relationship) {
        return AttachmentsByRelationship.TryGetValue(relationship, out IReadOnlyList<PdfAttachmentInfo>? attachments)
            ? attachments
            : Array.Empty<PdfAttachmentInfo>();
    }

    private static void AddAttachment(Dictionary<string, List<PdfAttachmentInfo>> grouped, string key, PdfAttachmentInfo attachment) {
        if (!grouped.TryGetValue(key, out List<PdfAttachmentInfo>? attachments)) {
            attachments = new List<PdfAttachmentInfo>();
            grouped.Add(key, attachments);
        }

        attachments.Add(attachment);
    }
}
