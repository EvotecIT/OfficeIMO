using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Core.Internal;
using DirectoryModel = OfficeIMO.Core.Internal.OfficeVbaDirectoryCodec.DirectoryModel;

namespace OfficeIMO;

/// <summary>Reads and edits a host-neutral VBA compound project without executing or compiling its source.</summary>
public sealed partial class OfficeVbaProject {
    private readonly byte[] _originalBytes;
    private readonly OfficeCompoundFile _compound;
    private readonly DirectoryModel _directory;
    private readonly List<OfficeVbaModule> _modules = new();
    private readonly List<OfficeVbaReference> _references = new();
    private readonly List<string> _deletedModules = new();
    private bool _referencesChanged;
    private int _originalExpandedBytes;

    private OfficeVbaProject(byte[] originalBytes, OfficeCompoundFile compound, DirectoryModel directory, string projectText) {
        _originalBytes = originalBytes; _compound = compound; _directory = directory; _projectText = projectText;
        Name = OfficeVbaText.Decode(directory.ProjectName, directory.CodePage);
        CodePage = directory.CodePage;
        Modules = _modules.AsReadOnly();
        References = _references.AsReadOnly();
    }

    /// <summary>Gets the project's persisted name.</summary>
    public string Name { get; }
    /// <summary>Gets the Windows code page used by source and ANSI project records.</summary>
    public int CodePage { get; }
    /// <summary>Gets the source modules in directory order.</summary>
    public IReadOnlyList<OfficeVbaModule> Modules { get; }
    /// <summary>Gets the project's explicit library references.</summary>
    public IReadOnlyList<OfficeVbaReference> References { get; }
    /// <summary>Gets whether the project declares password or access protection; protected projects remain readable.</summary>
    public bool IsProtected { get; private set; }
    /// <summary>Gets whether editing operations differ from the loaded source project.</summary>
    public bool HasChanges => _referencesChanged || _deletedModules.Count > 0
        || _modules.Any(module => module.IsNew || module.Name != module.OriginalName || module.Source != module.OriginalSource);

    /// <summary>Loads a detached project with aggregate input, compound-stream, and expanded-source limits.</summary>
    /// <exception cref="InvalidDataException">The project is malformed or its persisted text encoding is invalid or unavailable.</exception>
    public static OfficeVbaProject Load(byte[] projectBytes, OfficeVbaReadOptions? options = null) {
        if (projectBytes == null) throw new ArgumentNullException(nameof(projectBytes));
        options ??= new OfficeVbaReadOptions();
        ValidateLimits(options.MaximumProjectBytes, options.MaximumExpandedBytes);
        if (projectBytes.Length > options.MaximumProjectBytes) throw new InvalidDataException("The VBA project exceeds the configured input byte limit.");
        byte[] bytes = (byte[])projectBytes.Clone();
        var readOptions = new OfficeCompoundReadOptions(maxStreamBytes: options.MaximumProjectBytes,
            maxTotalStreamBytes: options.MaximumProjectBytes);
        if (!OfficeCompoundFileReader.TryRead(bytes, readOptions, out OfficeCompoundFile? compound, out string? error) || compound == null) {
            throw new InvalidDataException(error ?? "The VBA project is not a valid compound file.");
        }
        try { return LoadContents(bytes, compound, options); }
        catch (Exception exception) when (exception is System.Text.DecoderFallbackException || exception is NotSupportedException) {
            throw new InvalidDataException("The VBA project has invalid text or an unavailable persisted encoding.", exception);
        }
    }

    private static OfficeVbaProject LoadContents(byte[] bytes, OfficeCompoundFile compound, OfficeVbaReadOptions options) {
        if (!compound.Streams.TryGetValue("VBA/dir", out byte[]? compressedDirectory)
            || !compound.Streams.TryGetValue("PROJECT", out byte[]? projectText)) {
            throw new InvalidDataException("A VBA project must contain PROJECT and VBA/dir streams.");
        }
        if (!OfficeVbaCompression.TryDecompress(compressedDirectory, options.MaximumExpandedBytes, out byte[] directoryBytes, out string detail)
            || !DirectoryModel.TryParse(directoryBytes, options.MaximumExpandedBytes, out DirectoryModel? directory, out detail, includeSignatureTranscripts: false)
            || directory == null) throw new InvalidDataException(detail);
        var project = new OfficeVbaProject(bytes, compound, directory, OfficeVbaText.Decode(projectText, directory.CodePage));
        project.IsProtected = OfficeVbaProjectText.IsProtected(project._projectText);
        int remaining = options.MaximumExpandedBytes - directoryBytes.Length;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var streams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeVbaDirectoryCodec.ModuleModel module in directory.Modules) {
            string name = module.UnicodeName.Length > 0
                ? new System.Text.UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                : OfficeVbaText.Decode(module.AnsiName, directory.CodePage);
            if (!names.Add(name) || !streams.Add(module.StreamName)) throw new InvalidDataException("VBA modules have duplicate names or stream identities.");
            if (!compound.Streams.TryGetValue("VBA/" + module.StreamName, out byte[]? stream)
                || module.TextOffset > stream.Length) throw new InvalidDataException("A VBA module stream is missing or its source offset is outside the stream.");
            var compressedSource = new byte[stream.Length - module.TextOffset];
            Buffer.BlockCopy(stream, module.TextOffset, compressedSource, 0, compressedSource.Length);
            if (!OfficeVbaCompression.TryDecompress(compressedSource, remaining, out byte[] source, out detail)) throw new InvalidDataException(detail);
            remaining -= source.Length;
            OfficeVbaModuleKind kind = OfficeVbaProjectText.GetModuleKind(project._projectText, name, module.TypeId);
            project._modules.Add(new OfficeVbaModule(name, kind, OfficeVbaText.Decode(source, directory.CodePage).TrimEnd('\0'), module));
        }
        foreach (OfficeVbaDirectoryCodec.ReferenceModel reference in directory.References) {
            project._references.Add(OfficeVbaDirectoryWriter.ReadReference(reference.Serialized, directory.CodePage));
        }
        project._originalExpandedBytes = options.MaximumExpandedBytes - remaining;
        return project;
    }

    /// <summary>Creates a template-free VBA project with a native directory and source-only compilation state.</summary>
    public static OfficeVbaProject Create(string name = "VBAProject", int codePage = 1252) {
        OfficeVbaText.ValidateIdentifier(name, 128);
        if (codePage <= 0 || codePage > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(codePage));
        return Load(OfficeVbaDirectoryWriter.CreateProject(name, codePage));
    }

    /// <summary>Gets a module by its case-insensitive logical name.</summary>
    public OfficeVbaModule GetModule(string name) => _modules.FirstOrDefault(module => string.Equals(module.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new KeyNotFoundException("The VBA module '" + name + "' does not exist.");

    /// <summary>Replaces source while preserving module kind, attributes, and the rest of the project.</summary>
    public void SetModuleSource(string name, string source) {
        EnsureEditable();
        OfficeVbaModule module = GetModule(name);
        string normalized = OfficeVbaText.NormalizeSource(source, module.Name, module.Kind, module.Source);
        OfficeVbaText.Encode(normalized, CodePage);
        module.Source = normalized;
    }

    /// <summary>Adds a standard or ordinary class module; host document and designer identities require their own owners.</summary>
    public OfficeVbaModule AddModule(string name, string source, OfficeVbaModuleKind kind = OfficeVbaModuleKind.Standard) {
        if (kind != OfficeVbaModuleKind.Standard && kind != OfficeVbaModuleKind.Class) {
            throw new ArgumentException("Use AddDocumentModule for a host document. Form designers are preserved but not created by this source editor.", nameof(kind));
        }
        return AddModuleCore(name, source, kind);
    }

    /// <summary>Adds a host-owned module whose VB_Base class identity is supplied by the document adapter.</summary>
    public OfficeVbaModule AddDocumentModule(string name, string source, Guid baseClassId) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (baseClassId == Guid.Empty) throw new ArgumentException("A host document requires its base-class identity.", nameof(baseClassId));
        return AddDocumentModuleWithIdentity(name, source, "0" + baseClassId.ToString("B").ToUpperInvariant());
    }

    internal OfficeVbaModule AddDocumentModuleWithIdentity(string name, string source, string baseIdentity) {
        string attributes = "Attribute VB_Base = \"" + baseIdentity + "\"\r\n"
            + "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n";
        return AddModuleCore(name, attributes + source, OfficeVbaModuleKind.Document);
    }

    private OfficeVbaModule AddModuleCore(string name, string source, OfficeVbaModuleKind kind) {
        EnsureEditable();
        OfficeVbaText.ValidateIdentifier(name);
        if (_modules.Any(module => string.Equals(module.Name, name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("A VBA module with this name already exists.", nameof(name));
        EnsureFreeStreamIdentity(name, null);
        if (_modules.Count >= 4096) throw new InvalidOperationException("The project exceeds the supported module count.");
        string normalized = OfficeVbaText.NormalizeSource(source, name, kind);
        OfficeVbaText.Encode(normalized, CodePage);
        var directory = OfficeVbaDirectoryWriter.NewModule(name, kind, CodePage);
        var module = new OfficeVbaModule(name, kind, normalized, directory) { IsNew = true };
        _modules.Add(module);
        return module;
    }

    /// <summary>Renames a standard or class module and its project/stream identities; source code references are not rewritten.</summary>
    public void RenameModule(string name, string newName) {
        EnsureEditable();
        OfficeVbaModule module = GetModule(name);
        EnsureStructuralModule(module);
        OfficeVbaText.ValidateIdentifier(newName);
        if (_modules.Any(other => other != module && string.Equals(other.Name, newName, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("A VBA module with this name already exists.", nameof(newName));
        EnsureFreeStreamIdentity(newName, module);
        string source = OfficeVbaText.NormalizeSource(module.Source, newName, module.Kind);
        OfficeVbaText.Encode(source, CodePage);
        module.Name = newName;
        module.Source = source;
    }

    /// <summary>Deletes a standard or class module. Host-owned and designer modules remain under their owning document.</summary>
    public void DeleteModule(string name) {
        EnsureEditable();
        OfficeVbaModule module = GetModule(name);
        EnsureStructuralModule(module);
        if (!module.IsNew) _deletedModules.Add(module.OriginalName);
        _modules.Remove(module);
    }

    /// <summary>Adds a registered type-library reference without installing or executing that library.</summary>
    /// <remarks>Major and minor versions must each be between 0 and 65535, as required by the native reference format.</remarks>
    public void AddRegisteredReference(string name, Guid typeLibraryId, int majorVersion = 1, int minorVersion = 0, string path = "") {
        EnsureEditable();
        OfficeVbaText.ValidateIdentifier(name, 128);
        if (typeLibraryId == Guid.Empty) throw new ArgumentException("A type library identity is required.", nameof(typeLibraryId));
        if (majorVersion < 0 || majorVersion > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(majorVersion));
        if (minorVersion < 0 || minorVersion > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(minorVersion));
        if (path == null || path.IndexOfAny(new[] { '#', '\r', '\n', '\0' }) >= 0) throw new ArgumentException("The type-library path cannot contain reference separators.", nameof(path));
        string identity = typeLibraryId.ToString("B").ToUpperInvariant();
        if (_references.Any(reference => reference.LibraryId?.StartsWith("*\\G" + identity + "#", StringComparison.OrdinalIgnoreCase) == true)) return;
        if (_references.Any(reference => string.Equals(reference.Name, name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("A reference with this name already exists.", nameof(name));
        string libid = "*\\G" + identity + "#" + majorVersion.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + "."
            + minorVersion.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + "#0#" + path + "#" + name;
        _references.Add(new OfficeVbaReference(name, libid, OfficeVbaDirectoryWriter.RegisteredReference(name, libid, CodePage)));
        _referencesChanged = true;
    }

    /// <summary>Removes one explicit reference without rewriting source that uses its types.</summary>
    public bool RemoveReference(string name) {
        EnsureEditable();
        OfficeVbaReference[] matches = _references.Where(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length > 1) throw new ArgumentException("The reference name is ambiguous; no reference was removed.", nameof(name));
        OfficeVbaReference? reference = matches.FirstOrDefault();
        if (reference == null) return false;
        _references.Remove(reference); _referencesChanged = true; return true;
    }

    private void EnsureEditable() {
        if (IsProtected) throw new InvalidOperationException("This VBA project is protected. Source editing does not bypass or remove project protection.");
    }

    private void EnsureFreeStreamIdentity(string name, OfficeVbaModule? except) {
        if (name.Equals("dir", StringComparison.OrdinalIgnoreCase) || name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("__SRP_", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("The module name is reserved for VBA project infrastructure.", nameof(name));
        if (_modules.Any(module => module != except && string.Equals(module.Directory.StreamName, name, StringComparison.OrdinalIgnoreCase)
            && module.Name == module.OriginalName && !module.IsNew)) throw new ArgumentException("A module stream with this identity already exists.", nameof(name));
        if (_compound.Entries.Any(entry => entry.IsStorage && string.Equals(entry.Path, "VBA/" + name, StringComparison.OrdinalIgnoreCase))) {
            throw new ArgumentException("The module name collides with a preserved project storage.", nameof(name));
        }
        if (_compound.Streams.ContainsKey("VBA/" + name)
            && !_directory.Modules.Any(module => string.Equals(module.StreamName, name, StringComparison.OrdinalIgnoreCase))) {
            throw new ArgumentException("The module name collides with an opaque project stream.", nameof(name));
        }
    }

    private static void EnsureStructuralModule(OfficeVbaModule module) {
        if (module.Kind == OfficeVbaModuleKind.Document || module.Kind == OfficeVbaModuleKind.Designer) {
            throw new NotSupportedException("Host document and designer modules can be edited, but their structural identities require the owning document or form designer.");
        }
    }

    private static void ValidateLimits(int project, int expanded) {
        if (project <= 0 || expanded <= 0) throw new ArgumentOutOfRangeException(nameof(project), "VBA byte limits must be positive.");
    }
}
