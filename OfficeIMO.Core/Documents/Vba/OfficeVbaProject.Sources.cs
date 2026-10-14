using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public sealed partial class OfficeVbaProject {
    /// <summary>Exports UTF-8 source files and a typed XML manifest to a new directory for version control.</summary>
    /// <remarks>The destination must not exist. Numbered file names avoid untrusted or platform-reserved module names.</remarks>
    public void ExportSources(string directory) {
        if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A source directory is required.", nameof(directory));
        string destination = Path.GetFullPath(directory);
        if (Directory.Exists(destination) || File.Exists(destination)) throw new IOException("The export destination already exists.");
        string parent = Path.GetDirectoryName(destination) ?? throw new ArgumentException("The source directory must have a parent.", nameof(directory));
        Directory.CreateDirectory(parent);
        string staging = Path.Combine(parent, ".vba-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(staging);
        try {
            var manifest = new XElement("vbaProject", new XAttribute("name", Name), new XAttribute("codePage", CodePage));
            for (int index = 0; index < _modules.Count; index++) {
                OfficeVbaModule module = _modules[index];
                string fileName = "module-" + (index + 1).ToString("D4", CultureInfo.InvariantCulture)
                    + (module.Kind == OfficeVbaModuleKind.Standard ? ".bas" : ".cls");
                string source = module.Source;
                if (module.Kind != OfficeVbaModuleKind.Standard) source = "VERSION 1.0 CLASS\r\nBEGIN\r\nEND\r\n" + source;
                File.WriteAllText(Path.Combine(staging, fileName), source, new UTF8Encoding(false, true));
                manifest.Add(new XElement("module", new XAttribute("name", module.Name), new XAttribute("kind", module.Kind), new XAttribute("file", fileName)));
            }
            new XDocument(manifest).Save(Path.Combine(staging, "vba-project.xml"));
            Directory.Move(staging, destination);
        } finally {
            if (Directory.Exists(staging)) Directory.Delete(staging, recursive: true);
        }
    }

    /// <summary>Imports the source files named in an exported manifest, with validation before any project mutation.</summary>
    /// <remarks>Missing modules may be added only for standard/class kinds. Omitted modules are retained; deletions are explicit.</remarks>
    public void ImportSources(string directory, int maximumSourceBytes = 64 * 1024 * 1024) {
        EnsureEditable();
        if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A source directory is required.", nameof(directory));
        if (maximumSourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumSourceBytes));
        string root = Path.GetFullPath(directory);
        RejectLinkedSource(root);
        RejectLinkedSource(Path.Combine(root, "vba-project.xml"));
        byte[] manifestBytes;
        using (var input = File.OpenRead(Path.Combine(root, "vba-project.xml"))) manifestBytes = OfficeStreamReader.ReadAllBytes(input, 1024 * 1024);
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 };
        XDocument manifest;
        using (var input = new MemoryStream(manifestBytes, writable: false))
        using (XmlReader reader = XmlReader.Create(input, settings)) manifest = XDocument.Load(reader);
        if (manifest.Root?.Name != "vbaProject") throw new InvalidDataException("The VBA source manifest has an invalid root.");
        var pending = new List<(string Name, OfficeVbaModuleKind Kind, string Source, OfficeVbaModule? Existing)>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int remaining = maximumSourceBytes;
        foreach (XElement entry in manifest.Root.Elements("module")) {
            string name = (string?)entry.Attribute("name") ?? throw new InvalidDataException("A source module has no name.");
            string file = (string?)entry.Attribute("file") ?? throw new InvalidDataException("A source module has no file.");
            if (!names.Add(name) || !files.Add(file) || file != Path.GetFileName(file) || file.IndexOfAny(new[] { '/', '\\', ':', '\0' }) >= 0
                || file == "." || file == "..") throw new InvalidDataException("The source manifest has duplicate or non-local file identities.");
            if (!Enum.TryParse((string?)entry.Attribute("kind"), out OfficeVbaModuleKind kind) || !Enum.IsDefined(typeof(OfficeVbaModuleKind), kind)) throw new InvalidDataException("The source module kind is invalid.");
            byte[] bytes;
            RejectLinkedSource(Path.Combine(root, file));
            using (var input = File.OpenRead(Path.Combine(root, file))) {
                if (remaining == 0) {
                    if (input.ReadByte() >= 0) throw new InvalidDataException("VBA source import exceeds the configured aggregate byte limit.");
                    bytes = Array.Empty<byte>();
                } else bytes = OfficeStreamReader.ReadAllBytes(input, remaining);
            }
            remaining -= bytes.Length;
            int offset = bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf ? 3 : 0;
            string source = new UTF8Encoding(false, true).GetString(bytes, offset, bytes.Length - offset);
            OfficeVbaModule? existing = _modules.FirstOrDefault(module => string.Equals(module.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null && existing.Kind != kind) throw new InvalidDataException("A source import cannot change a module's persistence kind.");
            if (existing == null) {
                OfficeVbaText.ValidateIdentifier(name);
                if (kind != OfficeVbaModuleKind.Standard && kind != OfficeVbaModuleKind.Class) throw new NotSupportedException("Source import cannot create a host document or form designer.");
                EnsureFreeStreamIdentity(name, null);
            }
            source = OfficeVbaText.NormalizeSource(source, existing?.Name ?? name, kind, existing?.Source);
            OfficeVbaText.Encode(source, CodePage);
            pending.Add((name, kind, source, existing));
        }
        if (_modules.Count + pending.Count(item => item.Existing == null) > 4096) throw new InvalidDataException("The source import exceeds the supported module count.");
        foreach (var item in pending) {
            if (item.Existing != null) item.Existing.Source = item.Source;
            else AddModuleCore(item.Name, item.Source, item.Kind);
        }
    }

    private static void RejectLinkedSource(string path) {
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) {
            throw new InvalidDataException("Source import does not follow linked directories or files.");
        }
    }
}
