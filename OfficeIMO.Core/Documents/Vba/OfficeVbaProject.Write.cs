using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public sealed partial class OfficeVbaProject {
    private readonly string _projectText;

    /// <summary>Produces a bounded complete project; unchanged projects retain their original bytes and signatures.</summary>
    /// <remarks>Edits emit source-only changed module streams, normalize existing raw source chunks, and invalidate project/compiled caches. No source is executed.</remarks>
    public OfficeVbaWriteResult Write(OfficeVbaWriteOptions? options = null) {
        options ??= new OfficeVbaWriteOptions();
        ValidateLimits(options.MaximumProjectBytes, options.MaximumExpandedBytes);
        if (!HasChanges) {
            if (_originalBytes.Length > options.MaximumProjectBytes) throw new InvalidDataException("The VBA project exceeds the configured output byte limit.");
            if (_originalExpandedBytes > options.MaximumExpandedBytes) throw new InvalidDataException("The VBA directory and source exceed the configured aggregate expanded byte limit.");
            return new OfficeVbaWriteResult((byte[])_originalBytes.Clone(), false, Array.Empty<string>());
        }
        EnsureEditable();
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        // VBA signatures belong to the containing document's signature carrier, not to arbitrary
        // similarly named opaque streams in the MS-OVBA project. The host adapter owns that policy.
        var removals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeCompoundFileEntry entry in _compound.Entries) {
            if (entry.IsStream && entry.Path.StartsWith("VBA/__SRP_", StringComparison.OrdinalIgnoreCase)) removals.Add(entry.Path);
        }
        foreach (OfficeVbaDirectoryCodec.ModuleModel oldModule in _directory.Modules) {
            OfficeVbaModule? current = _modules.FirstOrDefault(module => !module.IsNew && module.Directory == oldModule);
            if (current == null || current.Name != current.OriginalName) removals.Add("VBA/" + oldModule.StreamName);
        }
        byte[] directory = OfficeVbaDirectoryWriter.SerializeDirectory(_directory, _modules, _references, CodePage);
        int remaining = options.MaximumExpandedBytes - directory.Length;
        if (remaining < 0) throw new InvalidDataException("The VBA directory exceeds the aggregate expanded byte limit.");
        var changed = new List<string>(_deletedModules);
        foreach (OfficeVbaModule module in _modules) {
            bool moduleChanged = module.IsNew || module.Name != module.OriginalName || module.Source != module.OriginalSource;
            byte[] source = OfficeVbaText.Encode(module.Source, CodePage);
            if (source.Length > remaining) throw new InvalidDataException("VBA module source exceeds the aggregate expanded byte limit.");
            remaining -= source.Length;
            if (moduleChanged) {
                string streamName = module.IsNew || module.Name != module.OriginalName ? module.Name : module.Directory.StreamName;
                string path = "VBA/" + streamName;
                // A deleted/recreated or renamed module can intentionally take an old module's stream identity.
                removals.Remove(path);
                replacements[path] = OfficeVbaCompression.Compress(source);
                changed.Add(module.Name);
            } else {
                string path = "VBA/" + module.Directory.StreamName;
                byte[] original = _compound.Streams[path];
                if (OfficeVbaCompression.ContainsRawChunk(original, module.Directory.TextOffset)) {
                    // Project cache invalidation also makes untouched source reachable by the
                    // native loader. Preserve its source offset and prefix while replacing raw chunks.
                    byte[] compressed = OfficeVbaCompression.Compress(source);
                    byte[] normalized = new byte[checked(module.Directory.TextOffset + compressed.Length)];
                    Buffer.BlockCopy(original, 0, normalized, 0, module.Directory.TextOffset);
                    Buffer.BlockCopy(compressed, 0, normalized, module.Directory.TextOffset, compressed.Length);
                    replacements[path] = normalized;
                    changed.Add(module.Name);
                }
            }
        }
        replacements["VBA/dir"] = OfficeVbaCompression.Compress(directory);
        replacements["VBA/_VBA_PROJECT"] = new byte[] { 0xcc, 0x61, 0xff, 0xff, 0, 1, 0 };
        replacements["PROJECT"] = OfficeVbaText.Encode(OfficeVbaProjectText.Update(_projectText, _modules, _deletedModules), CodePage);
        replacements["PROJECTwm"] = OfficeVbaDirectoryWriter.ProjectNames(_modules, CodePage);
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(_compound, replacements, removals, options.MaximumProjectBytes);
        // Validate the exact produced artifact through the common parser before exposing it to a document adapter.
        Load(bytes, new OfficeVbaReadOptions { MaximumProjectBytes = options.MaximumProjectBytes, MaximumExpandedBytes = options.MaximumExpandedBytes });
        return new OfficeVbaWriteResult(bytes, true, changed.AsReadOnly());
    }
}
