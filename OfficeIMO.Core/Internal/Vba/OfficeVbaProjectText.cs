using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Preserves PROJECT declarations and reads the standard obfuscated protection fields.</summary>
internal static class OfficeVbaProjectText {
    internal static OfficeVbaModuleKind GetModuleKind(string text, string name, ushort type) {
        foreach (string line in Lines(text)) {
            if (line.Equals("BaseClass=" + name, StringComparison.OrdinalIgnoreCase)) return OfficeVbaModuleKind.Designer;
            if (line.StartsWith("Document=" + name + "/", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("DocClass=" + name + "/", StringComparison.OrdinalIgnoreCase)
                || line.Equals("DocModule=" + name, StringComparison.OrdinalIgnoreCase)) return OfficeVbaModuleKind.Document;
        }
        return type == 0x0021 ? OfficeVbaModuleKind.Standard : OfficeVbaModuleKind.Class;
    }

    internal static string Update(string original, IReadOnlyList<OfficeVbaModule> modules, IReadOnlyList<string> deletedNames) {
        var result = new List<string>();
        var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string line in Lines(original)) {
            int equals = line.IndexOf('=');
            if (equals <= 0 || !IsDeclaration(line.Substring(0, equals))) continue;
            string name = line.Substring(equals + 1).Split('/')[0];
            if (!declarations.ContainsKey(name)) declarations.Add(name, line);
        }
        bool inserted = false, inWorkspace = false;
        foreach (string line in Lines(original)) {
            int equals = line.IndexOf('=');
            string key = equals > 0 ? line.Substring(0, equals) : string.Empty;
            if (line.StartsWith("[", StringComparison.Ordinal)) inWorkspace = line.Equals("[Workspace]", StringComparison.OrdinalIgnoreCase);
            if (IsDeclaration(key)) {
                if (!inserted) { AddDeclarations(result, modules, declarations); inserted = true; }
                continue;
            }
            if (!inserted && (key.Equals("Name", StringComparison.OrdinalIgnoreCase) || line.StartsWith("[", StringComparison.Ordinal))) {
                AddDeclarations(result, modules, declarations); inserted = true;
            }
            if (inWorkspace && deletedNames.Any(name => line.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase))) continue;
            string rewritten = line;
            foreach (OfficeVbaModule module in modules.Where(module => module.Name != module.OriginalName)) {
                if (inWorkspace && line.StartsWith(module.OriginalName + "=", StringComparison.OrdinalIgnoreCase)) {
                    rewritten = module.Name + line.Substring(module.OriginalName.Length); break;
                }
            }
            result.Add(rewritten);
        }
        if (!inserted) AddDeclarations(result, modules, declarations);
        return string.Join("\r\n", result);
    }

    private static bool IsDeclaration(string key) => new[] { "Module", "Class", "Document", "BaseClass", "DocModule", "DocClass" }
        .Contains(key, StringComparer.OrdinalIgnoreCase);

    private static void AddDeclarations(ICollection<string> output, IReadOnlyList<OfficeVbaModule> modules, IReadOnlyDictionary<string, string> original) {
        foreach (OfficeVbaModule module in modules) {
            if (!module.IsNew && original.TryGetValue(module.OriginalName, out string? declaration)) {
                int start = declaration.IndexOf('=') + 1;
                output.Add(declaration.Substring(0, start) + module.Name + declaration.Substring(start + module.OriginalName.Length));
                continue;
            }
            string prefix = module.Kind == OfficeVbaModuleKind.Standard ? "Module="
                : module.Kind == OfficeVbaModuleKind.Class ? "Class="
                : module.Kind == OfficeVbaModuleKind.Designer ? "BaseClass=" : module.IsDocumentClass ? "DocClass=" : "Document=";
            output.Add(prefix + module.Name + (module.Kind == OfficeVbaModuleKind.Document ? "/&H00000000" : string.Empty));
        }
    }

    internal static bool IsProtected(string text) {
        foreach (string line in Lines(text)) {
            if (line.StartsWith("CMG=", StringComparison.OrdinalIgnoreCase)) {
                byte[] value = DecodeProtection(GetQuotedValue(line));
                if (value.Length != 4) throw new InvalidDataException("The VBA project protection state is invalid.");
                if (value.Any(item => item != 0)) return true;
            }
            if (line.StartsWith("DPB=", StringComparison.OrdinalIgnoreCase)) {
                byte[] value = DecodeProtection(GetQuotedValue(line));
                if (value.Length != 1 || value[0] != 0) return true;
            }
        }
        return false;
    }

    internal static string Create(string name) {
        string id = Guid.NewGuid().ToString("B").ToUpperInvariant();
        return "ID=\"" + id + "\"\r\nName=\"" + name + "\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\n"
            + "CMG=\"" + EncodeProtection(new byte[4], id) + "\"\r\nDPB=\"" + EncodeProtection(new byte[1], id) + "\"\r\n"
            + "GC=\"" + EncodeProtection(new byte[] { 0xff }, id) + "\"\r\n\r\n[Host Extender Info]\r\n"
            + "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\r\n[Workspace]\r\n";
    }

    private static string EncodeProtection(byte[] data, string projectId) {
        const byte seed = 0; // Obfuscation, not confidentiality; a deterministic seed is permitted by MS-OVBA.
        byte key = unchecked((byte)Encoding.ASCII.GetBytes(projectId).Sum(value => (int)value));
        var output = new List<byte> { seed, 2, key };
        byte plainPrevious = key, encryptedPrevious = key, encryptedBefore = 2;
        foreach (byte value in BitConverter.GetBytes(data.Length).Concat(data)) {
            byte encoded = (byte)(value ^ unchecked((byte)(encryptedBefore + plainPrevious)));
            output.Add(encoded); encryptedBefore = encryptedPrevious; encryptedPrevious = encoded; plainPrevious = value;
        }
        return BitConverter.ToString(output.ToArray()).Replace("-", string.Empty);
    }

    private static byte[] DecodeProtection(string hex) {
        if (hex.Length < 14 || (hex.Length & 1) != 0) throw new InvalidDataException("The VBA protection field is truncated.");
        byte[] input = new byte[hex.Length / 2];
        for (int index = 0; index < input.Length; index++) {
            if (!byte.TryParse(hex.Substring(index * 2, 2), System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out input[index])) throw new InvalidDataException("The VBA protection field is not hexadecimal.");
        }
        byte seed = input[0];
        if ((seed ^ input[1]) != 2) throw new InvalidDataException("The VBA protection field has an unsupported version.");
        byte plainPrevious = (byte)(seed ^ input[2]), encryptedPrevious = input[2], encryptedBefore = input[1];
        var decoded = new List<byte>();
        for (int index = 3; index < input.Length; index++) {
            byte plain = (byte)(input[index] ^ unchecked((byte)(encryptedBefore + plainPrevious)));
            decoded.Add(plain); encryptedBefore = encryptedPrevious; encryptedPrevious = input[index]; plainPrevious = plain;
        }
        int ignored = (seed & 6) / 2;
        if (decoded.Count < ignored + 4) throw new InvalidDataException("The VBA protection field has no length.");
        byte[] decodedBytes = decoded.ToArray();
        uint length = BitConverter.ToUInt32(decodedBytes, ignored);
        if (length != decoded.Count - ignored - 4) throw new InvalidDataException("The VBA protection field length is inconsistent.");
        return decoded.Skip(ignored + 4).ToArray();
    }

    private static string GetQuotedValue(string line) {
        int start = line.IndexOf('"'); int end = line.LastIndexOf('"');
        if (start < 0 || end <= start) throw new InvalidDataException("The VBA protection value is not quoted.");
        return line.Substring(start + 1, end - start - 1);
    }

    private static IEnumerable<string> Lines(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
}
