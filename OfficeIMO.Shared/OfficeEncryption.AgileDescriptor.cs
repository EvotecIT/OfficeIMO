#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Shared {
    internal static partial class OfficeEncryption {
        private sealed class AgileDescriptor {
            public byte[] KeyDataSaltValue = Array.Empty<byte>();
            public int KeyDataSaltSize;
            public int KeyDataKeyBits;
            public string KeyDataHashAlgorithm = "SHA512";
            public byte[] PasswordSaltValue = Array.Empty<byte>();
            public int PasswordKeyBits;
            public int PasswordHashSize;
            public string PasswordHashAlgorithm = "SHA512";
            public int SpinCount;
            public byte[] EncryptedVerifierHashInput = Array.Empty<byte>();
            public byte[] EncryptedVerifierHashValue = Array.Empty<byte>();
            public byte[] EncryptedKeyValue = Array.Empty<byte>();
            public byte[] EncryptedHmacKey = Array.Empty<byte>();
            public byte[] EncryptedHmacValue = Array.Empty<byte>();

            public static AgileDescriptor Parse(byte[] encryptionInfoBytes) {
                if (encryptionInfoBytes.Length < 8) {
                    throw new InvalidDataException("EncryptionInfo stream is too small.");
                }

                ushort major = ReadUInt16(encryptionInfoBytes, 0);
                ushort minor = ReadUInt16(encryptionInfoBytes, 2);
                if (major != 4 || minor != 4) {
                    throw new NotSupportedException("Only Office Agile encryption is supported.");
                }

                string xml = Encoding.UTF8.GetString(encryptionInfoBytes, 8, encryptionInfoBytes.Length - 8).TrimEnd('\0', ' ', '\r', '\n', '\t');
                var document = XDocument.Parse(xml);
                XNamespace ns = EncryptionNamespace;
                XNamespace p = PasswordNamespace;

                var keyData = document.Root?.Element(ns + "keyData") ?? throw new InvalidDataException("EncryptionInfo is missing keyData.");
                var dataIntegrity = document.Root.Element(ns + "dataIntegrity") ?? throw new InvalidDataException("EncryptionInfo is missing dataIntegrity.");
                var encryptedKey = document.Root
                    .Element(ns + "keyEncryptors")?
                    .Elements(ns + "keyEncryptor")
                    .Select(e => e.Element(p + "encryptedKey"))
                    .FirstOrDefault(e => e != null) ?? throw new InvalidDataException("EncryptionInfo is missing password encryptedKey.");

                return new AgileDescriptor {
                    KeyDataSaltSize = ReadRequiredInt(keyData, "saltSize"),
                    KeyDataKeyBits = ReadRequiredInt(keyData, "keyBits"),
                    KeyDataHashAlgorithm = ReadRequiredString(keyData, "hashAlgorithm"),
                    KeyDataSaltValue = Convert.FromBase64String(ReadRequiredString(keyData, "saltValue")),
                    EncryptedHmacKey = Convert.FromBase64String(ReadRequiredString(dataIntegrity, "encryptedHmacKey")),
                    EncryptedHmacValue = Convert.FromBase64String(ReadRequiredString(dataIntegrity, "encryptedHmacValue")),
                    SpinCount = ReadRequiredInt(encryptedKey, "spinCount"),
                    PasswordKeyBits = ReadRequiredInt(encryptedKey, "keyBits"),
                    PasswordHashSize = ReadRequiredInt(encryptedKey, "hashSize"),
                    PasswordHashAlgorithm = ReadRequiredString(encryptedKey, "hashAlgorithm"),
                    PasswordSaltValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "saltValue")),
                    EncryptedVerifierHashInput = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedVerifierHashInput")),
                    EncryptedVerifierHashValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedVerifierHashValue")),
                    EncryptedKeyValue = Convert.FromBase64String(ReadRequiredString(encryptedKey, "encryptedKeyValue"))
                };
            }

            private static string ReadRequiredString(XElement element, string name) {
                return element.Attribute(name)?.Value ?? throw new InvalidDataException($"EncryptionInfo is missing '{name}'.");
            }

            private static int ReadRequiredInt(XElement element, string name) {
                return int.Parse(ReadRequiredString(element, name), CultureInfo.InvariantCulture);
            }
        }
    }
}
