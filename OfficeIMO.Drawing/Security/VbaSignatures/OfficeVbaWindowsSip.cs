using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Security;

internal readonly struct OfficeVbaContentBindingResult {
    internal OfficeVbaContentBindingResult(bool isSupported, bool isValid, string detail) {
        IsSupported = isSupported;
        IsValid = isValid;
        Detail = detail;
    }
    internal bool IsSupported { get; }
    internal bool IsValid { get; }
    internal string Detail { get; }
}

internal static class OfficeVbaWindowsSip {
    private static readonly Guid MicrosoftOfficeOpenXmlSip =
        new("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");

    internal static OfficeVbaContentBindingResult ValidateContentBinding(
        string filePath,
        string digestAlgorithmOid,
        byte[] expectedDigest) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return new OfficeVbaContentBindingResult(false, false,
                "Microsoft Office SIP content validation is available on Windows only.");
        }
        if (string.IsNullOrWhiteSpace(digestAlgorithmOid) || expectedDigest == null ||
            expectedDigest.Length == 0 || expectedDigest.Length > 1024) {
            return new OfficeVbaContentBindingResult(true, false,
                "The selected VBA signature does not contain a bounded Authenticode subject digest.");
        }
        if (!TryGetSubjectInterfacePackage(filePath, out Guid subjectGuid, out string detail)) {
            return new OfficeVbaContentBindingResult(false, false, detail);
        }

        IntPtr subjectGuidPointer = IntPtr.Zero;
        IntPtr fileNamePointer = IntPtr.Zero;
        IntPtr algorithmPointer = IntPtr.Zero;
        IntPtr indirectDataPointer = IntPtr.Zero;
        try {
            subjectGuidPointer = Marshal.AllocHGlobal(Marshal.SizeOf<Guid>());
            Marshal.StructureToPtr(subjectGuid, subjectGuidPointer, false);
            fileNamePointer = Marshal.StringToHGlobalUni(Path.GetFullPath(filePath));
            algorithmPointer = Marshal.StringToHGlobalAnsi(digestAlgorithmOid);
            var subject = new SipSubjectInfo {
                cbSize = checked((uint)Marshal.SizeOf<SipSubjectInfo>()),
                pgSubjectType = subjectGuidPointer,
                hFile = new IntPtr(-1),
                pwsFileName = fileNamePointer,
                pwsDisplayName = fileNamePointer,
                DigestAlgorithm = new CryptAlgorithmIdentifier { pszObjId = algorithmPointer },
                dwEncodingType = 0x00010001,
                dwUnionChoice = 0
            };

            uint signedDataBytes = 0;
            bool selected = CryptSIPGetSignedDataMsg(ref subject, out _, 0, ref signedDataBytes, IntPtr.Zero);
            int selectionError = Marshal.GetLastWin32Error();
            if ((!selected && selectionError != 122) || signedDataBytes == 0) {
                return new OfficeVbaContentBindingResult(true, false,
                    "The Office SIP could not select the highest-precedence VBA signature. Windows error " + selectionError + ".");
            }
            uint indirectDataBytes = 0;
            if (!CryptSIPCreateIndirectData(ref subject, ref indirectDataBytes, IntPtr.Zero) ||
                indirectDataBytes == 0 || indirectDataBytes > 1024 * 1024) {
                return NativeFailure("The Office SIP could not size the current VBA subject digest.");
            }
            indirectDataPointer = Marshal.AllocHGlobal(checked((int)indirectDataBytes));
            if (!CryptSIPCreateIndirectData(ref subject, ref indirectDataBytes, indirectDataPointer)) {
                return NativeFailure("The Office SIP could not calculate the current VBA subject digest.");
            }
            SipIndirectData indirect = Marshal.PtrToStructure<SipIndirectData>(indirectDataPointer);
            if (indirect.Digest.cbData == 0 || indirect.Digest.cbData > 1024 || indirect.Digest.pbData == IntPtr.Zero) {
                return new OfficeVbaContentBindingResult(true, false,
                    "The Office SIP returned an invalid subject digest.");
            }
            string? actualAlgorithm = Marshal.PtrToStringAnsi(indirect.DigestAlgorithm.pszObjId);
            byte[] actualDigest = new byte[indirect.Digest.cbData];
            Marshal.Copy(indirect.Digest.pbData, actualDigest, 0, actualDigest.Length);
            bool valid = string.Equals(actualAlgorithm, digestAlgorithmOid, StringComparison.Ordinal) &&
                FixedTimeEquals(actualDigest, expectedDigest);
            return new OfficeVbaContentBindingResult(true, valid,
                valid
                    ? "Microsoft's Office SIP reproduced the signed VBA subject digest."
                    : "Microsoft's Office SIP subject digest does not match the selected VBA signature.");
        } catch (Exception exception) when (exception is DllNotFoundException or EntryPointNotFoundException or
            BadImageFormatException or ArgumentException or OverflowException or IOException) {
            return new OfficeVbaContentBindingResult(false, false,
                "Microsoft Office SIP content validation failed. " + exception.Message);
        } finally {
            if (indirectDataPointer != IntPtr.Zero) Marshal.FreeHGlobal(indirectDataPointer);
            if (algorithmPointer != IntPtr.Zero) Marshal.FreeHGlobal(algorithmPointer);
            if (fileNamePointer != IntPtr.Zero) Marshal.FreeHGlobal(fileNamePointer);
            if (subjectGuidPointer != IntPtr.Zero) Marshal.FreeHGlobal(subjectGuidPointer);
        }
    }

    internal static bool IsAvailable(string filePath, out string detail) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            detail = "Microsoft Office SIP discovery is available on Windows only.";
            return false;
        }
        return TryGetSubjectInterfacePackage(filePath, out _, out detail);
    }

    private static bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
        subjectGuid = Guid.Empty;
        try {
            if (!CryptSIPRetrieveSubjectGuid(filePath, IntPtr.Zero, out subjectGuid)) {
                detail = "Windows could not resolve a registered Subject Interface Package for the Office file.";
                return false;
            }
            if (subjectGuid != MicrosoftOfficeOpenXmlSip) {
                detail = "The registered SIP subject GUID " + subjectGuid +
                    " is not Microsoft's OOXML Office SIP " + MicrosoftOfficeOpenXmlSip + ".";
                return false;
            }
            detail = "Microsoft's OOXML Office SIP is registered for the Office file.";
            return true;
        } catch (Exception exception) when (exception is DllNotFoundException or
            EntryPointNotFoundException or BadImageFormatException) {
            detail = "Microsoft Office SIP discovery failed. " + exception.Message;
            return false;
        }
    }

    private static OfficeVbaContentBindingResult NativeFailure(string message) =>
        new(true, false, message + " Windows error " + Marshal.GetLastWin32Error() + ".");

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        int difference = left.Length ^ right.Length;
        int maximum = Math.Max(left.Length, right.Length);
        for (int index = 0; index < maximum; index++) {
            byte leftValue = index < left.Length ? left[index] : (byte)0;
            byte rightValue = index < right.Length ? right[index] : (byte)0;
            difference |= leftValue ^ rightValue;
        }
        return difference == 0;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct CryptDataBlob { internal uint cbData; internal IntPtr pbData; }
    [StructLayout(LayoutKind.Sequential)]
    private struct CryptAlgorithmIdentifier { internal IntPtr pszObjId; internal CryptDataBlob Parameters; }
    [StructLayout(LayoutKind.Sequential)]
    private struct CryptAttributeTypeValue { internal IntPtr pszObjId; internal CryptDataBlob Value; }
    [StructLayout(LayoutKind.Sequential)]
    private struct SipIndirectData {
        internal CryptAttributeTypeValue Data;
        internal CryptAlgorithmIdentifier DigestAlgorithm;
        internal CryptDataBlob Digest;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct SipSubjectInfo {
        internal uint cbSize;
        internal IntPtr pgSubjectType;
        internal IntPtr hFile;
        internal IntPtr pwsFileName;
        internal IntPtr pwsDisplayName;
        internal uint dwReserved1;
        internal uint dwIntVersion;
        internal IntPtr hProv;
        internal CryptAlgorithmIdentifier DigestAlgorithm;
        internal uint dwFlags;
        internal uint dwEncodingType;
        internal uint dwReserved2;
        internal uint fdwCAPISettings;
        internal uint fdwSecuritySettings;
        internal uint dwIndex;
        internal uint dwUnionChoice;
        internal IntPtr psFlat;
        internal IntPtr pClientData;
    }

    [DllImport("crypt32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CryptSIPRetrieveSubjectGuid(string fileName, IntPtr fileHandle, out Guid subjectGuid);
    [DllImport("crypt32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CryptSIPGetSignedDataMsg(ref SipSubjectInfo subjectInfo, out uint encodingType,
        uint index, ref uint signedDataBytes, IntPtr signedData);
    [DllImport("crypt32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CryptSIPCreateIndirectData(ref SipSubjectInfo subjectInfo,
        ref uint indirectDataBytes, IntPtr indirectData);
}
