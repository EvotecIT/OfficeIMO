using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Security;

public static partial class OfficeVbaSignatureService {
    /// <summary>Creates legacy, agile, and V3 VBA signatures through Microsoft's registered Office SIP.</summary>
    public static OfficeVbaSigningResult TrySign(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        string certificateThumbprint,
        OfficeVbaSigningOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        options ??= new OfficeVbaSigningOptions();
        ValidateOptions(options);
        string fullPath = NormalizePath(filePath);
        var findings = new List<OfficeVbaSignatureFinding>();
        OfficeVbaSignatureInfo source = Inspect(fullPath, options);
        if (!source.IsMacroEnabledFormat || !source.HasMacroProject ||
            source.Findings.Any(finding => finding.State == OfficePackageSignatureValidationState.Failed)) {
            findings.AddRange(source.Findings);
            findings.Add(Finding("VbaSigningPreflightFailed", OfficePackageSignatureValidationState.Failed,
                "VBA signing requires a valid macro-enabled package with vbaProject.bin."));
            return SigningResult(fullPath, true, false, null, findings);
        }
        if (!TryNormalizeThumbprint(certificateThumbprint, out string thumbprint, out string thumbprintDetail)) {
            findings.Add(Finding("CertificateThumbprintInvalid", OfficePackageSignatureValidationState.Failed, thumbprintDetail));
            return SigningResult(fullPath, true, false, null, findings);
        }
        if (options.ToolTimeout <= TimeSpan.Zero || options.ToolTimeout.TotalMilliseconds > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options.ToolTimeout));
        }
        if (options.MaxToolOutputCharacters <= 0 || options.MaxToolOutputCharacters > 4 * 1024 * 1024) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxToolOutputCharacters));
        }
        if (options.TimestampAuthorityUrl != null && (!options.TimestampAuthorityUrl.IsAbsoluteUri ||
            !(options.TimestampAuthorityUrl.Scheme == Uri.UriSchemeHttps || options.TimestampAuthorityUrl.Scheme == Uri.UriSchemeHttp))) {
            throw new ArgumentException("TimestampAuthorityUrl must be an absolute HTTP or HTTPS URL.", nameof(options));
        }
        if (!OfficeVbaWindowsSip.IsAvailable(fullPath, out string sipDetail)) {
            findings.Add(Finding("MicrosoftOfficeSipUnavailable", OfficePackageSignatureValidationState.Unsupported, sipDetail));
            return SigningResult(fullPath, false, false, null, findings);
        }
        string signTool = string.IsNullOrWhiteSpace(options.SignToolPath)
            ? "signtool.exe" : Path.GetFullPath(options.SignToolPath!);
        string clearTool = string.IsNullOrWhiteSpace(options.OfficeSipsDirectory)
            ? string.Empty : Path.Combine(Path.GetFullPath(options.OfficeSipsDirectory!), "offclearsig.exe");
        if (string.IsNullOrWhiteSpace(clearTool) || !File.Exists(clearTool)) {
            findings.Add(Finding("OfficeSignatureClearToolUnavailable", OfficePackageSignatureValidationState.Unsupported,
                "OfficeSipsDirectory must contain Microsoft's offclearsig.exe."));
            return SigningResult(fullPath, false, false, null, findings);
        }

        OfficePackageSignatureInfo packageSignatures = OfficePackageSignatureService.Inspect(fullPath,
            new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
        if (packageSignatures.HasSignatures && !options.AllowPackageSignatureInvalidation) {
            findings.Add(Finding("ExistingPackageSignatureInvalidationBlocked", OfficePackageSignatureValidationState.Failed,
                "VBA signing would invalidate existing OPC package signatures."));
            return SigningResult(fullPath, true, false, null, findings);
        }

        string stagingPath = string.Empty;
        try {
            stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
            OfficePackageFileSnapshot.CopyBounded(fullPath, stagingPath, options.Package.MaxPackageBytes);
            string sourceHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.Package.MaxPackageBytes);
            ProcessResult clear = RunTool(clearTool, new[] { stagingPath }, options);
            if (!clear.Succeeded) {
                findings.Add(ToolFinding("OfficeSignatureClearFailed", "offclearsig.exe", clear));
                return SigningResult(fullPath, true, false, null, findings);
            }

            foreach (OfficeVbaSignatureProfile profile in new[] {
                OfficeVbaSignatureProfile.Legacy,
                OfficeVbaSignatureProfile.Agile,
                OfficeVbaSignatureProfile.V3 }) {
                ProcessResult sign = RunTool(signTool, BuildSignArguments(stagingPath, thumbprint, options), options);
                if (!sign.Succeeded) {
                    findings.Add(ToolFinding("VbaSignatureToolFailed", "signtool.exe", sign, profile));
                    return SigningResult(fullPath, true, false, null, findings);
                }
                OfficeVbaSignatureInfo profileReadback = Inspect(stagingPath, options);
                if (!profileReadback.Signatures.Any(signature => signature.Profile == profile)) {
                    findings.Add(Finding("VbaSignatureProfileMissing", OfficePackageSignatureValidationState.Failed,
                        "SignTool did not create the expected " + profile + " VBA signature profile.", profile));
                    return SigningResult(fullPath, true, false, null, findings);
                }
            }

            OfficeVbaSignatureValidationResult validation = Validate(stagingPath, securityProvider, options);
            bool hasAllProfiles = validation.SignatureInfo.Signatures.Select(signature => signature.Profile).Distinct().Count() == 3;
            if (!hasAllProfiles || !validation.IsValidUnderPolicy) {
                findings.AddRange(validation.Findings);
                findings.Add(Finding("VbaSignatureReadbackFailed", OfficePackageSignatureValidationState.Failed,
                    "The completed VBA signatures did not satisfy profile, content-binding, CMS, and trust policy."));
                return SigningResult(fullPath, true, false, validation, findings);
            }
            string validatedHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.Package.MaxPackageBytes);
            if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                stagingPath, fullPath,
                displaced => string.Equals(sourceHash,
                    OfficePackageFileSnapshot.ComputeSha256(displaced, options.Package.MaxPackageBytes), StringComparison.Ordinal),
                installed => string.Equals(validatedHash,
                    OfficePackageFileSnapshot.ComputeSha256(installed, options.Package.MaxPackageBytes), StringComparison.Ordinal))) {
                stagingPath = string.Empty;
                findings.Add(Finding("SourcePackageChangedDuringSigning", OfficePackageSignatureValidationState.Failed,
                    "The package changed while VBA signatures were staged; the current source was preserved."));
                return SigningResult(fullPath, true, false, validation, findings);
            }
            stagingPath = string.Empty;
            findings.Add(Finding("VbaSignaturesCommitted", OfficePackageSignatureValidationState.Passed,
                "Legacy, agile, and V3 VBA signatures were validated and atomically committed."));
            return SigningResult(fullPath, true, true, validation, findings);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
            InvalidDataException or ArgumentException or OverflowException) {
            findings.Add(Finding("VbaSigningFailed", OfficePackageSignatureValidationState.Failed,
                "VBA signing failed before atomic commit. " + exception.Message));
            return SigningResult(fullPath, true, false, null, findings);
        } finally {
            if (!string.IsNullOrWhiteSpace(stagingPath)) OfficeFileCommit.DeleteIfExists(stagingPath);
        }
    }

    /// <summary>Creates VBA profiles and throws if validated atomic signing does not complete.</summary>
    public static OfficeVbaSigningResult Sign(string filePath, IOfficeSecurityProvider securityProvider,
        string certificateThumbprint, OfficeVbaSigningOptions? options = null) {
        OfficeVbaSigningResult result = TrySign(filePath, securityProvider, certificateThumbprint, options);
        if (!result.Succeeded) throw new InvalidOperationException(string.Join(" ", result.Findings.Select(finding => finding.Message)));
        return result;
    }

    private static IReadOnlyList<string> BuildSignArguments(string filePath, string thumbprint,
        OfficeVbaSigningOptions options) {
        var arguments = new List<string> { "sign", "/sha1", thumbprint, "/s", StoreName(options.StoreName) };
        if (options.StoreLocation == System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine) arguments.Add("/sm");
        arguments.Add("/fd");
        arguments.Add("SHA256");
        if (options.TimestampAuthorityUrl != null) {
            arguments.Add("/tr");
            arguments.Add(options.TimestampAuthorityUrl.AbsoluteUri);
            arguments.Add("/td");
            arguments.Add("SHA256");
        }
        arguments.Add(filePath);
        return arguments;
    }

    private static string StoreName(System.Security.Cryptography.X509Certificates.StoreName value) =>
        value == System.Security.Cryptography.X509Certificates.StoreName.CertificateAuthority ? "CA" : value.ToString();

    private static ProcessResult RunTool(string executable, IReadOnlyList<string> arguments,
        OfficeVbaSigningOptions options) {
        var output = new StringBuilder();
        bool truncated = false;
        using var process = new Process {
            StartInfo = new ProcessStartInfo {
                FileName = executable,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };
        process.StartInfo.Arguments = string.Join(" ", arguments.Select(QuoteArgument));
        void Append(string? value) {
            if (string.IsNullOrEmpty(value) || truncated) return;
            string line = value!;
            int remaining = options.MaxToolOutputCharacters - output.Length;
            if (remaining <= 0) { truncated = true; return; }
            output.AppendLine(line.Length <= remaining ? line : line.Substring(0, remaining));
            if (line.Length > remaining) truncated = true;
        }
        process.OutputDataReceived += (_, eventArgs) => Append(eventArgs.Data);
        process.ErrorDataReceived += (_, eventArgs) => Append(eventArgs.Data);
        try {
            if (!process.Start()) return new ProcessResult(null, false, "The process did not start.");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            if (!process.WaitForExit(checked((int)options.ToolTimeout.TotalMilliseconds))) {
                try {
#if NETSTANDARD2_0 || NET472
                    process.Kill();
#else
                    process.Kill(entireProcessTree: true);
#endif
                } catch { }
                return new ProcessResult(null, true, output.ToString());
            }
            process.WaitForExit();
            return new ProcessResult(process.ExitCode, false,
                output + (truncated ? Environment.NewLine + "[output truncated]" : string.Empty));
        } catch (Exception exception) when (exception is InvalidOperationException or System.ComponentModel.Win32Exception) {
            return new ProcessResult(null, false, exception.Message);
        }
    }

    private static bool TryNormalizeThumbprint(string? value, out string thumbprint, out string detail) {
        thumbprint = string.Empty;
        if (string.IsNullOrWhiteSpace(value)) { detail = "A certificate SHA-1 thumbprint is required."; return false; }
        var characters = new List<char>();
        foreach (char character in value!) {
            if (Uri.IsHexDigit(character)) characters.Add(char.ToUpperInvariant(character));
            else if (!(char.IsWhiteSpace(character) || character == ':' || character == '-')) {
                detail = "The certificate thumbprint contains an invalid character.";
                return false;
            }
        }
        if (characters.Count != 40) { detail = "SignTool requires a 40-character SHA-1 thumbprint."; return false; }
        thumbprint = new string(characters.ToArray());
        detail = string.Empty;
        return true;
    }

    private static string QuoteArgument(string value) {
        if (value.Length > 0 && value.All(character => !char.IsWhiteSpace(character) && character != '"')) return value;
        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    private static OfficeVbaSignatureFinding ToolFinding(string code, string tool, ProcessResult result,
        OfficeVbaSignatureProfile? profile = null) => Finding(code,
        result.TimedOut ? OfficePackageSignatureValidationState.Unsupported : OfficePackageSignatureValidationState.Failed,
        tool + (result.TimedOut ? " exceeded its timeout. " : " failed. ") + result.Output, profile);

    private static OfficeVbaSigningResult SigningResult(string path, bool supported, bool succeeded,
        OfficeVbaSignatureValidationResult? validation, IReadOnlyList<OfficeVbaSignatureFinding> findings) =>
        new(path, supported, succeeded, validation, findings.ToArray());

    private readonly struct ProcessResult {
        internal ProcessResult(int? exitCode, bool timedOut, string output) {
            ExitCode = exitCode; TimedOut = timedOut; Output = output;
        }
        internal int? ExitCode { get; }
        internal bool TimedOut { get; }
        internal string Output { get; }
        internal bool Succeeded => ExitCode == 0 && !TimedOut;
    }
}
