using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Runtime.CompilerServices;

namespace PSWriteOffice.Foundation;

/// <summary>Binds an optional SecureString or credential password; warns before accepting plain text.</summary>
/// <para>A collection binds element by element for per-source password parameters; null elements stay null.</para>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficePasswordTransformationAttribute : ArgumentTransformationAttribute {
    private static readonly ConditionalWeakTable<SecureString, CredentialName> CredentialNames = new();

    /// <summary>Preserves a credential's user name for parameters that bind an account and secret together.</summary>
    public bool PreserveUserName { get; set; }

    private const string PlainTextWarning = "A plain-text password was supplied. Use Read-Host -AsSecureString or a PSCredential to keep passwords out of scripts and history.";

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        // The command decides whether null means an omitted source password or an explicit output reset.
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value == null) return null!;
        bool warned = false;
        Action<string> warning = message => {
            if (warned) return;
            warned = true;
            engineIntrinsics.InvokeCommand.InvokeScript("param($message) Microsoft.PowerShell.Utility\\Write-Warning -Message $message", false,
                System.Management.Automation.Runspaces.PipelineResultTypes.None, null, message);
        };
        if (PreserveUserName && value is PSCredential credential) {
            // A private copy keeps two accounts using the same SecureString independent.
            var password = credential.Password.Copy();
            password.MakeReadOnly();
            CredentialNames.Add(password, new CredentialName(credential.UserName));
            return password;
        }
        return value is IEnumerable values && value is not string ? ConvertPasswords(values, warning) : ConvertPassword(value, warning);
    }

    /// <summary>Uses the bound credential's account, rejecting an explicit conflicting account.</summary>
    internal static string? ResolveUserName(SecureString password, string? explicitUserName) {
        if (!CredentialNames.TryGetValue(password, out CredentialName? credential)) return explicitUserName;
        if (explicitUserName != null && !string.Equals(explicitUserName, credential.Value, StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("The explicit user name must match the PSCredential user name.");
        return credential.Value;
    }

    private sealed class CredentialName {
        internal CredentialName(string value) => Value = value;
        internal string Value { get; }
    }

    internal static SecureString ConvertPassword(object inputData, Action<string> warning) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value is SecureString secure) return secure;
        if (value is PSCredential credential) return credential.Password;
        if (value is string text) {
            warning(PlainTextWarning);
            var result = new SecureString();
            foreach (char character in text) result.AppendChar(character);
            result.MakeReadOnly();
            return result;
        }
        throw OfficeArgumentValue.Invalid("Expected a SecureString, PSCredential or plain-text password.");
    }

    internal static SecureString?[] ConvertPasswords(IEnumerable values, Action<string> warning) {
        var result = new List<SecureString?>();
        foreach (object? item in values)
            result.Add(OfficeArgumentValue.Unwrap(item) == null ? null : ConvertPassword(item!, warning));
        return result.ToArray();
    }
}

internal static class OfficePasswordExtensions {
    /// <summary>Reveals a bound password only at the native API boundary that requires a string.</summary>
    internal static string? ToPlainText(this SecureString? value) =>
        value == null ? null : new NetworkCredential(string.Empty, value).Password;
}
