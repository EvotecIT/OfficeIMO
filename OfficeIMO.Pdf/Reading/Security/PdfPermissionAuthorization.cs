namespace OfficeIMO.Pdf;

internal static class PdfPermissionAuthorization {
    internal static bool CanExtractText(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy) =>
        IsAllowed(security, policy, PdfStandardPermissions.CopyContents) ||
        IsAllowed(security, policy, PdfStandardPermissions.Accessibility);

    internal static bool CanExtractContent(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy) =>
        IsAllowed(security, policy, PdfStandardPermissions.CopyContents);

    internal static bool CanMutate(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy, PdfMutationOperation operation) {
        if (!security.HasEncryption || security.HasOwnerAuthorization) {
            return true;
        }

        if (operation == PdfMutationOperation.ChangeEncryption) {
            return false;
        }

        if (policy == PdfPermissionPolicy.IgnoreRestrictions) {
            return true;
        }

        switch (operation) {
            case PdfMutationOperation.ExtractPages:
            case PdfMutationOperation.MergeDocuments:
                return security.AllowsCopying == true && security.AllowsDocumentAssembly == true;
            case PdfMutationOperation.ModifyPageTree:
                return security.AllowsDocumentAssembly == true;
            case PdfMutationOperation.ModifyAnnotations:
                return security.AllowsAnnotationChanges == true || security.AllowsModification == true;
            case PdfMutationOperation.FillFormFields:
            case PdfMutationOperation.FlattenFormFields:
            case PdfMutationOperation.FillAndFlattenFormFields:
            case PdfMutationOperation.ModifyAcroForm:
                return security.AllowsFormFilling == true || security.AllowsModification == true;
            default:
                return security.AllowsModification == true;
        }
    }

    internal static bool CanRewriteFormFields(
        PdfDocumentSecurityInfo security,
        PdfPermissionPolicy policy,
        PdfMutationOperation operation) =>
        CanMutate(security, policy, operation) &&
        (!security.HasEncryption ||
         security.HasOwnerAuthorization ||
         policy == PdfPermissionPolicy.IgnoreRestrictions);

    internal static void DemandTextExtraction(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy) {
        if (CanExtractText(security, policy)) {
            return;
        }

        throw new PdfPermissionDeniedException(
            PdfStandardPermissions.CopyContents,
            security.PasswordAuthenticationRole,
            "PDF text extraction is restricted by the authenticated user-password permissions. Supply owner authorization or set PermissionPolicy to IgnoreRestrictions after confirming that the operation is authorized.");
    }

    internal static void DemandContentExtraction(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy, string contentName) {
        if (CanExtractContent(security, policy)) {
            return;
        }

        throw new PdfPermissionDeniedException(
            PdfStandardPermissions.CopyContents,
            security.PasswordAuthenticationRole,
            "PDF " + contentName + " extraction is restricted by the authenticated user-password permissions. Supply owner authorization or set PermissionPolicy to IgnoreRestrictions after confirming that the operation is authorized.");
    }

    internal static bool RestrictionsIgnored(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy) =>
        security.HasEncryption &&
        !security.HasOwnerAuthorization &&
        policy == PdfPermissionPolicy.IgnoreRestrictions;

    private static bool IsAllowed(PdfDocumentSecurityInfo security, PdfPermissionPolicy policy, PdfStandardPermissions permission) {
        if (!security.HasEncryption || security.HasOwnerAuthorization || policy == PdfPermissionPolicy.IgnoreRestrictions) {
            return true;
        }

        PdfStandardPermissions? allowed = security.AllowedStandardPermissions;
        return allowed.HasValue && (allowed.Value & permission) == permission;
    }
}
