namespace OfficeIMO.Email;

/// <summary>Selects the MIME representation used for an S/MIME signature.</summary>
public enum EmailSmimeSignatureMode {
    /// <summary>Writes multipart/signed with readable MIME content and a detached CMS signature.</summary>
    ClearSigned,
    /// <summary>Writes application/pkcs7-mime with encapsulated CMS SignedData.</summary>
    OpaqueSigned
}
