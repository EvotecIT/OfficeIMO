using System;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    private bool _xmlChanged;
    private bool _cachedXmlDifference;
    private bool _hadDeclaration;
    private string? _declarationVersion;
    private string? _declarationEncoding;
    private string? _declarationStandalone;

    private bool HasChanges {
        get {
            if (_modified) return true;
            XDeclaration? declaration = _xml.Declaration;
            // XDeclaration is mutable but does not raise XObject.Changed.
            bool declarationChanged = _hadDeclaration != (declaration != null) ||
                _declarationVersion != declaration?.Version || _declarationEncoding != declaration?.Encoding ||
                _declarationStandalone != declaration?.Standalone;
            if (_xmlChanged || declarationChanged) {
                _cachedXmlDifference = !string.Equals(_originalXmlFingerprint, GetXmlFingerprint(_xml), StringComparison.Ordinal);
                _xmlChanged = false;
                CaptureDeclaration();
            }
            return _cachedXmlDifference;
        }
    }

    private void CaptureDeclaration() {
        _hadDeclaration = _xml.Declaration != null;
        _declarationVersion = _xml.Declaration?.Version;
        _declarationEncoding = _xml.Declaration?.Encoding;
        _declarationStandalone = _xml.Declaration?.Standalone;
    }
}
