using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Security;

/// <summary>Fail-closed structural binding between XML DSig SignedInfo references and signed payload elements.</summary>
internal static class OfficeXmlSignatureBinding {
    private static readonly XNamespace Ds = XmlDigitalSignatureAlgorithms.Namespace;

    internal static AuthenticatedContent Resolve(
        XElement signature,
        XName payloadName,
        int maxReferences,
        bool requirePayload = true) {
        if (signature.Name != Ds + "Signature") {
            throw new InvalidDataException("The XML signature does not have a ds:Signature root element.");
        }

        XElement[] signedInfos = signature.Elements(Ds + "SignedInfo").Take(2).ToArray();
        if (signedInfos.Length != 1) {
            throw new InvalidDataException("The XML signature must contain exactly one direct ds:SignedInfo element.");
        }

        XElement[] signedInfoReferences = signedInfos[0].Elements(Ds + "Reference")
            .Take(maxReferences + 1).ToArray();
        if (signedInfoReferences.Length == 0) {
            throw new InvalidDataException("The XML signature does not authenticate any SignedInfo references.");
        }
        if (signedInfoReferences.Length > maxReferences) {
            throw new InvalidDataException("The XML signature exceeds the configured authenticated-reference limit.");
        }

        Dictionary<string, XElement?> elementsById = BuildUniqueElementIdIndex(signature);
        var authenticatedPayloads = new HashSet<XElement>();
        int authenticatedReferenceCount = signedInfoReferences.Length;
        foreach (XElement reference in signedInfoReferences) {
            string uri = ((string?)reference.Attribute("URI"))?.Trim() ?? string.Empty;
            if (!uri.StartsWith("#", StringComparison.Ordinal) || uri.Length == 1) continue;
            string id = uri.Substring(1);
            if (!elementsById.TryGetValue(id, out XElement? target) || target == null) {
                throw new InvalidDataException("A SignedInfo fragment reference does not resolve to one unique XML element.");
            }
            XElement[] payloads = target.DescendantsAndSelf(payloadName).ToArray();
            if (payloads.Length == 0) continue;
            if (!PreservesCompleteSubtree(reference, target)) {
                throw new InvalidDataException("A SignedInfo transform does not preserve the complete signed payload subtree.");
            }
            foreach (XElement payload in payloads) {
                if (!authenticatedPayloads.Add(payload)) continue;
                authenticatedReferenceCount = checked(authenticatedReferenceCount + payload.Elements(Ds + "Reference").Count());
                if (authenticatedReferenceCount > maxReferences) {
                    throw new InvalidDataException("The XML signature exceeds the configured authenticated-reference limit.");
                }
            }
        }

        XElement[] allPayloads = signature.Descendants(payloadName).ToArray();
        if (allPayloads.Any(payload => !authenticatedPayloads.Contains(payload))) {
            throw new InvalidDataException("The XML signature contains a payload that is not authenticated by SignedInfo.");
        }
        if (authenticatedPayloads.Count == 0 && requirePayload) {
            throw new InvalidDataException("SignedInfo does not authenticate the required signed payload.");
        }

        return new AuthenticatedContent(signedInfoReferences, authenticatedPayloads.ToArray());
    }

    private static Dictionary<string, XElement?> BuildUniqueElementIdIndex(XElement signature) {
        var elementsById = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        foreach (XElement element in signature.DescendantsAndSelf()) {
            foreach (XAttribute attribute in element.Attributes().Where(attribute =>
                         attribute.Name.Namespace == XNamespace.None &&
                         attribute.Name.LocalName is "Id" or "ID" or "id")) {
                string id = attribute.Value;
                if (string.IsNullOrWhiteSpace(id)) {
                    throw new InvalidDataException("The XML signature contains an empty identifier.");
                }
                if (!elementsById.TryGetValue(id, out XElement? existing)) {
                    elementsById[id] = element;
                } else if (!ReferenceEquals(existing, element)) {
                    elementsById[id] = null;
                }
            }
        }
        if (elementsById.Values.Any(element => element == null)) {
            throw new InvalidDataException("The XML signature contains duplicate identifiers.");
        }
        return elementsById;
    }

    private static bool PreservesCompleteSubtree(XElement reference, XElement target) {
        XElement? transforms = reference.Element(Ds + "Transforms");
        if (transforms == null) return true;
        return transforms.Elements(Ds + "Transform").All(transform => {
            string? algorithm = ((string?)transform.Attribute("Algorithm"))?.Trim();
            if (algorithm == "http://www.w3.org/2000/09/xmldsig#enveloped-signature") {
                return !target.DescendantsAndSelf(Ds + "Signature").Any();
            }
            return algorithm == XmlDigitalSignatureAlgorithms.CanonicalXml ||
                   algorithm == XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments ||
                   algorithm == "http://www.w3.org/2001/10/xml-exc-c14n#" ||
                   algorithm == "http://www.w3.org/2001/10/xml-exc-c14n#WithComments";
        });
    }

    internal readonly struct AuthenticatedContent {
        internal AuthenticatedContent(
            IReadOnlyList<XElement> signedInfoReferences,
            IReadOnlyList<XElement> payloads) {
            SignedInfoReferences = signedInfoReferences;
            Payloads = payloads;
        }

        internal IReadOnlyList<XElement> SignedInfoReferences { get; }
        internal IReadOnlyList<XElement> Payloads { get; }
    }
}
