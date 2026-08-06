#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeIMO.Security;

internal static class XmlDigitalSignatureReferenceWorkCalculator {
    private static readonly string[] LocalReferenceIdAttributeNames = { "Id", "ID", "id" };

    internal static long Measure(
        XmlDocument document,
        XmlElement? signedInfo,
        int certificateCandidateCount,
        long maxDigestWorkBytes) {
        if (certificateCandidateCount <= 0 || signedInfo == null) return 0L;

        Dictionary<string, XmlElement?> targetsById = IndexLocalReferenceTargets(document);
        var targetSizes = new Dictionary<string, long>(StringComparer.Ordinal);
        long totalDigestWorkBytes = 0L;
        foreach (XmlElement reference in signedInfo.ChildNodes
                     .OfType<XmlElement>()
                     .Where(element =>
                         element.LocalName == "Reference" &&
                         element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)) {
            string uri = reference.GetAttribute("URI");
            if (uri.Length > 0 && uri[0] != '#') continue;
            if (!targetSizes.TryGetValue(uri, out long targetBytes)) {
                XmlElement? target = ResolveLocalReferenceTarget(document, targetsById, uri);
                targetBytes = target == null ? 0L : Encoding.UTF8.GetByteCount(target.OuterXml);
                targetSizes.Add(uri, targetBytes);
            }
            long transformPasses = reference.ChildNodes
                .OfType<XmlElement>()
                .Where(element =>
                    element.LocalName == "Transforms" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                .SelectMany(element => element.ChildNodes.OfType<XmlElement>())
                .LongCount(element =>
                    element.LocalName == "Transform" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace) + 1L;
            long referenceWorkBytes = CheckedMultiply(targetBytes, transformPasses, maxDigestWorkBytes);
            long candidateWorkBytes = CheckedMultiply(referenceWorkBytes, certificateCandidateCount, maxDigestWorkBytes);
            if (candidateWorkBytes > maxDigestWorkBytes - totalDigestWorkBytes) {
                throw CreateLimitException(maxDigestWorkBytes);
            }
            totalDigestWorkBytes += candidateWorkBytes;
        }
        return totalDigestWorkBytes;
    }

    private static long CheckedMultiply(long value, long multiplier, long limit) {
        if (value > 0 && multiplier > long.MaxValue / value) throw CreateLimitException(limit);
        return value * multiplier;
    }

    private static InvalidDataException CreateLimitException(long limit) =>
        new("Local SignedInfo references exceed the " + limit +
            " byte remaining aggregate digest-work limit across signature parts and certificate candidates.");

    private static Dictionary<string, XmlElement?> IndexLocalReferenceTargets(XmlDocument document) {
        var targets = new Dictionary<string, XmlElement?>(StringComparer.Ordinal);
        if (document.DocumentElement == null) return targets;
        foreach (XmlElement element in EnumerateElements(document.DocumentElement)) {
            foreach (string attributeName in LocalReferenceIdAttributeNames) {
                if (!element.HasAttribute(attributeName)) continue;
                string id = element.GetAttribute(attributeName);
                if (id.Length == 0) continue;
                if (targets.TryGetValue(id, out XmlElement? existing)) {
                    if (!ReferenceEquals(existing, element)) targets[id] = null;
                } else {
                    targets.Add(id, element);
                }
            }
        }
        return targets;
    }

    private static IEnumerable<XmlElement> EnumerateElements(XmlElement root) {
        yield return root;
        foreach (XmlElement descendant in root.GetElementsByTagName("*").OfType<XmlElement>()) {
            yield return descendant;
        }
    }

    private static XmlElement? ResolveLocalReferenceTarget(
        XmlDocument document,
        IReadOnlyDictionary<string, XmlElement?> targetsById,
        string uri) {
        if (uri.Length == 0) return document.DocumentElement;
        if (uri.Length == 1) return null;
        return targetsById.TryGetValue(uri.Substring(1), out XmlElement? target) ? target : null;
    }
}
