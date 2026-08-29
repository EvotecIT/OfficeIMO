using System;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Core.Internal {
    /// <summary>Enforces structural XML limits before a consumer materializes each element.</summary>
    internal sealed class OfficeXmlLimitingReader : XmlReader, IXmlLineInfo {
        private readonly XmlReader _inner;
        private readonly string _formatName;
        private readonly int _maxDepth;
        private readonly int _maxElements;
        private readonly int _maxAttributes;
        private readonly CancellationToken _cancellationToken;
        private int _elements;
        private long _attributes;

        internal OfficeXmlLimitingReader(
            XmlReader inner,
            string formatName,
            int maxDepth,
            int maxElements,
            int maxAttributes,
            CancellationToken cancellationToken) {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _formatName = formatName ?? throw new ArgumentNullException(nameof(formatName));
            _maxDepth = maxDepth;
            _maxElements = maxElements;
            _maxAttributes = maxAttributes;
            _cancellationToken = cancellationToken;
        }

        public override int AttributeCount => _inner.AttributeCount;
        public override string BaseURI => _inner.BaseURI;
        public override bool CanResolveEntity => _inner.CanResolveEntity;
        public override int Depth => _inner.Depth;
        public override bool EOF => _inner.EOF;
        public override bool HasValue => _inner.HasValue;
        public override bool IsEmptyElement => _inner.IsEmptyElement;
        public override string LocalName => _inner.LocalName;
        public override string NamespaceURI => _inner.NamespaceURI;
        public override XmlNameTable NameTable => _inner.NameTable;
        public override XmlNodeType NodeType => _inner.NodeType;
        public override string Prefix => _inner.Prefix;
        public override char QuoteChar => _inner.QuoteChar;
        public override ReadState ReadState => _inner.ReadState;
        public override string Value => _inner.Value;
        public override string XmlLang => _inner.XmlLang;
        public override XmlSpace XmlSpace => _inner.XmlSpace;

        public int LineNumber => (_inner as IXmlLineInfo)?.LineNumber ?? 0;
        public int LinePosition => (_inner as IXmlLineInfo)?.LinePosition ?? 0;
        public bool HasLineInfo() => (_inner as IXmlLineInfo)?.HasLineInfo() == true;

        public override string GetAttribute(int i) => _inner.GetAttribute(i);
        public override string? GetAttribute(string name) => _inner.GetAttribute(name);
        public override string? GetAttribute(string name, string? namespaceURI) => _inner.GetAttribute(name, namespaceURI);
        public override string? LookupNamespace(string prefix) => _inner.LookupNamespace(prefix);
        public override bool MoveToAttribute(string name) => _inner.MoveToAttribute(name);
        public override bool MoveToAttribute(string name, string? ns) => _inner.MoveToAttribute(name, ns);
        public override void MoveToAttribute(int i) => _inner.MoveToAttribute(i);
        public override bool MoveToElement() => _inner.MoveToElement();
        public override bool MoveToFirstAttribute() => _inner.MoveToFirstAttribute();
        public override bool MoveToNextAttribute() => _inner.MoveToNextAttribute();
        public override bool ReadAttributeValue() => _inner.ReadAttributeValue();
        public override void ResolveEntity() => _inner.ResolveEntity();

        public override bool Read() {
            _cancellationToken.ThrowIfCancellationRequested();
            bool result = _inner.Read();
            if (!result || _inner.NodeType != XmlNodeType.Element) return result;
            if (_inner.Depth > _maxDepth) throw Limit("MaxDepth");
            if (++_elements > _maxElements) throw Limit("MaxElements");
            _attributes += _inner.AttributeCount;
            if (_attributes > _maxAttributes) throw Limit("MaxAttributes");
            return true;
        }

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }

        private InvalidDataException Limit(string name) =>
            new InvalidDataException(_formatName + " input exceeds " + name + ".");
    }
}
