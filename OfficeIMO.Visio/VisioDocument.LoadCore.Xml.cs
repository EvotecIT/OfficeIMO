using System;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        internal const long MaxPackageXmlPartBytes = 32_000_000;
        internal const long MaxPackageXmlCharacters = 30_000_000;

        private static XDocument LoadPackageXml(PackagePart part, string description, CancellationToken cancellationToken = default,
            long maxCharacters = MaxPackageXmlCharacters) {
            if (maxCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maxCharacters));
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.GetStream();
            using Stream boundedStream = new BoundedReadStream(stream, MaxPackageXmlPartBytes, description);
            using XmlReader reader = XmlReader.Create(boundedStream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = Math.Min(maxCharacters, MaxPackageXmlCharacters),
                MaxCharactersFromEntities = 0
            });
            using var cancellable = new OfficeXmlLimitingReader(reader, description,
                int.MaxValue, int.MaxValue, int.MaxValue, cancellationToken);
            return XDocument.Load(cancellable, LoadOptions.None);
        }
    }
}
