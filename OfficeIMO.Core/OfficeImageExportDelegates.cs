using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>Consumes one result from a streaming batch export.</summary>
public delegate void OfficeImageExportConsumer(OfficeImageExportResult result);

/// <summary>Asynchronously consumes one result from a streaming batch export.</summary>
public delegate Task OfficeImageExportAsyncConsumer(
    OfficeImageExportResult result,
    CancellationToken cancellationToken);
