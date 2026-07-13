using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>Builds an immutable, ordered document processor pipeline.</summary>
public sealed class OfficeDocumentProcessorPipelineBuilder {
    private readonly List<IOfficeDocumentProcessor> _processors = new List<IOfficeDocumentProcessor>();
    private readonly HashSet<string> _ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Number of processors currently configured.</summary>
    public int Count => _processors.Count;

    /// <summary>Adds a processor at the end of the deterministic execution order.</summary>
    public OfficeDocumentProcessorPipelineBuilder Add(IOfficeDocumentProcessor processor) {
        if (processor == null) throw new ArgumentNullException(nameof(processor));
        string id = OfficeDocumentProcessorPipeline.ValidateProcessorId(processor.Id, nameof(processor));
        if (!_ids.Add(id)) {
            throw new InvalidOperationException($"A document processor with id '{id}' is already registered.");
        }
        if (!(processor is ISynchronousOfficeDocumentProcessor) && !(processor is IAsyncOfficeDocumentProcessor)) {
            _ids.Remove(id);
            throw new ArgumentException(
                "Processor must implement the synchronous or asynchronous processing contract.",
                nameof(processor));
        }
        _processors.Add(processor);
        return this;
    }

    /// <summary>Adds processors in enumeration order.</summary>
    public OfficeDocumentProcessorPipelineBuilder AddRange(IEnumerable<IOfficeDocumentProcessor> processors) {
        if (processors == null) throw new ArgumentNullException(nameof(processors));
        foreach (IOfficeDocumentProcessor processor in processors) Add(processor);
        return this;
    }

    /// <summary>Removes a processor by stable identifier.</summary>
    public bool Remove(string processorId) {
        if (string.IsNullOrWhiteSpace(processorId)) return false;
        for (int index = 0; index < _processors.Count; index++) {
            if (!string.Equals(_processors[index].Id, processorId, StringComparison.OrdinalIgnoreCase)) continue;
            _processors.RemoveAt(index);
            _ids.Remove(processorId);
            return true;
        }
        return false;
    }

    /// <summary>Removes every configured processor.</summary>
    public OfficeDocumentProcessorPipelineBuilder Clear() {
        _processors.Clear();
        _ids.Clear();
        return this;
    }

    /// <summary>Creates an immutable pipeline snapshot.</summary>
    public OfficeDocumentProcessorPipeline Build() {
        return _processors.Count == 0
            ? OfficeDocumentProcessorPipeline.Empty
            : new OfficeDocumentProcessorPipeline(_processors.ToArray());
    }
}
