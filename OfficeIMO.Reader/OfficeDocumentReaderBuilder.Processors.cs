using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReaderBuilder {
    private readonly OfficeDocumentProcessorPipelineBuilder _processorPipelineBuilder = new OfficeDocumentProcessorPipelineBuilder();
    private OfficeDocumentProcessingOptions _processingOptions = new OfficeDocumentProcessingOptions();

    /// <summary>Adds a document processor after existing processors.</summary>
    public OfficeDocumentReaderBuilder AddProcessor(IOfficeDocumentProcessor processor) {
        _processorPipelineBuilder.Add(processor);
        return this;
    }

    /// <summary>Adds document processors in enumeration order.</summary>
    public OfficeDocumentReaderBuilder AddProcessors(IEnumerable<IOfficeDocumentProcessor> processors) {
        _processorPipelineBuilder.AddRange(processors);
        return this;
    }

    /// <summary>Replaces configured processors with a snapshot of an existing pipeline.</summary>
    public OfficeDocumentReaderBuilder UseProcessorPipeline(OfficeDocumentProcessorPipeline pipeline) {
        if (pipeline == null) throw new ArgumentNullException(nameof(pipeline));
        _processorPipelineBuilder.Clear().AddRange(pipeline.Processors);
        return this;
    }

    /// <summary>Removes a configured processor by stable identifier.</summary>
    public bool RemoveProcessor(string processorId) {
        return _processorPipelineBuilder.Remove(processorId);
    }

    /// <summary>Removes all configured processors.</summary>
    public OfficeDocumentReaderBuilder ClearProcessors() {
        _processorPipelineBuilder.Clear();
        return this;
    }

    /// <summary>Sets processor failure behavior for the built reader.</summary>
    public OfficeDocumentReaderBuilder WithProcessorFailureBehavior(
        OfficeDocumentProcessorFailureBehavior failureBehavior) {
        if (!Enum.IsDefined(typeof(OfficeDocumentProcessorFailureBehavior), failureBehavior)) {
            throw new ArgumentOutOfRangeException(nameof(failureBehavior));
        }
        _processingOptions = new OfficeDocumentProcessingOptions { FailureBehavior = failureBehavior };
        return this;
    }
}
