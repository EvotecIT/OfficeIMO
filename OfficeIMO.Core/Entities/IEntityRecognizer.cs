using System.Collections.Generic;

namespace OfficeIMO.Entities;

/// <summary>Submits readings for one or more schemes using shared tokens.</summary>
public interface IEntityRecognizer {
    /// <summary>Stable recognizer identifier.</summary>
    string Id { get; }
    /// <summary>Schemes this recognizer can submit.</summary>
    IReadOnlyList<EntityScheme> Schemes { get; }
    /// <summary>Token shapes used by this recognizer.</summary>
    EntityTokenShape Shapes { get; }
    /// <summary>Examines tokens and submits validated readings from the current segment.</summary>
    void Recognize(EntityRecognitionContext context);
}
