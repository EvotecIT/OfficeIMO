using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Explicitly registers schemes, recognizers and keywords for an immutable catalog.</summary>
public sealed class EntityCatalogBuilder {
    private readonly List<EntityScheme> _schemes = new();
    private readonly Dictionary<string, EntityScheme> _byId = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<IEntityRecognizer> _recognizers = new();
    private readonly HashSet<string> _recognizerIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<EntityKeyword> _keywords = new();

    /// <summary>Registers a scheme for validation and optionally registers a recognizer.</summary>
    public EntityCatalogBuilder Add(EntityScheme scheme, IEntityRecognizer? recognizer = null) {
        if (scheme == null) throw new ArgumentNullException(nameof(scheme));
        VerifyScheme(scheme);
        if (recognizer != null) VerifyRecognizer(recognizer, scheme);
        if (!_byId.ContainsKey(scheme.Id)) {
            _schemes.Add(scheme);
            _byId.Add(scheme.Id, scheme);
        }
        if (recognizer != null) RegisterRecognizer(recognizer);
        return this;
    }

    /// <summary>Registers a recognizer for schemes already present in the builder.</summary>
    public EntityCatalogBuilder AddRecognizer(IEntityRecognizer recognizer) {
        if (recognizer == null) throw new ArgumentNullException(nameof(recognizer));
        VerifyRecognizer(recognizer, pending: null);
        RegisterRecognizer(recognizer);
        return this;
    }

    /// <summary>Registers a whole-word context keyword and its BCP 47 language tag.</summary>
    public EntityCatalogBuilder AddKeyword(string schemeId, string language, string keyword) {
        if (schemeId == null) throw new ArgumentNullException(nameof(schemeId));
        if (!_byId.ContainsKey(schemeId)) throw new ArgumentException("The entity scheme is not registered.", nameof(schemeId));
        if (string.IsNullOrWhiteSpace(language)) throw new ArgumentException("A keyword language is required.", nameof(language));
        if (string.IsNullOrWhiteSpace(keyword)) throw new ArgumentException("A non-empty keyword is required.", nameof(keyword));
        _keywords.Add(new EntityKeyword(schemeId, language, keyword));
        return this;
    }

    /// <summary>Creates a snapshot that can be shared across threads.</summary>
    public EntityCatalog Build() => new(_schemes.ToArray(), _recognizers.ToArray(), _keywords.ToArray());

    private void VerifyScheme(EntityScheme scheme) {
        if (string.IsNullOrWhiteSpace(scheme.Id)) throw new ArgumentException("A scheme ID is required.", nameof(scheme));
        if (_byId.TryGetValue(scheme.Id, out EntityScheme? existing) && !ReferenceEquals(existing, scheme))
            throw new ArgumentException("A different scheme already uses this ID.", nameof(scheme));
    }

    private void VerifyRecognizer(IEntityRecognizer recognizer, EntityScheme? pending) {
        if (string.IsNullOrWhiteSpace(recognizer.Id)) throw new ArgumentException("A recognizer ID is required.", nameof(recognizer));
        if (_recognizerIds.Contains(recognizer.Id)) throw new ArgumentException("The recognizer ID is already registered.", nameof(recognizer));
        foreach (EntityScheme scheme in recognizer.Schemes) {
            if (!ReferenceEquals(pending, scheme) && (!_byId.TryGetValue(scheme.Id, out EntityScheme? found) || !ReferenceEquals(found, scheme)))
                throw new ArgumentException("Register the recognizer's schemes before registering the recognizer.", nameof(recognizer));
        }
    }

    private void RegisterRecognizer(IEntityRecognizer recognizer) {
        _recognizerIds.Add(recognizer.Id);
        _recognizers.Add(recognizer);
    }
}
