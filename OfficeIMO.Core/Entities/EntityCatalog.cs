using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>An immutable registry of schemes, recognizers and their context keywords.</summary>
public sealed class EntityCatalog {
    private readonly Dictionary<string, EntityScheme> _byId = new(StringComparer.OrdinalIgnoreCase);

    internal EntityCatalog(EntityScheme[] schemes, IEntityRecognizer[] recognizers, EntityKeyword[] keywords) {
        Schemes = Array.AsReadOnly(schemes);
        Recognizers = Array.AsReadOnly(recognizers);
        Keywords = new Dictionary<string, List<EntityKeyword>>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < schemes.Length; index++) {
            _byId.Add(schemes[index].Id, schemes[index]);
            RegistrationOrder.Add(schemes[index].Id, index);
        }
        foreach (EntityKeyword keyword in keywords) {
            if (!Keywords.TryGetValue(keyword.SchemeId, out List<EntityKeyword>? entries)) {
                entries = new List<EntityKeyword>();
                Keywords.Add(keyword.SchemeId, entries);
            }
            entries.Add(keyword);
        }
    }

    /// <summary>Every rule-based scheme included in Core.</summary>
    public static EntityCatalog Default => DefaultCatalog.Instance;
    /// <summary>Registered schemes in registration order.</summary>
    public IReadOnlyList<EntityScheme> Schemes { get; }

    internal IReadOnlyList<IEntityRecognizer> Recognizers { get; }
    internal Dictionary<string, List<EntityKeyword>> Keywords { get; }
    internal Dictionary<string, int> RegistrationOrder { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Returns the scheme registered under the supplied ID.</summary>
    public EntityScheme GetScheme(string id) {
        if (TryGetScheme(id, out EntityScheme? scheme)) return scheme!;
        throw new KeyNotFoundException("The entity scheme is not registered.");
    }

    /// <summary>Attempts to locate a scheme using a case-insensitive ID.</summary>
    public bool TryGetScheme(string id, out EntityScheme? scheme) {
        if (id == null) throw new ArgumentNullException(nameof(id));
        return _byId.TryGetValue(id, out scheme);
    }

    /// <summary>Validates a value against a registered scheme.</summary>
    public EntityValidationResult Validate(string schemeId, string? value) => GetScheme(schemeId).Validate(value);

    private static class DefaultCatalog {
        internal static readonly EntityCatalog Instance = new EntityCatalogBuilder()
            .AddPoland().AddEuVat().AddUnitedKingdom().AddUnitedStates().AddFinancial().AddInternet().Build();
    }
}
