using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Named collection of OfficeIMO-native stencil shape definitions.
    /// </summary>
    public sealed class VisioStencilCatalog {
        private readonly Dictionary<string, VisioStencilShape> _lookup = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Initializes a new stencil catalog.
        /// </summary>
        public VisioStencilCatalog(string name, IEnumerable<VisioStencilShape> shapes) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Catalog name cannot be null or whitespace.", nameof(name));
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));

            Name = name;
            Shapes = shapes.ToList().AsReadOnly();
            Categories = Shapes
                .Select(shape => shape.Category)
                .Where(category => !string.IsNullOrWhiteSpace(category))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            foreach (VisioStencilShape shape in Shapes) {
                AddLookup(shape.Id, shape);
                AddLookup(shape.Name, shape);
                AddLookup(shape.MasterNameU, shape);
                foreach (string keyword in shape.Keywords) {
                    AddLookup(keyword, shape);
                }

                foreach (string alias in shape.Aliases) {
                    AddLookup(alias, shape);
                }

                foreach (string tag in shape.Tags) {
                    AddLookup(tag, shape);
                }
            }
        }

        /// <summary>
        /// Creates a stencil catalog using the fluent catalog builder.
        /// </summary>
        public static VisioStencilCatalog Create(string name, Action<VisioStencilCatalogBuilder> configure) {
            if (configure == null) throw new ArgumentNullException(nameof(configure));

            VisioStencilCatalogBuilder builder = new(name);
            configure(builder);
            return builder.Build();
        }

        /// <summary>
        /// Loads an OfficeIMO-native stencil catalog manifest.
        /// </summary>
        public static VisioStencilCatalog Load(string path) {
            return VisioStencilCatalogManifest.Load(path);
        }

        /// <summary>
        /// Loads an OfficeIMO-native stencil catalog manifest.
        /// </summary>
        public static VisioStencilCatalog Load(Stream stream) {
            return VisioStencilCatalogManifest.Load(stream);
        }

        /// <summary>
        /// Gets the catalog name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets shapes in this catalog.
        /// </summary>
        public IReadOnlyList<VisioStencilShape> Shapes { get; }

        /// <summary>
        /// Gets category names represented by this catalog.
        /// </summary>
        public IReadOnlyList<string> Categories { get; }

        /// <summary>
        /// Saves this catalog as an OfficeIMO-native stencil catalog manifest.
        /// </summary>
        public void Save(string path) {
            VisioStencilCatalogManifest.Save(this, path);
        }

        /// <summary>
        /// Saves this catalog as an OfficeIMO-native stencil catalog manifest.
        /// </summary>
        public void Save(Stream stream) {
            VisioStencilCatalogManifest.Save(this, stream);
        }

        /// <summary>
        /// Attempts to find a stencil shape by id, name, master name, keyword, alias, or tag.
        /// </summary>
        public bool TryGet(string idOrName, out VisioStencilShape? shape) {
            if (string.IsNullOrWhiteSpace(idOrName)) {
                shape = null;
                return false;
            }

            return _lookup.TryGetValue(idOrName, out shape);
        }

        /// <summary>
        /// Gets a stencil shape by id, name, master name, keyword, alias, or tag.
        /// </summary>
        public VisioStencilShape Get(string idOrName) {
            if (TryGet(idOrName, out VisioStencilShape? shape) && shape != null) {
                return shape;
            }

            throw new KeyNotFoundException($"Stencil shape '{idOrName}' was not found in catalog '{Name}'.");
        }

        /// <summary>
        /// Attempts to find the first matching stencil shape from a prioritized set of lookup or search queries.
        /// Exact id/name/master/keyword/alias/tag matches are preferred before search matches for each query.
        /// </summary>
        /// <param name="queries">Prioritized lookup or search queries.</param>
        /// <param name="shape">The matched stencil shape, when one is found.</param>
        public bool TryFindBest(IEnumerable<string> queries, out VisioStencilShape? shape) {
            if (queries == null) throw new ArgumentNullException(nameof(queries));

            foreach (string query in queries.Where(value => !string.IsNullOrWhiteSpace(value))) {
                if (TryGet(query, out shape) && shape != null) {
                    return true;
                }

                shape = Search(query).FirstOrDefault();
                if (shape != null) {
                    return true;
                }
            }

            shape = null;
            return false;
        }

        /// <summary>
        /// Finds the first matching stencil shape from a prioritized set of lookup or search queries.
        /// Exact id/name/master/keyword/alias/tag matches are preferred before search matches for each query.
        /// </summary>
        /// <param name="queries">Prioritized lookup or search queries.</param>
        public VisioStencilShape FindBest(params string[] queries) {
            if (queries == null) throw new ArgumentNullException(nameof(queries));
            if (TryFindBest(queries, out VisioStencilShape? shape) && shape != null) {
                return shape;
            }

            string attempted = string.Join(", ", queries.Where(value => !string.IsNullOrWhiteSpace(value)));
            throw new KeyNotFoundException($"No stencil shape matching '{attempted}' was found in catalog '{Name}'.");
        }

        /// <summary>
        /// Searches stencil shapes by id, name, master name, category, keyword, alias, or tag.
        /// </summary>
        /// <param name="query">Search text.</param>
        public IReadOnlyList<VisioStencilShape> Search(string query) {
            if (string.IsNullOrWhiteSpace(query)) {
                return Shapes;
            }

            string normalizedQuery = query.Trim();
            return Shapes
                .Select(shape => new { Shape = shape, Score = Score(shape, normalizedQuery) })
                .Where(match => match.Score > 0)
                .OrderByDescending(match => match.Score)
                .ThenBy(match => match.Shape.Category, StringComparer.OrdinalIgnoreCase)
                .ThenBy(match => match.Shape.Name, StringComparer.OrdinalIgnoreCase)
                .Select(match => match.Shape)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Gets shapes in a category.
        /// </summary>
        /// <param name="category">Category name.</param>
        public IReadOnlyList<VisioStencilShape> InCategory(string category) {
            if (string.IsNullOrWhiteSpace(category)) {
                return Array.Empty<VisioStencilShape>();
            }

            return Shapes
                .Where(shape => string.Equals(shape.Category, category, StringComparison.OrdinalIgnoreCase))
                .OrderBy(shape => shape.Name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private void AddLookup(string key, VisioStencilShape shape) {
            if (!_lookup.ContainsKey(key)) {
                _lookup.Add(key, shape);
            }
        }

        private static int Score(VisioStencilShape shape, string query) {
            int best = 0;
            best = Math.Max(best, ScoreText(shape.Id, query, 100, 70));
            best = Math.Max(best, ScoreText(shape.Name, query, 95, 65));
            best = Math.Max(best, ScoreText(shape.MasterNameU, query, 90, 60));
            best = Math.Max(best, ScoreText(shape.Category, query, 80, 50));
            best = Math.Max(best, ScoreMany(shape.Aliases, query, 88, 58));
            best = Math.Max(best, ScoreMany(shape.Tags, query, 86, 56));
            best = Math.Max(best, ScoreMany(shape.Keywords, query, 84, 54));
            return best;
        }

        private static int ScoreMany(IEnumerable<string> values, string query, int exactScore, int containsScore) {
            int best = 0;
            foreach (string value in values) {
                best = Math.Max(best, ScoreText(value, query, exactScore, containsScore));
            }

            return best;
        }

        private static int ScoreText(string value, string query, int exactScore, int containsScore) {
            if (string.Equals(value, query, StringComparison.OrdinalIgnoreCase)) {
                return exactScore;
            }

            return value.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ? containsScore : 0;
        }
    }
}
