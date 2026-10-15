import Foundation

/// Coalesces reads for one engine revision and retries when the caller's revision changes while waiting.
@MainActor
final class RevisionRead<Value: Sendable> {
    private struct Request {
        let source: ObjectIdentifier
        let revision: Int64
        let id: UUID
        let task: Task<Value, Error>
    }

    private final class Cached {
        weak var source: AnyObject?
        let revision: Int64
        let value: Value

        init(source: AnyObject, revision: Int64, value: Value) {
            self.source = source
            self.revision = revision
            self.value = value
        }
    }

    private var pending: Request?
    private var cached: Cached?
    var isLoading: Bool { pending != nil }

    func read(source: AnyObject, revision: @MainActor () -> Int64,
              operation: @escaping @Sendable () async throws -> Value) async throws -> Value {
        while !Task.isCancelled {
            let requested = revision()
            do {
                let value = try await value(source: source, revision: requested, operation: operation)
                try Task.checkCancellation()
                if requested == revision() { return value }
            } catch {
                if requested == revision() { throw error }
            }
        }
        throw CancellationError()
    }

    private func value(source: AnyObject, revision: Int64,
                       operation: @escaping @Sendable () async throws -> Value) async throws -> Value {
        let identity = ObjectIdentifier(source)
        if let cached, cached.source === source, cached.revision == revision { return cached.value }
        if let pending, pending.source == identity, pending.revision == revision {
            return try await pending.task.value
        }
        let id = UUID()
        let task = Task { try await operation() }
        pending = Request(source: identity, revision: revision, id: id, task: task)
        defer {
            // An older waiter must not clear the request a newer revision started.
            if pending?.id == id { pending = nil }
        }
        let value = try await task.value
        if pending?.id == id { cached = Cached(source: source, revision: revision, value: value) }
        return value
    }
}
