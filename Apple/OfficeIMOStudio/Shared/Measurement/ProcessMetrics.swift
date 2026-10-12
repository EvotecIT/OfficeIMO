import Darwin
import Foundation

/// Process facts used by the feasibility measurements.
enum ProcessMetrics {
    struct Footprint: Codable, Sendable {
        /// The memory the system charges to the process (what iOS memory limits apply to), in bytes.
        let current: UInt64
        /// The highest footprint since launch, in bytes.
        let peak: UInt64
    }

    static func footprint() -> Footprint {
        var info = task_vm_info_data_t()
        var count = mach_msg_type_number_t(MemoryLayout<task_vm_info_data_t>.size / MemoryLayout<natural_t>.size)
        let result = withUnsafeMutablePointer(to: &info) {
            $0.withMemoryRebound(to: integer_t.self, capacity: Int(count)) {
                task_info(mach_task_self_, task_flavor_t(TASK_VM_INFO), $0, &count)
            }
        }
        guard result == KERN_SUCCESS else { return Footprint(current: 0, peak: 0) }
        return Footprint(current: info.phys_footprint, peak: UInt64(info.ledger_phys_footprint_peak))
    }

    /// When the kernel started this process, for launch timing.
    static func processStart() -> Date? {
        var info = kinfo_proc()
        var size = MemoryLayout<kinfo_proc>.stride
        var name: [Int32] = [CTL_KERN, KERN_PROC, KERN_PROC_PID, getpid()]
        guard sysctl(&name, u_int(name.count), &info, &size, nil, 0) == 0 else { return nil }
        let start = info.kp_proc.p_un.__p_starttime
        return Date(timeIntervalSince1970: TimeInterval(start.tv_sec) + TimeInterval(start.tv_usec) / 1_000_000)
    }

    static var deviceModel: String {
        #if os(macOS)
        return "\(systemString("hw.model")) (\(systemString("hw.machine")))"
        #elseif targetEnvironment(simulator)
        return "Simulator: \(ProcessInfo.processInfo.environment["SIMULATOR_MODEL_IDENTIFIER"] ?? "?") on \(systemString("hw.machine"))"
        #else
        return systemString("hw.machine")
        #endif
    }

    private static func systemString(_ name: String) -> String {
        var size = 0
        sysctlbyname(name, nil, &size, nil, 0)
        var bytes = [UInt8](repeating: 0, count: size)
        sysctlbyname(name, &bytes, &size, nil, 0)
        return String(decoding: bytes.prefix { $0 != 0 }, as: UTF8.self)
    }
}

/// Elapsed milliseconds of a block, measured with the monotonic clock.
func measure<T>(isolation: isolated (any Actor)? = #isolation, _ work: () async throws -> T) async rethrows -> (T, Double) {
    let start = ContinuousClock.now
    let value = try await work()
    return (value, milliseconds(since: start))
}

func milliseconds(since start: ContinuousClock.Instant) -> Double {
    let elapsed = ContinuousClock.now - start
    return Double(elapsed.components.seconds) * 1000 + Double(elapsed.components.attoseconds) / 1e15
}
