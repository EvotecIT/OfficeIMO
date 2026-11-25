#!/usr/bin/env ruby
# Writes the App Store screenshot sync configs for every listing locale from
# store-art.json, pointing each display type at the composed store art for that
# locale, and lists them in powerforge.release.json. Run with --check to verify
# that the committed files are current.
require 'json'

APPLE = File.expand_path('..', __dir__)
ROOT = File.expand_path('../../..', APPLE)
spec = JSON.parse(File.read(File.join(__dir__, 'store-art.json')))
locales = JSON.parse(File.read(File.join(__dir__, 'locales.json'))).fetch('locales').map { |l| l.fetch('store') }
abort('store-art.json and locales.json list different locales') unless spec.fetch('locales').sort == locales.sort

APP_ID = '6818096229'
STORE_ART = 'Artifacts/Studio/Apple/screenshots/store-art'
APPROVALS = 'Artifacts/Studio/Apple/screenshots'
# App Store display types are listing policy; rendering belongs to PowerForge.
DISPLAYS = {
  'iphone-69' => ['APP_IPHONE_67', 1320, 2868],
  'ipad-13' => ['APP_IPAD_PRO_3GEN_129', 2064, 2752],
  'mac' => ['APP_DESKTOP', 2880, 1800]
}.freeze
PLATFORM_SETS = { 'macos' => %w[mac], 'ios' => %w[iphone-69 ipad-13] }.freeze
MAXIMUM_SCREENSHOTS = 10

def frame_count(spec, device)
  count = spec.fetch('frames').count { |frame| frame.fetch('devices').include?(device) }
  abort("#{device} has #{count} frames; App Store Connect accepts at most #{MAXIMUM_SCREENSHOTS}") if count > MAXIMUM_SCREENSHOTS
  abort("#{device} has no frames") if count.zero?
  count
end

expected = {}
locales.each do |locale|
  PLATFORM_SETS.each do |platform, devices|
    platform_name = platform == 'ios' ? 'iOS' : 'macOS'
    if locale == 'en-US'
      path = "Build/Studio/Apple/appstoreconnect-screenshots-#{platform}.json"
      approval = "approval-#{platform_name}.json"
    else
      path = "Build/Studio/Apple/appstore/localizations/#{locale}/screenshots-#{platform}.json"
      approval = "approval-#{platform_name}.#{locale}.json"
    end
    prefix = Array.new(File.dirname(path).split('/').length, '..').join('/')
    sets = devices.map do |device|
      abort("store-art.json does not render #{device}") unless spec.fetch('devices').include?(device)
      display_type, width, height = DISPLAYS.fetch(device)
      {
        'screenshotDisplayType' => display_type,
        'path' => "#{prefix}/#{STORE_ART}/#{locale}/#{device}",
        'filter' => '*.png',
        'maxCount' => frame_count(spec, device),
        'allowedDimensions' => ["#{width}x#{height}"]
      }
    end
    expected[path] = {
      'appId' => APP_ID,
      'useReleaseVersion' => true,
      'locale' => locale,
      'platform' => platform_name,
      'screenshotSets' => sets,
      'quality' => {
        'enabled' => true,
        'rejectDuplicates' => true,
        'requireConsistentDimensions' => true,
        'minimumFileBytes' => 4096,
        'minimumKilobytesPerMegapixel' => 12,
        'requireApprovalManifest' => true,
        'approvalManifestPath' => "#{prefix}/#{APPROVALS}/#{approval}"
      }
    }
  end
end

# The release file lists every screenshot config; keep its formatting.
release_path = File.join(APPLE, 'powerforge.release.json')
release_text = File.read(release_path)
items = expected.keys.map { |path| "      #{JSON.generate(path)}" }.join(",\n")
release_expected = release_text.sub(/"ScreenshotConfigPaths": \[[^\]]*\]/m, "\"ScreenshotConfigPaths\": [\n#{items}\n    ]")
abort('powerforge.release.json has no ScreenshotConfigPaths list') if release_expected == release_text && !release_text.include?('"ScreenshotConfigPaths"')

stale = expected.reject do |path, data|
  file = File.join(ROOT, path)
  File.exist?(file) && JSON.parse(File.read(file)) == data
end
stale_release = release_expected != release_text

if ARGV.include?('--check')
  out_of_date = stale.keys + (stale_release ? ['Build/Studio/Apple/powerforge.release.json'] : [])
  unless out_of_date.empty?
    abort("Screenshot configs are out of date; run Build/Studio/Apple/store-art/screenshot-configs.rb:\n  #{out_of_date.join("\n  ")}")
  end
  puts "Screenshot configs current: #{expected.length} files."
else
  stale.each do |path, data|
    file = File.join(ROOT, path)
    Dir.mkdir(File.dirname(file)) unless Dir.exist?(File.dirname(file))
    File.write(file, JSON.pretty_generate(data) + "\n")
  end
  File.write(release_path, release_expected) if stale_release
  puts "Wrote #{stale.length} of #{expected.length} screenshot configs."
end
