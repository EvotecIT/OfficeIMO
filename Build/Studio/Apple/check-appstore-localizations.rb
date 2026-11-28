#!/usr/bin/env ruby
# Validates the authored App Store listings for every Studio language without Apple credentials.
require 'json'

APPLE = __dir__
ROOT = File.expand_path('../../..', APPLE)
APP_ID = '6818096229'
MARKETING_URL = 'https://officeimo.com/studio/'
SUPPORT_URL = 'https://github.com/EvotecIT/OfficeIMO/issues'
PRIVACY_URL = 'https://officeimo.com/studio/privacy/'
CATALOG = JSON.parse(File.read(File.join(APPLE, 'store-art/locales.json'))).fetch('locales').freeze
LOCALES = CATALOG.map { |entry| entry.fetch('store') }.freeze
abort('listing locale codes must be unique') unless LOCALES.uniq == LOCALES
abort('en-US must be a listing locale') unless LOCALES.include?('en-US')

# Every language the app ships in its String Catalogs has a listing, and no listing lacks the app's language.
def catalog_languages(path)
  catalog = JSON.parse(File.read(File.join(ROOT, path)))
  languages = catalog.fetch('strings').values.flat_map { |entry| entry.fetch('localizations', {}).keys }
  (languages + [catalog.fetch('sourceLanguage')]).uniq.sort
end
listing_languages = CATALOG.map { |entry| entry.fetch('language') }.sort
%w[Apple/OfficeIMOStudio/Shared/Localizable.xcstrings Apple/OfficeIMOStudio/Shared/InfoPlist.xcstrings].each do |path|
  languages = catalog_languages(path)
  abort("#{path} languages #{languages.join(', ')} differ from the listing languages #{listing_languages.join(', ')}") unless languages == listing_languages
end

def read_config(path, locale)
  data = JSON.parse(File.read(File.join(ROOT, path)))
  abort("#{path}: wrong app ID") unless data.fetch('appId') == APP_ID
  abort("#{path}: wrong locale") unless data.fetch('locale') == locale
  data
end

def text_field(metadata, field, maximum, path, bytes: false)
  value = metadata.fetch(field)
  abort("#{path}: #{field} must be non-empty text") unless value.is_a?(String) && value.strip.match?(/[^[:space:]]/)
  size = bytes ? value.bytesize : value.length
  abort("#{path}: #{field} exceeds #{maximum} #{bytes ? 'bytes' : 'characters'} (#{size})") if size > maximum
  value
end

def url_field(metadata, field, expected, path)
  abort("#{path}: unexpected #{field}") unless metadata.fetch(field) == expected
end

localizations = File.join(APPLE, 'appstore/localizations')
actual = Dir.children(localizations).select { |name| File.directory?(File.join(localizations, name)) }.sort
abort('listing locale directories differ from store-art/locales.json') unless actual == (LOCALES - ['en-US']).sort

app_info_paths = []
version_paths = []
LOCALES.each do |locale|
  base = locale == 'en-US' ? 'Build/Studio/Apple/appstoreconnect-' : "Build/Studio/Apple/appstore/localizations/#{locale}/"
  app_path = "#{base}app-info.json"
  app_info_paths << app_path
  app = read_config(app_path, locale).fetch('metadata')
  abort("#{app_path}: the app name is OfficeIMO Studio in every language") unless text_field(app, 'name', 30, app_path) == 'OfficeIMO Studio'
  text_field(app, 'subtitle', 30, app_path)
  url_field(app, 'privacyPolicyUrl', PRIVACY_URL, app_path)

  { 'ios' => 'iOS', 'macos' => 'macOS' }.each do |slug, platform|
    path = locale == 'en-US' ? "#{base}metadata-#{slug}.json" : "#{base}#{slug}.json"
    version_paths << path
    config = read_config(path, locale)
    abort("#{path}: wrong platform") unless config.fetch('platform') == platform
    abort("#{path}: release binding must follow the release version") unless config.fetch('useReleaseVersion') == true
    metadata = config.fetch('metadata')
    description = text_field(metadata, 'description', 4000, path)
    # Each listing names the read-only Office viewer and the source repository.
    %w[Word Excel PowerPoint https://github.com/EvotecIT/OfficeIMO].each do |required|
      abort("#{path}: description must mention #{required}") unless description.include?(required)
    end
    text_field(metadata, 'promotionalText', 170, path) if metadata.key?('promotionalText')
    keywords = text_field(metadata, 'keywords', 100, path, bytes: true)
    terms = keywords.split(',', -1)
    if terms.any? { |term| term.length < 3 || term.match?(/\A[[:space:]]|[[:space:]]\z/) }
      abort("#{path}: keywords must be at least three characters and separated by commas without surrounding spaces")
    end
    abort("#{path}: keywords contain repeated terms") if terms.map(&:downcase).uniq.length != terms.length
    url_field(metadata, 'marketingUrl', MARKETING_URL, path)
    url_field(metadata, 'supportUrl', SUPPORT_URL, path)
  end

  captions = File.join(APPLE, "store-art/captions/store-captions.#{locale}.json")
  abort("missing store captions for #{locale}") unless File.file?(captions)
end

release = JSON.parse(File.read(File.join(APPLE, 'powerforge.release.json'))).fetch('AppleApps')
abort('release must list every app information localization exactly once') unless release.fetch('AppInfoConfigPaths').sort == app_info_paths.sort
abort('release must list every platform listing exactly once') unless release.fetch('MetadataConfigPaths').sort == version_paths.sort
abort('release must not also set a single AppInfoConfigPath') if release.key?('AppInfoConfigPath')

caption_keys = JSON.parse(File.read(File.join(APPLE, 'store-art/captions/store-captions.en-US.json'))).keys.sort
LOCALES.each do |locale|
  keys = JSON.parse(File.read(File.join(APPLE, "store-art/captions/store-captions.#{locale}.json"))).keys.sort
  abort("store-captions.#{locale}.json keys differ from en-US") unless keys == caption_keys
end
puts "App Store metadata valid: #{LOCALES.length} app information localizations, #{LOCALES.length * 2} platform listings."
