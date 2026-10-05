export { partUri, relationshipPartUri, relativePartTarget } from "./uri.js";
export { corePropertiesXml, appPropertiesXml } from "./properties.js";
export type { CoreProperties, AppProperties } from "./properties.js";
import { partUri, relationshipPartUri, relativePartTarget } from "./uri.js";
import { corePropertiesXml, appPropertiesXml } from "./properties.js";
import type { CoreProperties, AppProperties } from "./properties.js";
import { ZipWriter } from "../zip/index.js";
import type { ZipWriterOptions, ZipEntrySource } from "../zip/index.js";
import type { PreparedEntry } from "../zip/entry.js";
import { ChunkedTextSink } from "../core/sinks.js";
import type { ByteSink } from "../core/sinks.js";
import { OfficeIMOError } from "../core/errors.js";
import { escapeXml, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";

export const relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
export const officeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
export const relationshipTypes = {
  officeDocument: officeRelationshipsNamespace + "/officeDocument",
  worksheet: officeRelationshipsNamespace + "/worksheet",
  styles: officeRelationshipsNamespace + "/styles",
  theme: officeRelationshipsNamespace + "/theme",
  customXml: officeRelationshipsNamespace + "/customXml",
  coreProperties: relationshipsNamespace + "/metadata/core-properties",
  appProperties: officeRelationshipsNamespace + "/extended-properties"
} as const;

export type PartSource = string | Blob | ZipEntrySource;
export interface PackagePart { readonly uri: string; readonly contentType: string; readonly data: PartSource; }
export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** Internal targets are absolute part URIs; serialized relative to the source part. */
  readonly target: string;
  readonly external?: boolean;
}
export interface OpcPackageOptions extends ZipWriterOptions { readonly invalidCharacterPolicy?: InvalidCharacterPolicy; }

function contentType(value: string): string {
  if (typeof value !== "string" || !/^[\w!#$&^.+-]+\/[\w!#$&^.+-]+$/.test(value)) throw new TypeError("Invalid content type.");
  return value;
}

/** Content-type catalog usable by every OPC format. Duplicate overrides are rejected. */
export class ContentTypes {
  private readonly defaults = new Map<string, string>([["rels", "application/vnd.openxmlformats-package.relationships+xml"], ["xml", "application/xml"]]);
  private readonly overrides = new Map<string, { uri: string; type: string }>();
  addDefault(extension: string, type: string): void {
    if (!/^[A-Za-z0-9]+$/.test(extension)) throw new TypeError("Invalid part extension.");
    const key = extension.toLowerCase(), previous = this.defaults.get(key);
    if (previous !== undefined && previous !== type) throw new TypeError("Conflicting default content type.");
    this.defaults.set(key, contentType(type));
  }
  addOverride(uri: string, type: string): void {
    uri = partUri(uri);
    if (this.overrides.has(uri.toLowerCase())) throw new TypeError("Duplicate content-type override.");
    this.overrides.set(uri.toLowerCase(), { uri, type: contentType(type) });
  }
  toXml(): string {
    return xmlDeclaration + '<Types xmlns="' + contentTypesNamespace + '">' + [...this.defaults].map(([ext, type]) =>
      '<Default Extension="' + escapeXml(ext) + '" ContentType="' + escapeXml(type) + '"/>').join("") + [...this.overrides.values()].map(({ uri, type }) =>
      '<Override PartName="' + escapeXml(uri) + '" ContentType="' + escapeXml(type) + '"/>').join("") + '</Types>';
  }
}

/** Format-neutral part and relationship owner. A source iterable is consumed once at finalization. */
export class OpcPackage {
  private readonly types = new ContentTypes();
  private readonly parts = new Map<string, PackagePart | { uri: string; prepared: PreparedEntry }>();
  private readonly relationships = new Map<string, Relationship[]>();
  private state = "open";
  private result: Promise<Blob> | undefined;
  constructor(private readonly options: OpcPackageOptions = {}) {}
  private open(): void { if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "OPC package is finalized."); }
  private reserve(uri: string, type: string): string {
    this.open(); uri = partUri(uri);
    if (/\/_rels\//i.test(uri) || uri.toLowerCase() === "/[content_types].xml") throw new TypeError("Package metadata part names are reserved.");
    if (this.parts.has(uri.toLowerCase())) throw new TypeError("Duplicate OPC part: " + uri);
    this.types.addOverride(uri, type); return uri;
  }
  addPart(part: PackagePart): void {
    const uri = this.reserve(part.uri, part.contentType);
    this.parts.set(uri.toLowerCase(), { ...part, uri, data: part.data instanceof Uint8Array ? new Uint8Array(part.data) : part.data });
  }
  /** @internal */
  addPrepared(uri: string, type: string, prepared: PreparedEntry): void {
    uri = this.reserve(uri, type); this.parts.set(uri.toLowerCase(), { uri, prepared });
  }
  addRelationship(source: string, relationship: Relationship): void {
    this.open(); source = source === "/" ? source : partUri(source);
    if (!/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(relationship.id)) throw new TypeError("Relationship id must be an XML ID.");
    if (!/^[A-Za-z][A-Za-z0-9+.-]*:[^\s<>"']+$/.test(relationship.type)) throw new TypeError("Relationship type must be an absolute URI.");
    const target = relationship.external ? relationship.target : partUri(relationship.target);
    if (relationship.external && (!target || /[\u0000-\u0020]/.test(target))) throw new TypeError("Invalid external relationship target.");
    const key = source.toLowerCase(), list = this.relationships.get(key) ?? [];
    if (list.some(r => r.id === relationship.id)) throw new TypeError("Duplicate relationship id: " + relationship.id);
    list.push({ ...relationship, target }); this.relationships.set(key, list);
  }
  setProperties(core: CoreProperties = {}, app: AppProperties = {}): void {
    // Build and validate both documents before reserving any parts.
    const coreXml = corePropertiesXml(core, this.options.invalidCharacterPolicy), appXml = appPropertiesXml(app, this.options.invalidCharacterPolicy);
    this.addPart({ uri: "/docProps/core.xml", contentType: "application/vnd.openxmlformats-package.core-properties+xml", data: coreXml });
    this.addPart({ uri: "/docProps/app.xml", contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml", data: appXml });
    this.addRelationship("/", { id: "core", type: relationshipTypes.coreProperties, target: "/docProps/core.xml" });
    this.addRelationship("/", { id: "app", type: relationshipTypes.appProperties, target: "/docProps/app.xml" });
  }
  private source(data: PartSource): ZipEntrySource {
    if (typeof data === "string") return async (sink: ByteSink) => { const text = new ChunkedTextSink(sink, this.options.signal); await text.write(data); await text.close(); };
    if (data instanceof Blob) return (async function* () {
      for (let i = 0; i < data.size; i += 65536) {
        let buffer: ArrayBuffer;
        try { buffer = await data.slice(i, i + 65536).arrayBuffer(); }
        catch (error) {
          if ((error as { name?: unknown } | null)?.name !== "NotReadableError") throw error;
          throw new OfficeIMOError("PLATFORM_UNAVAILABLE", "This environment cannot read Blob part data. Supply text or byte chunks; read Blobs before sending data to a worker.", { cause: error });
        }
        yield new Uint8Array(buffer);
      }
    })();
    return data;
  }
  toBlob(type = "application/zip"): Promise<Blob> {
    if (this.result) return this.result;
    this.open(); this.state = "finalizing";
    this.result = (async () => {
      const zip = new ZipWriter(undefined, this.options);
      try {
        for (const [source, list] of this.relationships) {
          if (source !== "/" && !this.parts.has(source)) throw new TypeError("Missing relationship source part: " + source);
          for (const rel of list) if (!rel.external && !this.parts.has(rel.target.toLowerCase())) throw new TypeError("Missing relationship target: " + rel.target);
        }
        for (const part of this.parts.values()) {
          if ("prepared" in part) await zip.addPrepared(part.uri.slice(1), part.prepared);
          else await zip.add(part.uri.slice(1), this.source(part.data));
        }
        for (const [key, list] of this.relationships) {
          const source = key === "/" ? "/" : this.parts.get(key)!.uri;
          const xml = xmlDeclaration + '<Relationships xmlns="' + relationshipsNamespace + '">' + list.map(r =>
            '<Relationship Id="' + escapeXml(r.id) + '" Type="' + escapeXml(r.type) + '" Target="' +
            escapeXml(r.external ? r.target : relativePartTarget(source, r.target)) + '"' + (r.external ? ' TargetMode="External"' : "") + '/>').join("") + '</Relationships>';
          await zip.add(relationshipPartUri(source).slice(1), this.source(xml));
        }
        await zip.add("[Content_Types].xml", this.source(this.types.toXml()));
        const blob = await zip.toBlob(type); this.state = "complete"; return blob;
      } catch (error) { this.state = "failed"; throw error; }
      finally { this.parts.clear(); this.relationships.clear(); }
    })();
    return this.result;
  }
}
