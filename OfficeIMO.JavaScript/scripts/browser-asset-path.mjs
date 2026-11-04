import { relative, isAbsolute } from "node:path";

export function browserAssetPathIsContained(base, path, pathApi = { relative, isAbsolute }) {
  const pathFromBase = pathApi.relative(base, path);
  return !pathApi.isAbsolute(pathFromBase) && !pathFromBase.startsWith("..");
}
