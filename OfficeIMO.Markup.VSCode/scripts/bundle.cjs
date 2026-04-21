const path = require('path');
const esbuild = require('esbuild');

const args = new Set(process.argv.slice(2));
const watch = args.has('--watch');

const extensionBuildOptions = {
  entryPoints: [path.join(__dirname, '..', 'src', 'extension.ts')],
  outfile: path.join(__dirname, '..', 'out', 'extension.js'),
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node20',
  sourcemap: true,
  external: ['vscode'],
  logLevel: 'info'
};

const mermaidBuildOptions = {
  entryPoints: [path.join(__dirname, '..', 'src', 'mermaidPreview.ts')],
  outfile: path.join(__dirname, '..', 'out', 'mermaidPreview.js'),
  bundle: true,
  platform: 'browser',
  format: 'iife',
  target: 'es2020',
  minify: true,
  sourcemap: true,
  logLevel: 'info'
};

async function run() {
  if (watch) {
    const extensionContext = await esbuild.context(extensionBuildOptions);
    const mermaidContext = await esbuild.context(mermaidBuildOptions);
    await Promise.all([extensionContext.watch(), mermaidContext.watch()]);
    return;
  }

  await Promise.all([
    esbuild.build(extensionBuildOptions),
    esbuild.build(mermaidBuildOptions)
  ]);
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
