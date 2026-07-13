#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const defaultRuntimeIdentifiers = ['win-x64', 'win-arm64', 'linux-x64', 'linux-arm64', 'osx-x64', 'osx-arm64'];
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';

function parseArgs(argv) {
  const options = {
    configuration: 'Release',
    framework: 'net8.0',
    runtimeIdentifiers: [...defaultRuntimeIdentifiers],
    outputDirectory: 'dist',
    skipNpmCi: false,
    skipRestore: false,
    publishMarketplace: false,
    preRelease: false
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = () => {
      index += 1;
      if (index >= argv.length) {
        throw new Error(`${arg} requires a value.`);
      }

      return argv[index];
    };

    switch (arg) {
      case '--configuration':
      case '-c':
        options.configuration = next();
        break;
      case '--framework':
      case '-f':
        options.framework = next();
        break;
      case '--runtimeIdentifiers':
      case '--runtime-identifiers':
      case '-r':
        options.runtimeIdentifiers = next().split(',').map((value) => value.trim()).filter(Boolean);
        break;
      case '--outputDirectory':
      case '--output-directory':
      case '-o':
        options.outputDirectory = next();
        break;
      case '--skipNpmCi':
      case '--skip-npm-ci':
        options.skipNpmCi = true;
        break;
      case '--skipRestore':
      case '--skip-restore':
        options.skipRestore = true;
        break;
      case '--publishMarketplace':
      case '--publish-marketplace':
        options.publishMarketplace = true;
        break;
      case '--preRelease':
      case '--pre-release':
        options.preRelease = true;
        break;
      default:
        throw new Error(`Unknown argument '${arg}'.`);
    }
  }

  if (!options.runtimeIdentifiers.length) {
    throw new Error('At least one runtime identifier must be provided.');
  }

  return options;
}

function assertChildPath(targetPath, parentPath) {
  const resolved = path.resolve(targetPath);
  const resolvedParent = path.resolve(parentPath);
  const parentWithSeparator = resolvedParent.endsWith(path.sep) ? resolvedParent : `${resolvedParent}${path.sep}`;
  const comparison = process.platform === 'win32'
    ? [resolved.toLowerCase(), resolvedParent.toLowerCase(), parentWithSeparator.toLowerCase()]
    : [resolved, resolvedParent, parentWithSeparator];

  if (comparison[0] !== comparison[1] && !comparison[0].startsWith(comparison[2])) {
    throw new Error(`Refusing to operate on '${resolved}' because it is outside '${resolvedParent}'.`);
  }

  return resolved;
}

function run(command, args, options = {}) {
  const isWindowsScript = process.platform === 'win32' && /\.(cmd|bat)$/i.test(command);
  const fileName = isWindowsScript ? (process.env.ComSpec || 'cmd.exe') : command;
  const argumentList = isWindowsScript
    ? ['/d', '/s', '/c', quoteWindowsCommand(command, args)]
    : args;
  const result = spawnSync(fileName, argumentList, {
    cwd: options.cwd,
    env: options.env ?? process.env,
    shell: false,
    stdio: 'inherit'
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`'${command} ${args.join(' ')}' failed with exit code ${result.status}.`);
  }
}

function quoteWindowsCommand(command, args) {
  return [command, ...args].map((value) => {
    const text = String(value);
    if (!/[ \t"&|<>^]/.test(text)) {
      return text;
    }

    return `"${text.replace(/"/g, '\\"')}"`;
  }).join(' ');
}

function removeDirectory(directoryPath) {
  fs.rmSync(directoryPath, { recursive: true, force: true });
}

function copyDirectoryContents(sourceDirectory, destinationDirectory) {
  fs.mkdirSync(destinationDirectory, { recursive: true });
  for (const entry of fs.readdirSync(sourceDirectory)) {
    fs.cpSync(path.join(sourceDirectory, entry), path.join(destinationDirectory, entry), {
      recursive: true,
      force: true
    });
  }
}

function vsceCommand(extensionRoot) {
  const windowsVsce = path.join(extensionRoot, 'node_modules', '.bin', 'vsce.cmd');
  if (fs.existsSync(windowsVsce)) {
    return { command: windowsVsce, prefix: [] };
  }

  const posixVsce = path.join(extensionRoot, 'node_modules', '.bin', 'vsce');
  if (fs.existsSync(posixVsce)) {
    return { command: posixVsce, prefix: [] };
  }

  const vsceMain = path.join(extensionRoot, 'node_modules', '@vscode', 'vsce', 'out', 'main.js');
  if (fs.existsSync(vsceMain)) {
    return { command: process.execPath, prefix: [vsceMain] };
  }

  throw new Error('VSCE was not found in node_modules. Run npm ci first.');
}

const options = parseArgs(process.argv.slice(2));
const extensionRoot = path.resolve(__dirname, '..');
const repoRoot = path.resolve(extensionRoot, '..');
const packagePath = path.join(extensionRoot, 'package.json');
const cliProject = path.join(repoRoot, 'OfficeIMO.Markup.Cli', 'OfficeIMO.Markup.Cli.csproj');

if (!fs.existsSync(packagePath)) {
  throw new Error(`package.json not found at ${packagePath}.`);
}

if (!fs.existsSync(cliProject)) {
  throw new Error(`OfficeIMO.Markup.Cli project not found at ${cliProject}.`);
}

if (!options.skipNpmCi) {
  if (fs.existsSync(path.join(extensionRoot, 'package-lock.json'))) {
    console.log('Installing extension dependencies with npm ci...');
    run(npmCommand, ['ci'], { cwd: extensionRoot });
  } else {
    console.log('Installing extension dependencies with npm install...');
    run(npmCommand, ['install'], { cwd: extensionRoot });
  }
}

const publishRoot = assertChildPath(path.join(extensionRoot, '.tmp', 'cli-publish'), extensionRoot);
removeDirectory(publishRoot);
fs.mkdirSync(publishRoot, { recursive: true });

const bundledCli = assertChildPath(path.join(extensionRoot, 'tools', 'OfficeIMO.Markup.Cli'), extensionRoot);
removeDirectory(bundledCli);
fs.mkdirSync(bundledCli, { recursive: true });
process.once('exit', () => removeDirectory(bundledCli));

const commonDotnetArgs = [
  'publish',
  cliProject,
  '-c', options.configuration,
  '-f', options.framework,
  '--nologo',
  '--verbosity', 'minimal',
  '-m:1',
  '-nr:false',
  '-p:BuildInParallel=false',
  '-p:UseSharedCompilation=false',
  '-p:DebugType=embedded'
];

if (options.skipRestore) {
  commonDotnetArgs.push('--no-restore');
}

const portablePublishRoot = assertChildPath(path.join(publishRoot, 'portable'), publishRoot);
fs.mkdirSync(portablePublishRoot, { recursive: true });
console.log(`Publishing OfficeIMO.Markup.Cli (${options.configuration}, ${options.framework}, portable fallback)...`);
run('dotnet', [...commonDotnetArgs, '-p:UseAppHost=false', '-o', portablePublishRoot], { cwd: extensionRoot });
copyDirectoryContents(portablePublishRoot, bundledCli);

for (const rid of options.runtimeIdentifiers) {
  const ridPublishRoot = assertChildPath(path.join(publishRoot, rid), publishRoot);
  fs.mkdirSync(ridPublishRoot, { recursive: true });

  console.log(`Publishing OfficeIMO.Markup.Cli (${options.configuration}, ${options.framework}, ${rid}, self-contained)...`);
  run('dotnet', [
    ...commonDotnetArgs,
    '-r', rid,
    '--self-contained', 'true',
    '-o', ridPublishRoot,
    '-p:PublishSingleFile=true',
    '-p:IncludeNativeLibrariesForSelfExtract=true',
    '-p:EnableCompressionInSingleFile=true'
  ], { cwd: extensionRoot });

  copyDirectoryContents(ridPublishRoot, path.join(bundledCli, rid));
}

removeDirectory(publishRoot);

console.log('Compiling VS Code extension...');
run(npmCommand, ['run', 'compile'], { cwd: extensionRoot });

const extensionPackage = JSON.parse(fs.readFileSync(packagePath, 'utf8'));
const outputRoot = assertChildPath(
  path.isAbsolute(options.outputDirectory)
    ? options.outputDirectory
    : path.join(extensionRoot, options.outputDirectory),
  extensionRoot
);
fs.mkdirSync(outputRoot, { recursive: true });

const vsixPath = path.join(outputRoot, `${extensionPackage.name}-${extensionPackage.version}.vsix`);
fs.rmSync(vsixPath, { force: true });

const vsce = vsceCommand(extensionRoot);
const packageArgs = [
  ...vsce.prefix,
  'package',
  '--allow-missing-repository',
  '--out',
  vsixPath
];
if (options.preRelease) {
  packageArgs.push('--pre-release');
}

console.log('Packaging VSIX...');
run(vsce.command, packageArgs, { cwd: extensionRoot });

if (options.publishMarketplace) {
  const token = process.env.VSCE_PAT;
  if (!token || !token.trim()) {
    throw new Error('VSCE_PAT is required when publishing to the Visual Studio Marketplace.');
  }

  const publishArgs = [
    ...vsce.prefix,
    'publish',
    '--packagePath',
    vsixPath
  ];
  if (options.preRelease) {
    publishArgs.push('--pre-release');
  }

  console.log('Publishing VSIX to the Visual Studio Marketplace...');
  run(vsce.command, publishArgs, { cwd: extensionRoot });
}

console.log(`VSIX: ${vsixPath}`);
