#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Read version from package.json
const packageJson = JSON.parse(fs.readFileSync(path.join(__dirname, '../package.json'), 'utf-8'));
const version = packageJson.version;

// Execute bun build command with --define
// On Windows, we need to be careful with quotes
// Using \\\" to escape the quotes in the command line
const command = `bun build src/bin.ts --compile --minify --outfile bin/xlsx-cli --define __VERSION__=\\\"${version}\\\"`;

console.log(`Building xlsx-cli with version: ${version}`);
console.log(`Command: ${command}`);
console.log(`Version length: ${version.length}, Version value: "${version}"`);

try {
  execSync(command, {
    cwd: path.join(__dirname, '..'),
    stdio: 'inherit'
  });
  console.log(`\nBuild completed successfully! Version: ${version}`);
} catch (error) {
  console.error('Build failed:', error);
  process.exit(1);
}
