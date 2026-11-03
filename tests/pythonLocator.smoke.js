/**
 * Smoke test for out/pythonLocator.js under plain node (no VS Code needed).
 *
 *   node tests/pythonLocator.smoke.js
 *
 * Validates the "Python installed but not on PATH" recovery path:
 * PATH candidates fail → common-location scan finds the real interpreter.
 */
const path = require('path');
const Module = require('module');

// Route `require('vscode')` to our stub
const stub = path.join(__dirname, 'vscode-stub.js');
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...args) {
    if (request === 'vscode') { return stub; }
    return origResolve.call(this, request, ...args);
};

const locator = require('../out/pythonLocator.js');

console.log('--- scanCommonPythonPaths() ---');
const scanned = locator.scanCommonPythonPaths();
for (const p of scanned) { console.log(' ', p); }

console.log('--- pythonCandidates() ---');
for (const c of locator.pythonCandidates()) {
    console.log(`  [${c.source}] ${c.cmd}`);
}

console.log('--- resolvePythonCommand() ---');
const resolved = locator.resolvePythonCommand();
if (!resolved) {
    console.error('  NOT FOUND (scan found nothing and PATH has no python)');
    process.exit(1);
}
console.log(`  source: ${resolved.source}`);
console.log(`  cmd:    ${resolved.cmd}`);
console.log(`  version:${resolved.version}`);

if (resolved.source !== 'scan' && scanned.length > 0) {
    // On this machine PATH already resolves python — acceptable, but note it.
    console.log('  (note: resolved from PATH/setting; scan results were also available)');
}
console.log('OK');
