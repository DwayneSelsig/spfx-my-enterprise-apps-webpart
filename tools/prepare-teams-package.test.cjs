const { test } = require('node:test');
const assert = require('node:assert/strict');
const { mkdtempSync, mkdirSync, writeFileSync, readFileSync, existsSync, rmSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join } = require('node:path');
const { prepareTeamsPackage } = require('./prepare-teams-package.cjs');

function fixture(t, sourceVersion, stagedVersion) {
  const root = mkdtempSync(join(tmpdir(), 'enterprise-apps-release-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  for (const folder of ['copilot', 'temp/copilot', 'teams']) mkdirSync(join(root, folder), { recursive: true });
  writeFileSync(join(root, 'copilot/manifest.json'), JSON.stringify({ version: sourceVersion }));
  writeFileSync(join(root, 'temp/copilot/manifest.json'), JSON.stringify({ version: stagedVersion }));
  writeFileSync(join(root, 'teams/my-enterprise-apps.zip'), 'generated-package');
  return root;
}

for (const version of ['0.9.0', '0.10.0', '1.0.0']) {
  test(`preserves GitHub release ${version} without remapping`, t => {
    const root = fixture(t, version, version);
    prepareTeamsPackage(root, version);
    assert.equal(readFileSync(join(root, 'teams/TeamsSPFxApp.zip'), 'utf8'), 'generated-package');
  });
}

test('rejects an SDK hotfix version instead of silently publishing it', t => {
  const root = fixture(t, '0.9.0', '0.9.0.1');
  assert.throws(() => prepareTeamsPackage(root, '0.9.0'), /must match GitHub release/);
  assert.equal(existsSync(join(root, 'teams/TeamsSPFxApp.zip')), false);
});

test('rejects a source version not set by the release workflow', t => {
  const root = fixture(t, '1.0.0', '1.0.0');
  assert.throws(() => prepareTeamsPackage(root, '0.9.0'), /must match GitHub release/);
});
