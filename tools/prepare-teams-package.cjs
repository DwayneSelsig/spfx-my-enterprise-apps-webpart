// SharePoint Sync to Teams recognizes this exact developer-provided package name.
const { copyFileSync, readFileSync } = require('node:fs');
const { resolve } = require('node:path');

function prepareTeamsPackage(root, releaseVersion) {
  if (releaseVersion !== undefined) {
    // Keep GitHub's tag authoritative, including 0.x releases. Never publish an
    // SDK-generated local hotfix counter in place of the requested release.
    const source = JSON.parse(readFileSync(resolve(root, 'copilot/manifest.json'), 'utf8'));
    const staged = JSON.parse(readFileSync(resolve(root, 'temp/copilot/manifest.json'), 'utf8'));
    if (source.version !== releaseVersion || staged.version !== releaseVersion) {
      throw new Error(`Teams app version must match GitHub release ${releaseVersion}; ` +
        `source=${source.version}, built=${staged.version}.`);
    }
  }
  const teamsFolder = resolve(root, 'teams');
  copyFileSync(resolve(teamsFolder, 'my-enterprise-apps.zip'), resolve(teamsFolder, 'TeamsSPFxApp.zip'));
  console.log('Prepared teams/TeamsSPFxApp.zip for inclusion in the SharePoint solution.');
}

module.exports = { prepareTeamsPackage };
if (require.main === module) {
  prepareTeamsPackage(resolve(__dirname, '..'), process.env.RELEASE_VERSION);
}
