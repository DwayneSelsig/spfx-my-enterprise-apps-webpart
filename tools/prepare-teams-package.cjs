// SharePoint Sync to Teams recognizes this exact developer-provided package name.
const { copyFileSync } = require('node:fs');
const { resolve } = require('node:path');

const teamsFolder = resolve(__dirname, '../teams');
copyFileSync(resolve(teamsFolder, 'my-enterprise-apps.zip'), resolve(teamsFolder, 'TeamsSPFxApp.zip'));
console.log('Prepared teams/TeamsSPFxApp.zip for inclusion in the SharePoint solution.');
