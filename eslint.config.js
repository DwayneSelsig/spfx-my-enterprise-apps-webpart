const { FlatCompat } = require('@eslint/eslintrc');
const js = require('@eslint/js');

const compat = new FlatCompat({
  baseDirectory: __dirname,
  recommendedConfig: js.configs.recommended
});

module.exports = [
  { ignores: ['**/*.js', '**/*.cjs', 'lib/**', 'lib-commonjs/**', 'dist/**', 'temp/**', 'sharepoint/**', 'release/**'] },
  ...compat.config(require('./.eslintrc.legacy.cjs'))
];
