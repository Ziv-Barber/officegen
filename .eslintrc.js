const fs = require('fs')
const path = require('path')

const prettierOptions = JSON.parse(
  fs.readFileSync(path.resolve(__dirname, '.prettierrc'), 'utf8')
)

module.exports = {
  parser: 'babel-eslint',
  extends: [
    'standard',
    'prettier',
    'prettier/flowtype',
    'prettier/react',
    'prettier/standard',
    'plugin:flowtype/recommended'
  ],
  parserOptions: {
    ecmaVersion: 6,
    sourceType: 'module',
    ecmaFeatures: {
      jsx: true
    }
  },
  rules: {
	'prettier/prettier': ['error', prettierOptions],
    camelcase: 0,
    'new-cap': 0,
    indent: [
      2,
      2,
      {
        SwitchCase: 1
      }
    ]
  },
  parserOptions: {},
  plugins: [
    'prettier',
    'flowtype'
  ],
  env: {
    mocha: true,
    browser: true,
    node: true,
    es6: true
  },
  settings: {
  }
}
