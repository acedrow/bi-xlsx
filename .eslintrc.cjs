const OFF = 0; const WARN = 1; const ERROR = 2

module.exports = {
  env: {
    browser: true,
    es2021: true
  },
  extends: 'standard',
  rules: {
    'no-unused-vars': WARN,
    'object-curly-newline': ['error', { multiline: true }],
    'object-curly-spacing': ['error', 'always'],
    'object-property-newline': ['error', { allowAllPropertiesOnSameLine: true }]
  },
  overrides: [
    {
      env: { node: true },
      files: ['.eslintrc.{js,cjs}'],
      parserOptions: { sourceType: 'script' }
    }
  ],
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module'
  }
}
