const OFF = 0; const WARN = 1; const ERROR = 2

module.exports = {
  env: {
    browser: true,
    es2021: true
  },
  extends: 'standard',
  rules: {
    'no-unused-vars': WARN
  },
  overrides: [
    {
      env: {
        node: true
      },
      files: ['.eslintrc.{js,cjs}'],
      parserOptions: {
        sourceType: 'script'
      }
    }
  ],
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module'
  }
}
