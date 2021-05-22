module.exports = {
  env: {
    browser: true,
    commonjs: true,
    es2021: true
  },
  extends: [
    'eslint:recommended'
    // "plugin:@typescript-eslint/recommended"
  ],
  // "parser": "@typescript-eslint/parser",
  parserOptions: {
    ecmaVersion: 12
  },
  // "plugins": [
  //     "@typescript-eslint"
  // ],
  rules: {
    'object-curly-spacing': [
      'error',
      'always'
    ],
    indent: [
      'error',
      2
    ],
    'linebreak-style': [
      'error',
      'unix'
    ],
    quotes: [
      'error',
      'single'
    ],
    semi: [
      'error',
      'never'
    ],
    'comma-spacing': { before: false, after: true }
  }
}