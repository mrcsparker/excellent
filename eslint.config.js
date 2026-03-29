'use strict';

var js = require('@eslint/js');
var globals = require('globals');
var importPlugin = require('eslint-plugin-import');
var n = require('eslint-plugin-n');
var promise = require('eslint-plugin-promise');
var sonarjs = require('eslint-plugin-sonarjs');
var tseslint = require('typescript-eslint');

var generatedFiles = [
  'coverage/**',
  'dist/**',
  'generated/**'
];

var vendorFiles = [
  'demo/bs3/**'
];

var authoredNodeFiles = [
  'eslint.config.js',
  'index.js',
  'browser.js',
  'bin/**/*.js',
  'scripts/**/*.js',
  'test/**/*.js'
];

var authoredEsmNodeFiles = [
  'browser.mjs'
];

var browserFiles = [
  'demo/scripts/demo.js'
];

var devToolFiles = [
  'scripts/**/*.js'
];

var eslintConfigFiles = [
  'eslint.config.js'
];

var testFiles = [
  'test/**/*.js'
];

var commonRules = {
  'curly': ['error', 'all'],
  'default-case-last': 'error',
  'eqeqeq': ['error', 'always', { null: 'ignore' }],
  'no-alert': 'error',
  'no-console': 'error',
  'no-eval': 'error',
  'no-extend-native': 'error',
  'no-implied-eval': 'error',
  'no-loop-func': 'error',
  'no-prototype-builtins': 'error',
  'no-shadow': ['error', { hoist: 'functions' }],
  'no-unused-vars': ['error', {
    args: 'after-used',
    argsIgnorePattern: '^_',
    ignoreRestSiblings: true
  }],
  'no-use-before-define': ['error', {
    classes: true,
    functions: false,
    variables: true
  }],
  'radix': 'error',
  'yoda': 'error',
  'sonarjs/no-all-duplicated-branches': 'error',
  'sonarjs/no-identical-functions': 'error',
  'sonarjs/no-unused-collection': 'error'
};

var tsBaseConfig = tseslint.configs.strictTypeChecked[0];
var tsRules = Object.assign(
  {},
  tseslint.configs.strictTypeChecked[1].rules,
  tseslint.configs.strictTypeChecked[2].rules,
  tseslint.configs.stylisticTypeChecked[2].rules
);

module.exports = [
  {
    ignores: [
      vendorFiles[0],
      generatedFiles[0],
      generatedFiles[1],
      generatedFiles[2]
    ],
  },
  {
    linterOptions: {
      reportUnusedDisableDirectives: 'error'
    }
  },
  {
    files: authoredNodeFiles,
    languageOptions: {
      ecmaVersion: 'latest',
      globals: globals.node,
      sourceType: 'commonjs'
    },
    plugins: {
      import: importPlugin,
      n: n,
      promise: promise,
      sonarjs: sonarjs
    },
    settings: {
      node: {
        version: '>=20.13.0'
      },
      'import/resolver': {
        node: {
          extensions: ['.js', '.json', '.ts', '.d.ts', '.cjs', '.mjs']
        }
      }
    },
    rules: Object.assign(
      {},
      js.configs.recommended.rules,
      importPlugin.flatConfigs.recommended.rules,
      n.configs['flat/recommended-script'].rules,
      promise.configs['flat/recommended'].rules,
      commonRules,
      {
        'import/no-extraneous-dependencies': ['error', {
          devDependencies: [
            'eslint.config.js',
            'scripts/**',
            'test/**'
          ],
          includeTypes: true
        }],
        'import/no-unresolved': ['error', {
          amd: false,
          commonjs: true
        }],
        'n/no-process-exit': 'off',
        'promise/always-return': 'off'
      }
    )
  },
  {
    files: authoredEsmNodeFiles,
    languageOptions: {
      ecmaVersion: 'latest',
      globals: globals.node,
      sourceType: 'module'
    },
    plugins: {
      import: importPlugin,
      n: n,
      promise: promise,
      sonarjs: sonarjs
    },
    settings: {
      node: {
        version: '>=20.13.0'
      },
      'import/resolver': {
        node: {
          extensions: ['.js', '.json', '.ts', '.d.ts', '.cjs', '.mjs']
        }
      }
    },
    rules: Object.assign(
      {},
      js.configs.recommended.rules,
      importPlugin.flatConfigs.recommended.rules,
      n.configs['flat/recommended-module'].rules,
      promise.configs['flat/recommended'].rules,
      commonRules,
      {
        'import/no-extraneous-dependencies': ['error', {
          devDependencies: false,
          includeTypes: true
        }],
        'n/no-process-exit': 'off',
        'promise/always-return': 'off'
      }
    )
  },
  {
    files: devToolFiles,
    rules: {
      'n/no-unpublished-require': 'off'
    }
  },
  {
    files: testFiles,
    rules: {
      'n/no-unpublished-require': 'off'
    }
  },
  {
    files: eslintConfigFiles,
    rules: {
      'import/no-unresolved': 'off',
      'n/no-unpublished-require': 'off'
    }
  },
  {
    files: browserFiles,
    languageOptions: {
      ecmaVersion: 'latest',
      globals: Object.assign({}, globals.browser, {
        'Excellent': 'readonly'
      }),
      sourceType: 'script'
    },
    plugins: {
      promise: promise,
      sonarjs: sonarjs
    },
    rules: Object.assign(
      {},
      js.configs.recommended.rules,
      promise.configs['flat/recommended'].rules,
      commonRules,
      {
        'no-alert': 'off',
        'no-console': 'off',
        'no-implicit-globals': 'error',
        'promise/always-return': 'off'
      }
    )
  },
  {
    files: ['**/*.ts', '**/*.d.ts'],
    languageOptions: Object.assign(
      {},
      tsBaseConfig.languageOptions,
      {
        parserOptions: {
          projectService: true,
          tsconfigRootDir: __dirname
        }
      }
    ),
    plugins: Object.assign(
      {},
      tsBaseConfig.plugins,
      {
        import: importPlugin,
        promise: promise,
        sonarjs: sonarjs
      }
    ),
    settings: {
      'import/resolver': {
        node: {
          extensions: ['.js', '.json', '.ts', '.d.ts', '.cjs', '.mjs']
        }
      }
    },
    rules: Object.assign(
      {},
      tsRules,
      importPlugin.flatConfigs.recommended.rules,
      importPlugin.flatConfigs.typescript.rules,
      promise.configs['flat/recommended'].rules,
      {
        'import/no-extraneous-dependencies': ['error', {
          devDependencies: [
            'test/**'
          ],
          includeTypes: true
        }],
        'import/no-unresolved': 'error',
        '@typescript-eslint/consistent-type-imports': ['error', {
          prefer: 'type-imports'
        }],
        '@typescript-eslint/no-floating-promises': 'error',
        '@typescript-eslint/no-import-type-side-effects': 'error',
        '@typescript-eslint/no-unused-vars': ['error', {
          args: 'after-used',
          argsIgnorePattern: '^_',
          ignoreRestSiblings: true
        }],
        'promise/always-return': 'off',
        'sonarjs/no-identical-functions': 'error'
      }
    )
  },
  {
    files: [
      'src/formula/**/*.ts',
      'src/workbook/**/*.ts',
      'src/excellent.loader.ts',
      'src/excellent.parser.ts',
      'src/excellent.util.ts',
      'src/excellent.xlsx-simple.ts',
      'src/excellent.xlsx.shared.ts',
      'src/excellent.xlsx.ts',
      'src/index.ts'
    ],
    rules: {
      '@typescript-eslint/array-type': 'off',
      '@typescript-eslint/consistent-type-definitions': 'off',
      '@typescript-eslint/no-require-imports': 'off',
      '@typescript-eslint/no-unnecessary-condition': 'off',
      '@typescript-eslint/no-unnecessary-type-assertion': 'off',
      '@typescript-eslint/no-unnecessary-type-conversion': 'off',
      '@typescript-eslint/no-unnecessary-type-parameters': 'off',
      '@typescript-eslint/no-useless-default-assignment': 'off',
      '@typescript-eslint/no-unsafe-argument': 'off',
      '@typescript-eslint/no-unsafe-assignment': 'off',
      '@typescript-eslint/no-unsafe-call': 'off',
      '@typescript-eslint/no-unsafe-member-access': 'off',
      '@typescript-eslint/no-unsafe-return': 'off',
      '@typescript-eslint/prefer-nullish-coalescing': 'off',
      '@typescript-eslint/prefer-optional-chain': 'off',
      '@typescript-eslint/restrict-plus-operands': 'off'
    }
  },
  {
    files: testFiles,
    rules: {
      'n/no-unsupported-features/node-builtins': 'off'
    }
  },
  {
    files: [
      'browser.js',
      'index.js',
      'test/browser-smoke/browser_bundle_smoke_spec.js',
      'test/evaluator/formula_runtime_spec.js',
      'test/parser/formula_parser_regression_spec.js',
      'test/parser/formula_parser_spec.js'
    ],
    rules: {
      'n/no-unpublished-require': 'off'
    }
  }
];
