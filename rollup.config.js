import resolve from 'rollup-plugin-node-resolve'
import commonjs from 'rollup-plugin-commonjs'
import babel from 'rollup-plugin-babel'
import json from 'rollup-plugin-json'
import packageJson from './package.json'

export default [
  // browser-friendly UMD build
  {
    entry: 'index.js',
    dest: packageJson.browser,
    format: 'umd',
    moduleName: 'officegen',
    plugins: [
      json({
        // All JSON files will be parsed by default,
        // but you can also specifically include/exclude files
        include: 'node_modules/**',
        // exclude: [ 'node_modules/foo/**', 'node_modules/bar/**' ],

        // for tree-shaking, properties will be declared as
        // variables, using either `var` or `const`
        preferConst: true, // Default: false

        // specify indentation for the generated default export —
        // defaults to '\t'
        indent: '  '
      }),
      resolve(),
      commonjs(),
      babel({
        exclude: ['node_modules/**']
      })
    ]
  },

  // CommonJS (for Node) and ES module (for bundlers) build.
  // (We could have three entries in the configuration array
  // instead of two, but it's quicker to generate multiple
  // builds from a single configuration where possible, using
  // the `targets` option which can specify `dest` and `format`)
  {
    entry: 'index.js',
    targets: [
      { dest: packageJson.main, format: 'cjs' },
      { dest: packageJson.module, format: 'es' }
    ],
    plugins: [
      json({
        // All JSON files will be parsed by default,
        // but you can also specifically include/exclude files
        include: 'node_modules/**',
        // exclude: [ 'node_modules/foo/**', 'node_modules/bar/**' ],

        // for tree-shaking, properties will be declared as
        // variables, using either `var` or `const`
        preferConst: true, // Default: false

        // specify indentation for the generated default export —
        // defaults to '\t'
        indent: '  '
      }),
      babel({
        exclude: ['node_modules/**']
      })
    ]
  }
]
