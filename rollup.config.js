import { nodeResolve } from '@rollup/plugin-node-resolve'
import commonjs from '@rollup/plugin-commonjs'
import typescript from '@rollup/plugin-typescript'
import terser from '@rollup/plugin-terser'
import globals from 'rollup-plugin-node-globals'
import builtins from 'rollup-plugin-node-builtins'

const onwarn = (warning) => {
  if (warning.code === 'CIRCULAR_DEPENDENCY') return
  console.warn(`(!) ${warning.message}`)
}

export default {
  input: 'src/index.ts',
  onwarn,
  output: [
    {
      file: 'dist/index.umd.js',
      format: 'umd',
      name: 'pptxtojson',
      sourcemap: true,
    },
    {
      file: 'dist/index.cjs',
      format: 'cjs',
      sourcemap: true,
    },
    {
      file: 'dist/index.js',
      format: 'es',
      sourcemap: true,
    },
  ],
  plugins: [
    nodeResolve({ preferBuiltins: false }),
    commonjs(),
    typescript({ tsconfig: './tsconfig.json' }),
    terser(),
    globals(),
    builtins(),
  ],
}
