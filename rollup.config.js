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
  external: ['pdfjs-dist/legacy/build/pdf.mjs', 'canvas', 'jpegxr'],
  output: [
    {
      file: 'dist/index.umd.js',
      format: 'umd',
      name: 'pptxtojson-pro',
      globals: {
        'pdfjs-dist/legacy/build/pdf.mjs': 'pdfjsLib',
        'canvas': 'canvas',
        'jpegxr': 'JpegXR',
      },
    },
    {
      file: 'dist/index.cjs',
      format: 'cjs',
    },
    {
      file: 'dist/index.js',
      format: 'es',
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
