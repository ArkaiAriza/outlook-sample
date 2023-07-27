// @eslint-disable no-restricted-imports
import * as path from 'path'
import * as OfficeDevCerts from 'office-addin-dev-certs'
import copy from 'rollup-plugin-copy'
import del from 'rollup-plugin-delete'
import serve from 'rollup-plugin-serve'
import watch from 'rollup-plugin-watch'

export const BASE_PATH = path.resolve(__dirname, '..')
export const PATH_SRC = path.resolve(BASE_PATH, 'src')
export const PATH_OFFICE = path.resolve(BASE_PATH, 'office')

const config = async () => {
  const server = await OfficeDevCerts.getHttpsServerOptions()

  const devConfig = {
    input: {},
    output: {
      dir: 'dist',
      format: 'iife',
      entryFileNames: '[name].[hash].js',
    },
    plugins: [
      del({ targets: 'dist/*' }),
      copy({
        targets: [
          {
            src: 'src/manifest.xml',
            dest: 'dist',
            rename: 'manifest.xml',
          },
          {
            src: 'src/launchevents.js',
            dest: 'dist',
            rename: 'launchevents.js',
          },
          {
            src: 'src/launch-events.html',
            dest: 'dist',
            rename: 'launch-events.html',
          },
          {
            src: 'src/icons',
            dest: 'dist',
          },
          {
            src: 'src/assets',
            dest: 'dist',
          },
        ],
      }),
      serve({
        contentBase: 'dist',
        host: 'localhost',
        port: 3000,
        https: server,
      }),
      watch({ dir: 'src' }),
    ],
  }
  return devConfig
}

export default config
