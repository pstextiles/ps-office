// Builds PSQueryRefresh.exe: one Windows executable (Node's single executable
// application) holding the refresh program, its worker and the xlsx sidecar.
//
//   node tools/query-refresh/build.mjs [--sidecar <path to xlsx-sidecar.exe>]
//
// Needs the workspace installed (npm ci) and the sidecar built
// (npm run native:build -w @genoffice/sheets). Output: tools/query-refresh/dist/.
import { build } from 'esbuild'
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '../..')
const dist = join(here, 'dist')
const argSidecar = process.argv.indexOf('--sidecar')
const sidecar =
  argSidecar > 0
    ? resolve(process.argv[argSidecar + 1])
    : join(root, 'apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar.exe')
if (!existsSync(sidecar)) throw new Error(`xlsx sidecar not found: ${sidecar}`)

rmSync(dist, { recursive: true, force: true })
mkdirSync(dist, { recursive: true })

// Vite's `?modulePath` worker import (builtin-refresh.ts): the program passes its
// own worker, so the app's worker chunk path is empty here
const modulePath = {
  name: 'module-path',
  setup(b) {
    b.onResolve({ filter: /\?modulePath$/ }, (args) => ({
      path: args.path,
      namespace: 'module-path',
    }))
    b.onLoad({ filter: /.*/, namespace: 'module-path' }, () => ({
      contents: 'export default ""',
      loader: 'js',
    }))
  },
}

const common = {
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node22',
  logLevel: 'warning',
  plugins: [modulePath],
  // the program never runs Electron code; a stray import must fail the build, not the run
  external: [],
  legalComments: 'none',
}

await build({
  ...common,
  entryPoints: [join(here, 'refresh-cli.ts')],
  outfile: join(dist, 'refresh-cli.cjs'),
})
await build({
  ...common,
  entryPoints: [join(root, 'apps/sheets/src/main/query-refresh-worker.ts')],
  outfile: join(dist, 'refresh-worker.cjs'),
})
copyFileSync(sidecar, join(dist, 'xlsx-sidecar.exe'))

const seaConfig = {
  main: join(dist, 'refresh-cli.cjs'),
  output: join(dist, 'sea-prep.blob'),
  disableExperimentalSEAWarning: true,
  useCodeCache: false,
  assets: {
    'refresh-worker.cjs': join(dist, 'refresh-worker.cjs'),
    'xlsx-sidecar.exe': join(dist, 'xlsx-sidecar.exe'),
  },
}
writeFileSync(join(dist, 'sea-config.json'), JSON.stringify(seaConfig, null, 2))
execFileSync(process.execPath, ['--experimental-sea-config', join(dist, 'sea-config.json')], {
  stdio: 'inherit',
})

const exe = join(dist, 'PSQueryRefresh.exe')
copyFileSync(process.execPath, exe)
brandExecutable(exe)
execFileSync(
  process.execPath,
  [
    join(root, 'node_modules/postject/dist/cli.js'),
    exe,
    'NODE_SEA_BLOB',
    join(dist, 'sea-prep.blob'),
    '--sentinel-fuse',
    'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
  ],
  { stdio: 'inherit' },
)
console.log(`built ${exe}`)

/**
 * The PS Office icon and this program's name in the copied node.exe (resedit,
 * already a dependency of the installer build). Node's signature is dropped:
 * injecting the program would break it anyway, and the program is unsigned
 * as the installer is.
 */
function brandExecutable(path) {
  const require = createRequire(import.meta.url)
  const ResEdit = require('resedit')
  const executable = ResEdit.NtExecutable.from(readFileSync(path), { ignoreCert: true })
  const resources = ResEdit.NtExecutableResource.from(executable)
  const icon = ResEdit.Data.IconFile.from(readFileSync(join(root, 'apps/shell/build/icon.ico')))
  const groups = ResEdit.Resource.IconGroupEntry.fromEntries(resources.entries)
  const group = groups[0]
  ResEdit.Resource.IconGroupEntry.replaceIconsForResource(
    resources.entries,
    group ? group.id : 1,
    group ? group.lang : 1033,
    icon.icons.map((item) => item.data),
  )
  const [version] = ResEdit.Resource.VersionInfo.fromEntries(resources.entries)
  if (version) {
    const language = version.getAllLanguagesForStringValues()[0] ?? { lang: 1033, codepage: 1200 }
    version.setStringValues(language, {
      ProductName: 'PS Query Refresh',
      FileDescription: 'PS Query Refresh',
      CompanyName: 'PS Textiles',
      InternalName: 'PSQueryRefresh',
      OriginalFilename: 'PSQueryRefresh.exe',
      LegalCopyright: 'PS Textiles. Apache-2.0, see NOTICE',
    })
    version.outputToResourceEntries(resources.entries)
  }
  resources.outputResource(executable)
  writeFileSync(path, Buffer.from(executable.generate()))
}
