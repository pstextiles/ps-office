import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join, resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

// PS Office: a Setup.exe run by hand over an existing install updates it in
// place, asking nothing (no install-mode page, no folder page).
const SHELL = resolve(import.meta.dirname, '..')
const require = createRequire(import.meta.url)
const nsh = readFileSync(join(SHELL, 'build/installer.nsh'), 'utf8')

function loadBuilderConfig() {
  const configModule = { exports: {} as ReturnType<typeof require> }
  runInNewContext(readFileSync(join(SHELL, 'electron-builder.cjs'), 'utf8'), {
    module: configModule,
    __dirname: SHELL,
    process: { platform: 'win32', arch: 'x64', env: {} },
    require: (id: string) =>
      id === 'node:fs' ? { ...require(id), existsSync: () => true } : require(id),
  })
  return configModule.exports
}

function macro(name: string): string {
  const match = new RegExp(`!macro ${name}\\r?\\n([\\s\\S]*?)!macroend`).exec(nsh)
  if (!match) throw new Error(`no ${name} macro in installer.nsh`)
  return match[1]!
}

describe('running a newer installer over an existing install', () => {
  it('has no folder page, so an existing install keeps its own folder', () => {
    const config = loadBuilderConfig()
    expect(config.nsis.oneClick).toBe(false)
    expect(config.nsis.allowToChangeInstallationDirectory).toBe(false)
  })

  it('keeps the mode of the one install on record instead of asking', () => {
    const body = macro('customInstallMode')
    // the in-app update path stays as it was
    expect(body).toMatch(/\$\{if\} \$\{isUpdated\}/)
    // a per-user install alone: update it for the current user
    expect(body).toMatch(
      /\$\{elseIf\} \$hasPerUserInstallation == "1"\s+\$\{andIf\} \$hasPerMachineInstallation == "0"\s+StrCpy \$isForceCurrentInstall "1"/,
    )
    // an all-users install alone: update it for all users (the page elevates)
    expect(body).toMatch(
      /\$\{elseIf\} \$hasPerMachineInstallation == "1"\s+\$\{andIf\} \$hasPerUserInstallation == "0"\s+StrCpy \$isForceMachineInstall "1"/,
    )
  })
})
