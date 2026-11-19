import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { appBinaryForResources } from '../src/resources'

describe('appBinaryForResources', () => {
  it('finds the app in a custom Windows install directory', () => {
    expect(appBinaryForResources(join('D:\\Apps\\PSOffice', 'resources'), 'win32')).toBe(
      join('D:\\Apps\\PSOffice', 'PSOffice.exe'),
    )
  })

  it('resolves the same binary for the default Windows install directory', () => {
    const localAppData = 'C:\\Users\\test\\AppData\\Local'
    const resources = join(localAppData, 'Programs', 'PSOffice', 'resources')
    expect(appBinaryForResources(resources, 'win32')).toBe(
      join(localAppData, 'Programs', 'PSOffice', 'PSOffice.exe'),
    )
  })

  it('finds the app in a custom macOS bundle location', () => {
    expect(
      appBinaryForResources(join('/Volumes/Work/PSOffice.app/Contents/Resources'), 'darwin'),
    ).toBe(join('/Volumes/Work/PSOffice.app/Contents/MacOS/PSOffice'))
  })

  it('finds the app in a custom Linux prefix', () => {
    expect(appBinaryForResources(join('/opt/psoffice-custom/resources'), 'linux')).toBe(
      join('/opt/psoffice-custom/psoffice'),
    )
  })
})
