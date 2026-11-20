/** The engine's file access on the PC: files and folder listings through node:fs. */
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import type { FolderEntry, QueryHost } from './library'

export const nodeQueryHost: QueryHost = {
  readFile: (path) => readFileSync(path),
  listFolder: (path) => {
    const folder = path.endsWith('\\') || path.endsWith('/') ? path : `${path}\\`
    const out: FolderEntry[] = []
    for (const dirent of readdirSync(path, { withFileTypes: true })) {
      if (!dirent.isFile()) continue
      const st = statSync(join(path, dirent.name))
      out.push({
        name: dirent.name,
        folder,
        modified: st.mtime,
        created: st.birthtime,
        accessed: st.atime,
        hidden: false,
      })
    }
    return out
  },
}
