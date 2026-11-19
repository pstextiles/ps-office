import type { zh } from './zh'

export const es = {
  docPickTitle: 'Adjuntar desde documentos de GenOffice',
  docPickPane: 'Otro panel',
  docPickOpen: 'Documentos abiertos',
  docPickRecent: 'Documentos recientes',
  docPickEmpty: 'No hay documentos para adjuntar.',
  docPickMissing: 'No se encontró el archivo.',
  docPickUnsaved: 'Guárdelo primero.',
  docPickModified: 'Guárdelo para adjuntarlo.',
  docPickUnavailable: 'No se pudo leer la lista de documentos.',
  docPickLimit: 'Solo puede elegir hasta 20 documentos.',
  docPickSelected: '{n} seleccionados',
  docPickAttach: 'Adjuntar',
  docPickCancel: 'Cancelar',
  docPickWorking: 'Adjuntando…',
} satisfies Record<keyof typeof zh, string>
