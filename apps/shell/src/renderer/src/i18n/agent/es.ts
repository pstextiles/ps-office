import type { zh } from './zh'

export const es = {
  windowTitle: 'Agente',
  homeCardSub: 'Trabaje en una carpeta con Claude Code o Codex',
  intro:
    'El agente ejecuta su Claude Code o Codex en la carpeta que elija. Puede leer los archivos que hay en ella, crear otros nuevos, usar sus habilidades, ejecutar scripts y modificar los documentos abiertos en GenOffice.',
  folderLabel: 'Carpeta del proyecto',
  folderPick: 'Elegir una carpeta…',
  folderNone: 'No se ha elegido ninguna carpeta',
  recentLabel: 'Carpetas recientes',
  providerLabel: 'Programa',
  providerClaude: 'Claude Code',
  providerCodex: 'Codex',
  providerClaudeNote:
    'En Windows, los comandos que permite se ejecutan sin aislamiento y pueden acceder a todo su PC y a internet.',
  providerCodexNote:
    'Los comandos se ejecutan en el espacio aislado de Codex: solo se puede escribir en la carpeta del proyecto y no hay acceso a internet.',
  providerBlocked:
    'Aún no disponible: todavía no se puede impedir que Codex lea su propio archivo de inicio de sesión.',
  connChecking: 'Comprobando la conexión…',
  connUnavailable: 'No conectado. Inicie sesión en Configuración → IA.',
  connMetered: 'Conectado con un inicio de sesión de pago por uso que aceptó.',
  modeLabel: 'Cuándo preguntarle',
  modeAskCommands: 'Preguntar antes de cada comando',
  modeAskCommandsDesc:
    'Leer y escribir archivos en la carpeta del proyecto se hace sin preguntar. Antes de ejecutar un comando o un script, se le pregunta.',
  modeAskWrites: 'Preguntar antes de escribir y de ejecutar comandos',
  modeAskWritesDesc:
    'Solo la lectura se hace sin preguntar. Antes de escribir un archivo o de ejecutar un comando, se le pregunta.',
  modeTrust: 'Confiar en este proyecto',
  modeTrustDesc:
    'No se pregunta nada. Codex sigue dentro de su espacio aislado; en Claude Code, su propio revisor comprueba cada acción en su lugar.',
  warningTitle: 'Antes de empezar',
  warningCommands:
    'Los comandos que ejecuta el agente pueden cambiar archivos y usar internet. Un documento que alguien le haya enviado puede contener instrucciones ocultas.',
  warningRedaction:
    'Las partes que ocultó a la AI solo quedan ocultas en los documentos abiertos en GenOffice. Los archivos que el agente lee de la carpeta se envían tal cual.',
  warningUsage:
    'Un agente trabaja en muchos pasos y gasta más de su suscripción que el panel de AI.',
  warningAccept: 'Entendido',
  start: 'Iniciar',
  starting: 'Iniciando…',
  stop: 'Detener',
  newSession: 'Nueva sesión',
  send: 'Enviar',
  inputPlaceholder: '¿Qué debe hacer el agente?',
  you: 'Usted',
  agent: 'Agente',
  startedLine: '{program} se inició en {folder}',
  skillsLine: '{n} habilidades disponibles',
  skillsLineOne: '{n} habilidad disponible',
  toolRunning: 'En curso',
  toolDone: 'Hecho',
  toolFailed: 'Error',
  toolDenied: 'Rechazado',
  actRead: 'Leer',
  actWrite: 'Escribir',
  actCommand: 'Ejecutar',
  actAppTool: 'Editar un documento abierto',
  actSkill: 'Usar una habilidad',
  actOther: 'Usar una herramienta',
  openFile: 'Abrir',
  permTitle: 'El agente pide permiso para:',
  permCommand: 'Ejecutar este comando',
  permWrite: 'Escribir en este archivo',
  permRead: 'Leer un archivo fuera de la carpeta del proyecto',
  permOther: 'Usar la herramienta «{tool}»',
  permReason: 'Motivo: {reason}',
  permAllow: 'Permitir',
  permDeny: 'Rechazar',
  turnFailed: 'El agente no pudo terminar esta solicitud.',
  ended: 'La sesión ha terminado.',
  errBadFolder:
    'No se puede usar esta carpeta. Elija una carpeta con su propio trabajo: no una unidad entera, ni su carpeta personal, ni una carpeta de datos de inicio de sesión o de aplicaciones.',
  errNotConnected:
    'Este programa no está conectado. Inicie sesión en Configuración → IA y vuelva a intentarlo.',
  errNotAccepted: 'Primero lea las notas de arriba y confírmelas.',
  errUnsafeTools:
    'El programa se inició con herramientas que la aplicación no le dio, así que se detuvo.',
  errRateLimit: 'Ha agotado el límite de su suscripción. Vuelva a intentarlo más tarde.',
  errSignedOut: 'El programa no tiene la sesión iniciada. Inicie sesión en Configuración → IA.',
  errMetered: 'El programa tiene iniciada una sesión de pago por uso que no aceptó.',
  errSandboxUnavailable:
    'Codex no puede abrir en este PC el entorno aislado que contiene sus comandos, así que no se inició. Actualice Windows o elija Claude Code.',
  errNeedsUpdate:
    'El programa es demasiado antiguo para el modo agente. Actualícelo y vuelva a intentarlo.',
  errNotInstalled: 'El programa no está instalado.',
  errGeneric: 'Algo salió mal. Vuelva a intentarlo.',
  errUnsupportedPath:
    'Codex no puede usar una carpeta o un archivo cuyo nombre contenga [ ] * o ?. Cambie el nombre de la carpeta del proyecto o cierre los documentos con ese tipo de nombre y vuelva a empezar.',
  errProgramExited:
    'El programa se cerró mientras se iniciaba la sesión. Actualícelo, compruebe que se abre en una terminal y vuelva a intentarlo.',
  errProgramNoAnswer:
    'El programa no respondió a tiempo mientras se iniciaba la sesión. Vuelva a intentarlo.',
  errProgramRefused:
    'El programa no aceptó la configuración de la sesión. Actualícelo y vuelva a intentarlo.',
  errNoModel:
    'El programa no ofreció ningún modelo. Elija un modelo en Configuración → IA y vuelva a intentarlo.',
  errPrepareFailed: 'La aplicación no pudo preparar la sesión. Vuelva a intentarlo.',
  errCode: 'Si sigue ocurriendo, incluya este código al enviar sus comentarios: {code}',
  documentsUnprotected:
    'Los documentos abiertos después de iniciar la sesión ({n}) quedan fuera de la protección de escritura del programa. Detenga e inicie la sesión de nuevo para protegerlos.',
  connOff: 'Desactivado en esta aplicación. Puede volver a usarlo en Configuración → IA.',
  errOffInApp:
    'Este programa está desactivado en esta aplicación, así que no se envió nada más. Puede volver a usarlo en Configuración → IA.',
} satisfies Record<keyof typeof zh, string>
