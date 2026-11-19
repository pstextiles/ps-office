import type { zh } from './zh'

export const es = {
  driveNav: 'Google Drive',
  driveDesc:
    'Elija archivos en Google Drive y ábralos y edítelos en GenOffice; al guardar vuelven a Drive. GenOffice solo puede abrir los archivos que usted elige.',
  driveConnect: 'Conectar Google Drive',
  driveConnecting: 'Inicie sesión con su cuenta de Google en el navegador y permita el acceso…',
  driveCancel: 'Cancelar',
  driveConnectedAs: 'Conectado como {email}',
  driveDisconnect: 'Desconectar',
  driveDisconnectConfirm:
    'Esto quita el inicio de sesión de Drive de este PC y pide a Google que retire el acceso de GenOffice. Las copias que ya están en este PC se conservan.',
  driveDisconnectYes: 'Desconectar',
  driveNotPersisted:
    'Este PC no puede guardar el inicio de sesión cifrado, así que la conexión termina al cerrar la aplicación.',
  driveExpired:
    'Google terminó la conexión: se retiró el acceso o la sesión caducó. Vuelva a conectar.',
  driveAiAccess: 'Dejar que la IA busque y lea archivos en Drive',
  driveAiAccessDesc:
    'Si está activado, la IA de los editores puede buscar y leer los archivos de Drive que eligió aquí. Lo que lee va al proveedor de IA que usa. La IA no puede modificar, eliminar ni compartir archivos de Drive.',
  drivePrivacy:
    'Las listas y el contenido de los archivos viajan solo entre este PC y Google. Nunca pasan por un servidor de los creadores de GenOffice.',
  driveFailDenied: 'No se conectó porque Google no lo permitió.',
  driveFailScope:
    'No se conectó porque no se marcó el acceso a Drive. Vuelva a conectar y marque la opción de Drive.',
  driveFailTimeout: 'El inicio de sesión no terminó en 10 minutos. Inténtelo de nuevo.',
  driveFailNetwork:
    'No se pudo contactar con Google. Revise su conexión a internet e inténtelo de nuevo.',
  driveRecent: 'Archivos elegidos',
  driveSearchPlaceholder: 'Buscar en los archivos elegidos',
  driveSearchResults: 'Resultados de «{q}»',
  driveRefresh: 'Actualizar',
  driveNotConnectedTitle: 'Google Drive no está conectado',
  driveNotConnectedBody:
    'Una vez conectado, puede elegir archivos de Drive, abrirlos y editarlos aquí, y al guardar vuelven a Drive.',
  driveOpenSettings: 'Conectar en Configuración',
  driveEmpty:
    'Aún no hay archivos. Haga clic en «Elegir de Google Drive» para seleccionar algunos.',
  driveLoading: 'Cargando…',
  driveLoadMore: 'Mostrar más',
  driveOpening: 'Abriendo…',
  driveGoogleBadge: 'Formato de Google',
  driveGoogleNote:
    'Un documento de Google se abre como archivo de Office. Al guardar, el original queda igual y se guarda un archivo de Office junto a él o en Mi unidad.',
  driveReadOnly: 'Solo lectura',
  driveReadOnlyNote:
    'No puede editar este archivo. Al guardar, se le preguntará si quiere guardar una copia en Drive.',
  driveUnsupported: 'Esta aplicación no puede abrir este tipo de archivo.',
  driveErrNotConnected: 'Google Drive está desconectado. Vuelva a conectarlo en Configuración.',
  driveErrNotFound: 'No se encontró el archivo. Se eliminó o esta cuenta no puede verlo.',
  driveErrForbidden: 'Esta cuenta no tiene permiso para descargar este archivo.',
  driveErrTooLarge: 'El archivo es demasiado grande para abrirlo.',
  driveErrRateLimit: 'Google está frenando las solicitudes por un momento. Inténtelo en breve.',
  driveErrNetwork: 'No se pudo contactar con Google Drive. Revise su conexión a internet.',
  driveErrDisk: 'No se pudo guardar el archivo en este PC.',
  drivePick: 'Elegir de Google Drive',
  drivePicking: 'Elija los archivos en la ventana de Google que se abrió en su navegador…',
  drivePickNote:
    'GenOffice solo puede abrir los archivos de Drive que elija en la ventana de Google. Los archivos que eligió una vez se quedan en esta lista.',
  driveTrainingTitle: 'Antes de usar archivos de Drive con la IA',
  driveTrainingBody:
    'Google no permite que los archivos de Drive vayan a un servicio de IA que los use para entrenar sus modelos. Los planes personales de Claude, ChatGPT y Codex pueden usar sus chats para entrenamiento si no lo desactiva: en Claude, en Configuración → Privacidad; en ChatGPT, en Configuración → Controles de datos → «Mejorar el modelo para todos». Los proveedores con clave de API no entrenan con datos de la API de forma predeterminada. Las conexiones de IA gratuitas nunca reciben archivos de Drive.',
  driveTrainingAck:
    'El entrenamiento de modelos está desactivado en las cuentas de IA que uso con archivos de Drive',
  driveFailTraining: 'Primero marque la opción sobre el entrenamiento de modelos de arriba.',
} satisfies Record<keyof typeof zh, string>
