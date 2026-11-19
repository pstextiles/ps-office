import type { zh } from './zh'

export const es = {
  aiAssistant: 'Asistente de IA',
  cardChecking: 'Comprobando…',
  cardNotInstalled: 'No instalado',
  cardSignedOut: 'Sesión no iniciada',
  cardSigningIn: 'Iniciando sesión…',
  cardReady: 'Conectado',
  cardReadyPlan: 'Conectado · {plan}',
  cardMetered: 'Sesión iniciada, pero no con una suscripción',
  cardMeteredAccepted: 'En uso · se cobra por consumo',
  cardError: 'No se pudo comprobar la conexión',
  cardErrNeedsUpdate:
    'El programa {provider} de este PC es demasiado antiguo. Actualícelo y compruebe de nuevo.',
  cardErrNeedsUpdateVersion:
    'El programa {provider} de este PC ({version}) es demasiado antiguo; se necesita la {minimum} o posterior. Ejecute {command} en una terminal y compruebe de nuevo.',
  cardErrNeedsUpdatePath: 'Archivo que ejecuta esta aplicación:',
  cardErrNeedsUpdateWindows:
    'Actualizar {provider} dentro de WSL (Ubuntu) o en otra carpeta no cambia este archivo. Ejecute {command} en Windows PowerShell.',
  claudeDesktopNote:
    'La aplicación de escritorio de Claude no es Claude Code, que es lo que ejecuta esta aplicación. Claude Code es un programa aparte que se usa en una terminal, así que instalar o actualizar la aplicación de escritorio no lo cambia.',
  cardErrTimeout: 'El programa {provider} no respondió a tiempo. Compruebe de nuevo en un momento.',
  cardErrFailed: 'Algo salió mal al comprobar {provider}. Compruebe de nuevo.',
  actionInstall: 'Instalar',
  actionLogin: 'Iniciar sesión',
  actionLoginWindow: '¿No se abrió el navegador?',
  actionLoginWithCode: '¿El navegador muestra un código?',
  actionLoginBrowser: 'Iniciar sesión en el navegador',
  actionRecheck: 'Comprobar de nuevo',
  actionAcceptMetered: 'Usar de todos modos',
  actionInstallHelp: 'Guía oficial de instalación',
  actionOpenSettings: 'Abrir ajustes',
  actionDismiss: 'Cerrar',
  actionSwitchAccount: 'Iniciar sesión con otra cuenta',
  actionLogout: 'Cerrar sesión',
  installNote:
    'El comando de instalación oficial de {vendor} se ejecuta en una nueva ventana de terminal.',
  loginNote: 'La página de inicio de sesión que se abre es la de {vendor}.',
  signingInBrowserNote:
    'Inicie sesión con su cuenta de {vendor} en el navegador que se abrió y permita el acceso. Se conecta solo cuando termine.',
  signingInWindowNote:
    'Siga las instrucciones de la parte superior de la ventana de inicio de sesión que se acaba de abrir. Inicie sesión en la página de {vendor} que se abrió en el navegador: la ventana se cierra sola y se completa la conexión. Una página de inicio de sesión abierta antes ya no funciona; ciérrela.',
  meteredWarning: 'Esto no es una suscripción: se cobra según lo que use.',
  meteredMeaning:
    'Este inicio de sesión usa una clave de API o una cuenta de pago por consumo, o no se pudo confirmar su tipo. «Usar de todos modos» permite que la IA funcione con esta conexión, y cada solicitud se factura a esa cuenta.',
  connectTitle: 'Conecte una IA para empezar',
  connectBody:
    'La IA funciona con la suscripción de Claude o ChatGPT que ya tiene. No hace falta ninguna clave de API.',
  connectProviderTitle: 'La conexión de {provider} no se puede usar ahora',
  connectKeyBody:
    'La configuración de {provider} no está terminada. Complétela en Ajustes o elija otra IA.',
  errNotConnected:
    'No hay ninguna IA conectada. Conecte primero una suscripción de Claude o Codex, o una IA gratuita.',
  errNotInstalled: 'No se encontró el programa {provider} en este PC.',
  errSignedOut: 'La sesión de {provider} ha caducado. Vuelva a iniciar sesión para continuar.',
  errRateLimit: 'Se ha agotado el límite de la suscripción de {provider}.',
  errRateLimitReset:
    'Se ha agotado el límite de la suscripción de {provider}. Se restablece a las {time}.',
  errChatGptAppLimit:
    'Esta aplicación ha agotado el uso de ChatGPT del que puede disponer. Puede ser el límite de todo su plan o uno fijado para esta aplicación. Revíselo en ChatGPT, en Configuración → Uso.',
  errUnsafeTools:
    'El programa {provider} se inició con sus propias herramientas activadas, así que la ejecución se detuvo por seguridad.',
  errMeteredNotAccepted:
    'La conexión de {provider} es de pago por consumo y aún no la ha permitido. Elija «Usar de todos modos» en los ajustes de conexión para usarla.',
  errNeedsUpdate:
    'El programa {provider} es demasiado antiguo para esta función. Actualícelo y vuelva a intentarlo.',
  errHeadlessLimit: 'Se alcanzó el límite por hora de solicitudes automáticas.',
  errHeadlessLimitReset:
    'Se alcanzó el límite por hora de solicitudes automáticas. Volverán a ser posibles a las {time}.',
  connectFreeHint:
    '¿No tiene suscripción? Puede conectar una IA gratuita en Ajustes en un minuto aproximadamente.',
  actionConnectFree: 'Conectar una IA gratuita',
  errFreeLimit:
    'Se agotó el uso gratuito de hoy. Vuelva a intentarlo mañana o elija otra IA en Ajustes.',
  errFreeLimitReset: 'Se agotó el uso gratuito de hoy. Podrá volver a usarlo a las {time}.',
  errFreeTierModel:
    'Este modelo no tiene nivel gratuito, así que una clave gratuita no puede usarlo. Elija un modelo incluido en el nivel gratuito (por ejemplo, uno Flash) o active la facturación en la consola del proveedor.',
  continueWith: 'Continuar con {provider}',
  continueMeteredNote:
    'La conexión de {provider} no es una suscripción: continuar se cobra por consumo.',
  usageLine: '{percent} % usado · se restablece a las {time}',
  usageWindowFiveHour: 'Límite de 5 horas',
  usageWindowSevenDay: 'Límite de 7 días',
  noticeTitle: 'Antes de usar la IA con su suscripción',
  noticeSend:
    'El contenido del documento y su conversación se envían a Anthropic u OpenAI con su propia cuenta. Cada solicitud lleva además el nombre, el título y la descripción de cada habilidad configurada como «Decide la IA»; también se envían las habilidades que se usan y los archivos de habilidad que lee la IA.',
  noticeTraining:
    'Que ese contenido se use para entrenar IA depende de los ajustes de su propia cuenta en esa empresa.',
  noticeCodexHistory:
    'Esta aplicación desactiva el guardado del historial de conversaciones de Codex, pero Codex puede seguir guardando sus propios registros de diagnóstico en este PC.',
  noticePdfWholeDocument:
    'En el editor de PDF, la lectura de páginas con IA y la traducción de documentos pueden enviar imágenes de las páginas y todo el texto del documento.',
  noticeAck: 'Entendido',
  codexNotice:
    'Usar una suscripción de Codex dentro de una aplicación no es un método recomendado oficialmente por OpenAI.',
  truncatedNotice:
    'La entrada era demasiado larga, así que se recortó una parte antes de enviarla.',
  pickerAuto: 'Automático',
  pickerAutoUsing: 'Automático · {provider}',
  pickerAutoUsingModel: 'Automático · {provider} · {model}',
  pickerDefaultModel: 'Modelo predeterminado',
  pickerDefaultModelNamed: 'Modelo predeterminado · {model}',
  pickerSubscription: 'Suscripción',
  pickerPaid: 'De pago',
  pickerPaidTitle: 'Se cobra por consumo, no lo cubre una suscripción',
  pickerNotConnected: 'Conectar IA',
  guideQuestion: '¿Qué IA usa ya?',
  guidePathClaudeTitle: 'Tengo una suscripción de pago a Claude',
  guidePathClaudeBody:
    'Necesita un plan Pro o Max. Una cuenta gratuita de Claude no puede conectarse.',
  guidePathCodexTitle: 'Uso ChatGPT',
  guidePathCodexBody:
    'Inicie sesión en el programa Codex de OpenAI con su cuenta de ChatGPT. El plan gratuito solo permite un uso reducido.',
  guidePathFreeTitle: 'Ninguna todavía',
  guidePathFreeBody:
    'Conecte una IA gratuita en un minuto aproximadamente. Mientras la use gratis, lo que envíe puede usarse para mejorar la IA de esa empresa.',
  guideBack: 'Elegir otra forma',
  guideBadgeReady: 'Conectado',
  guideBadgeInstalled: 'Instalado',
  guideBadgeCheck: 'Revisar',
  guideStepsLabel: 'Pasos de conexión',
  stepInstall: 'Instalar el programa',
  stepLogin: 'Iniciar sesión',
  stepReady: 'Conectado',
  installReturnNote:
    'Mientras descarga, esa ventana puede no mostrar nada durante unos minutos. No la cierre. Cuando termine la instalación, vuelva aquí. La aplicación vuelve a comprobarlo sola.',
  loginHowNote:
    'Al pulsarlo se abre un navegador. Inicie sesión con su cuenta de {vendor} y pulse Permitir; la conexión se completa sola.',
  signingInStalledNote:
    'Si no se abrió ninguna página de inicio de sesión en el navegador, abra la ventana de inicio de sesión con el botón de abajo. Si todavía está iniciando sesión en el navegador, mejor termine allí: el botón detiene ese inicio de sesión y empieza de nuevo.',
  actionLoginWindowOpen: 'Abrir la ventana de inicio de sesión',
  loginUnfinishedNote:
    'El inicio de sesión no terminó. Si el navegador muestra un código, hay que pegarlo en la ventana de inicio de sesión: ábrala y vuelva a iniciar sesión allí. Un código mostrado antes ya no sirve.',
  loginWindowTitle: 'Inicio de sesión en {provider}',
  loginWindowIntro:
    'Esta es la ventana de inicio de sesión del programa {provider}. El texto en inglés que aparece debajo es del programa {provider}.',
  loginWindowStepBrowser:
    'Cuando se abra en el navegador la página de inicio de sesión de {vendor}, inicie sesión y permita el acceso. Después, esta ventana se cierra sola.',
  loginWindowStepAddress:
    'Si no se abrió ningún navegador, seleccione la dirección de abajo que empieza por https:// arrastrando el puntero sobre ella y haga clic con el botón derecho para copiarla. Péguela en la barra de direcciones del navegador para abrir la página de inicio de sesión.',
  loginWindowStepCode:
    'Si el navegador muestra un código después de iniciar sesión, copie el código completo (use el botón de copiar si lo hay). Haga clic con el botón derecho una vez dentro de esta ventana para pegarlo y pulse Intro. Es normal que el código pegado no se vea.',
  loginWindowStepInvalid:
    'Si aparece "Invalid code", el código no se copió completo. No cierre esta ventana, vuelva a copiar el código y péguelo.',
  loginWindowClosing:
    'El inicio de sesión no terminó. Cuando haya leído los mensajes de arriba, pulse Intro para cerrar esta ventana. Para volver a intentarlo, pulse otra vez el botón de iniciar sesión.',
  actionCancelLogin: 'Cancelar inicio de sesión',
  actionUpdate: 'Actualizar',
  updateNote:
    'Al pulsarlo, la actualización oficial de {vendor} se ejecuta en una ventana nueva. Cuando termine allí, vuelva aquí.',
  guideInfoHint:
    '¿Sigue sin funcionar? Copie los datos de conexión y envíelos a quien le esté ayudando. No incluyen ninguna cuenta ni rutas de archivos.',
  actionCopyInfo: 'Copiar datos para pedir ayuda',
  copyInfoDone: 'Copiado',
  actionConnectHelp: 'Ver cada paso en pantalla',
  guidePathChatGptBody:
    'Inicie sesión con su cuenta de ChatGPT y use su plan. No hay que instalar nada. El plan gratuito solo permite un uso reducido.',
  guideUseCodex: '¿Ya usa el programa Codex en este PC? Conectar con Codex',
  guideUseDirect: 'Conectar directamente en la aplicación, sin Codex',
  chatgptLoginNote:
    'Se abre una página de inicio de sesión de ChatGPT en el navegador. Inicie sesión y permita que esta aplicación use su plan de ChatGPT; la conexión se completa sola.',
  chatgptLoginUnfinishedNote:
    'El inicio de sesión no terminó. Vuelva a pulsar para empezar de nuevo. Para conectar, tiene que permitir el uso de su plan de ChatGPT.',
  chatgptLoginNoPlanNote:
    'Inició sesión en ChatGPT, pero esta aplicación no recibió permiso para usar su plan. Vuelva a pulsar y permita el uso del plan en la página de autorización. Si ya lo permitió, puede que OpenAI no ofrezca este uso para su cuenta (por ejemplo, una gestionada por una empresa o un centro educativo) o para su plan. En ese caso puede conectarse con el programa Codex.',
  chatgptLoginDeniedNote:
    'Inició sesión en el navegador, pero OpenAI no aceptó esta conexión o no se pudo comprobar su respuesta. Compruebe que la fecha y la hora de este PC sean correctas y vuelva a pulsar para empezar de nuevo.',
  chatgptLoginNetworkNote:
    'No se pudo terminar el inicio de sesión: no se pudo contactar con OpenAI o no se abrió ningún navegador. Revise la conexión a internet (incluida la red de la empresa, una VPN o un proxy) y vuelva a pulsar.',
  chatgptLoginWorkspacePolicyNote:
    'Esta cuenta de ChatGPT pertenece a un espacio de trabajo gestionado por una empresa o un centro educativo, y OpenAI no permite que las aplicaciones usen el plan ChatGPT de un espacio de trabajo. El administrador del espacio de trabajo no tiene ningún ajuste para permitir esta aplicación, así que volver a iniciar sesión no servirá. Inicie sesión con una cuenta personal de ChatGPT (Go, Plus o Pro) o conéctese con la tarjeta de Codex.',
  chatgptLoginNotAllowedNote:
    'No permitió que esta aplicación use su plan de ChatGPT, o el navegador volvió del inicio de sesión con un error. Vuelva a pulsar, inicie sesión y permítalo en la página de autorización.',
  chatgptSigningInNote:
    'Inicie sesión con su cuenta de ChatGPT en el navegador que se abrió y permita que esta aplicación use su plan. La conexión se completa sola.',
  chatgptPlanNote:
    'La IA de esta aplicación funciona con su plan de ChatGPT. Al desconectar, solo se desconecta esta aplicación; los demás programas de este PC mantienen la sesión iniciada.',
  actionContinueWithChatGpt: 'Continuar con ChatGPT',
  actionManageUsage: 'Gestionar uso',
  browserPageDoneBody: 'Puede cerrar esta pestaña y volver a la aplicación.',
  browserPageFailedTitle: 'No se pudo completar la conexión',
  browserPageFailedBody: 'Vuelva a la aplicación e inténtelo de nuevo.',
  cardOff: 'Desactivado en esta aplicación',
  offNote:
    'Esta aplicación no ejecuta el programa {provider} ni comprueba su estado. La sesión de {provider} en este PC y su uso en una terminal siguen igual.',
  offElsewhereNote: 'Puede volver a usarlo en Configuración → IA.',
  offSwitchNote:
    '«No usar en esta aplicación» mantiene la sesión de {provider} en este PC y solo detiene esta aplicación. Cerrar sesión afecta a todo el PC.',
  actionTurnOff: 'No usar en esta aplicación',
  actionTurnOn: 'Volver a usar',
  connectOffTitle: '{provider} está desactivado en esta aplicación',
  connectOffBody:
    'Pulse «Volver a usar» en la tarjeta de {provider} en Configuración → IA, o elija otra IA.',
  errOffInApp:
    '{provider} está desactivado en esta aplicación, así que no se ejecutó nada. Puede volver a usarlo en Configuración → IA.',
  guideBadgeOff: 'Desactivado',
} satisfies Record<keyof typeof zh, string>
