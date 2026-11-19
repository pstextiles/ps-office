# Google Drive

Conecte Google Drive para abrir en GenOffice los archivos de Drive que elija y editarlos. Al guardar, el archivo vuelve a Drive. GenOffice solo puede abrir los archivos que usted elige en la ventana de Google; no ve el resto de su Drive.

## Conectar

1. Abra **Configuración ▸ Integraciones**.
2. Lea **Antes de usar archivos de Drive con la IA**, desactive el entrenamiento de modelos en las cuentas de IA que usa con archivos de Drive y luego marque esa opción.
   - Claude: en **Configuración ▸ Privacidad** de Claude, desactive el entrenamiento de modelos.
   - ChatGPT y Codex: en **Configuración ▸ Controles de datos** de ChatGPT, desactive **Mejorar el modelo para todos**.
   - Los proveedores con clave de API no usan los datos de la API para entrenamiento de forma predeterminada.
   - Las conexiones de IA gratuitas (modelos gratuitos de OpenRouter, Gemini) nunca reciben archivos de Drive.
3. Haga clic en **Conectar Google Drive**. La página de inicio de sesión de Google se abre en su navegador. Inicie sesión y permita el acceso.
4. Vuelva a GenOffice. La tarjeta muestra la cuenta conectada.

Google no permite que los archivos de Drive vayan a un servicio de IA que los use para entrenar sus modelos, así que no puede conectar ni elegir archivos hasta marcar la opción.

El inicio de sesión se guarda en este PC con el cifrado que su sistema operativo ofrece a su cuenta de usuario (la protección propia de Windows, el llavero de macOS o un llavero de escritorio en Linux). Si este PC no puede cifrarlo, la conexión solo dura hasta que cierre la aplicación.

## Abrir archivos de Drive

- Haga clic en **Google Drive** en la barra lateral de Inicio y luego en **Elegir de Google Drive**.
- El selector de archivos de Google se abre en su navegador. Elija los archivos que quiera abrir (puede elegir varios a la vez); GenOffice vuelve y los abre en el editor que corresponde.
- Los archivos que eligió una vez se quedan en la lista **Archivos elegidos**. Vuelva a abrirlos desde ahí o búsquelos por nombre y contenido.
- Los formatos que GenOffice no puede abrir aparecen atenuados. Los documentos, hojas de cálculo y presentaciones de Google se abren como archivos de Word, Excel y PowerPoint.

## Guardar de nuevo en Drive

- Guarde como siempre. Poco después, el archivo se sube a Drive.
- Si alguien cambió el archivo en Drive después de que lo abriera, se le pregunta si sobrescribir el archivo de Drive, guardar una copia en Drive o mantener el cambio solo en este PC.
- Un archivo que solo puede ver se guarda en Drive como copia o se queda en este PC.
- Un documento de Google se deja como está. Sus cambios se guardan como archivo de Office en la misma carpeta, o en Mi unidad si no se puede escribir en esa carpeta.
- Si una subida falla, sus cambios se quedan en este PC y GenOffice lo vuelve a intentar, también la próxima vez que se inicie la aplicación.

## La IA y Drive

- En el panel de IA de los editores de Word, Excel, PowerPoint, HTML, mapas mentales y Hangul, haga clic en **Adjuntar desde Google Drive**, junto al botón de adjuntar, para elegir archivos en la ventana de Google de su navegador y adjuntarlos como archivos de este PC. Hasta 20 a la vez.
- Mientras **Dejar que la IA busque y lea archivos en Drive** esté activado (se activa al conectar), la IA de cada editor puede buscar y leer los archivos de Drive que eligió en GenOffice. Por ejemplo: «Busque en mi Drive el acta de la reunión del mes pasado y resúmala». Si no encuentra el archivo, elíjalo primero en la pantalla de inicio.
- Lo que lee la IA va al proveedor de IA que usa. Las partes ocultas a la IA siguen ocultas.
- Si la IA que usa es una conexión gratuita, no se puede adjuntar desde Drive ni la IA lee Drive.
- La IA no puede modificar, eliminar ni compartir archivos de Drive.
- Cuando un chat contiene algo leído de Drive, ese chat deja de usar la búsqueda web, la búsqueda de imágenes, la lectura de páginas por dirección web y la inserción de imágenes por dirección web. Empiece un chat nuevo para volver a usarlas.

## Desconectar

Haga clic en **Desconectar** en **Configuración ▸ Integraciones**. GenOffice quita el inicio de sesión de este PC y pide a Google que retire su acceso. Las copias de archivos que ya están en este PC se conservan. También puede quitar el acceso en cualquier momento en **Conexiones de terceros** de su cuenta de Google.
