/**
 * PS Office, refresh-only run:
 *
 *   <app binary> --refresh-queries <workbook.xlsx> [<workbook.xlsm> …]
 *
 * Refreshes each workbook's Power Query tables and recalculates its formulas
 * with the built-in engine (the sheets tab's Data › Refresh All), saves the
 * file, and exits — without opening the workbook on screen. It is for PCs
 * that are too slow for the full editor or whose Excel is too old to refresh
 * the queries; the installer puts a "PS Query Refresh" shortcut on the
 * desktop and in Explorer's Send to menu, so a file dropped on the icon (or
 * sent to it) is refreshed.
 *
 * One small window shows which file is being refreshed; at the end one
 * message lists each file's result. No single-instance lock is taken, so it
 * runs beside an open PS Office. Exit code 0 when every file was refreshed,
 * 1 for no files, 3 when any file was not.
 */

export const REFRESH_QUERIES_FLAG = '--refresh-queries'

export type RefreshQueriesArgv =
  | { readonly kind: 'none' }
  | { readonly kind: 'run'; readonly paths: readonly string[] }

/** the workbooks after the flag (Explorer appends dropped or sent files after the shortcut's own arguments) */
export function parseRefreshQueriesArgv(argv: readonly string[]): RefreshQueriesArgv {
  const index = argv.indexOf(REFRESH_QUERIES_FLAG)
  if (index < 0) return { kind: 'none' }
  const paths = argv.slice(index + 1).filter((arg) => arg.trim() !== '' && !arg.startsWith('--'))
  return { kind: 'run', paths: [...new Set(paths)] }
}

/** what the sheets module answers for one workbook (refreshQueriesHeadless) */
export type HeadlessRefreshResult =
  | {
      readonly ok: true
      readonly tables: readonly unknown[]
      readonly formulas: 'recalculated' | 'stale'
      readonly formulasDetail?: string
      readonly recalcOnOpen: boolean
    }
  | {
      readonly ok: false
      readonly code: 'busy' | 'limit' | 'failed' | 'locked' | 'timeout' | 'not-file'
      readonly query?: string
      readonly detail?: string
    }

export interface RefreshRunLabels {
  readonly title: string
  /** {index} of {total}: the file being refreshed */
  readonly running: string
  readonly patience: string
  readonly doneTitle: string
  readonly someFailedTitle: string
  readonly ok: string
  readonly okStale: string
  readonly limit: string
  readonly failed: string
  readonly locked: string
  readonly timeout: string
  readonly notFile: string
  readonly noFiles: string
  readonly button: string
  /** the standalone program's console: the last line before it closes */
  readonly pressEnter: string
}

const KO: RefreshRunLabels = {
  title: 'PS 쿼리 새로 고침',
  running: '쿼리를 새로 고치는 중… ({index}/{total})',
  patience: '느린 PC에서는 몇 분 걸릴 수 있습니다. 이 창을 닫지 마세요.',
  doneTitle: '새로 고침을 마쳤습니다',
  someFailedTitle: '새로 고치지 못한 파일이 있습니다',
  ok: '{file}: 쿼리 표 {tables}개를 새로 고치고 수식도 다시 계산해 저장했습니다.',
  okStale:
    '{file}: 쿼리 표 {tables}개를 새로 고쳐 저장했습니다. 일부 수식은 PS 오피스가 계산하지 못해 이전 결과로 남아 있으며, 최신 Excel에서 열면 다시 계산됩니다. ({detail})',
  limit:
    '{file}: 새로 고치지 않았습니다. 쿼리 {query}에 PS 오피스가 Excel과 똑같이 계산할 수 없는 부분이 있습니다. ({detail})',
  failed: '{file}: 새로 고치지 못했습니다. ({detail})',
  locked:
    '{file}: 다른 프로그램(Excel 등)에서 열려 있어 저장하지 못했습니다. 파일을 닫고 다시 해 주세요.',
  timeout: '{file}: 15분이 넘게 걸려 멈췄습니다. 파일은 그대로입니다.',
  notFile: '{file}: Excel 통합 문서(.xlsx, .xlsm)가 아닙니다.',
  noFiles:
    '사용법: 이 창을 닫은 뒤, 새로 고칠 Excel 파일을 끌어다 바탕 화면의 "PS Query Refresh" 아이콘이나 PSQueryRefresh.exe 파일 위에 놓으세요. 또는 Excel 파일을 마우스 오른쪽 단추로 누르고 [보내기] › [PS Query Refresh]를 고르세요.',
  button: '확인',
  pressEnter: 'Enter 키를 누르면 창이 닫힙니다.',
}

const EN: RefreshRunLabels = {
  title: 'PS Query Refresh',
  running: 'Refreshing queries… ({index}/{total})',
  patience: 'This can take a few minutes on a slow PC. Please keep this window open.',
  doneTitle: 'Refresh finished',
  someFailedTitle: 'Some files were not refreshed',
  ok: '{file}: {tables} query tables refreshed, formulas recalculated, saved.',
  okStale:
    '{file}: {tables} query tables refreshed and saved. Some formulas could not be calculated by PS Office and keep their previous results until the file is opened in a current Excel. ({detail})',
  limit:
    '{file}: not refreshed. The query {query} uses something PS Office cannot compute exactly as Excel does. ({detail})',
  failed: '{file}: not refreshed. ({detail})',
  locked:
    '{file}: not saved because it is open in another program (Excel?). Close it and try again.',
  timeout: '{file}: stopped after 15 minutes. The file was left as it was.',
  notFile: '{file}: not an Excel workbook (.xlsx, .xlsm).',
  noFiles:
    'How to use: close this window, then drag the Excel file to refresh and drop it on the "PS Query Refresh" icon on the desktop or on the PSQueryRefresh.exe file. Or right-click the Excel file and choose Send to › PS Query Refresh.',
  button: 'OK',
  pressEnter: 'Press Enter to close this window.',
}

const LABELS: Record<string, RefreshRunLabels> = {
  ko: KO,
  en: EN,
  es: {
    title: 'PS Query Refresh',
    running: 'Actualizando las consultas… ({index}/{total})',
    patience: 'En un equipo lento puede tardar unos minutos. No cierre esta ventana.',
    doneTitle: 'Actualización terminada',
    someFailedTitle: 'Algunos archivos no se actualizaron',
    ok: '{file}: se actualizaron {tables} tablas de consulta, se recalcularon las fórmulas y se guardó.',
    okStale:
      '{file}: se actualizaron y guardaron {tables} tablas de consulta. Algunas fórmulas no las pudo calcular PS Office y conservan su resultado anterior hasta que abra el archivo en un Excel actual. ({detail})',
    limit:
      '{file}: no se actualizó. La consulta {query} usa algo que PS Office no puede calcular exactamente como Excel. ({detail})',
    failed: '{file}: no se pudo actualizar. ({detail})',
    locked:
      '{file}: no se guardó porque está abierto en otro programa (¿Excel?). Ciérrelo e inténtelo de nuevo.',
    timeout: '{file}: se detuvo después de 15 minutos. El archivo quedó como estaba.',
    notFile: '{file}: no es un libro de Excel (.xlsx, .xlsm).',
    noFiles:
      'Cómo usarlo: cierre esta ventana y arrastre el archivo de Excel que desea actualizar hasta el icono "PS Query Refresh" del escritorio o hasta el archivo PSQueryRefresh.exe, y suéltelo allí. O haga clic derecho en el archivo de Excel y elija Enviar a › PS Query Refresh.',
    button: 'Aceptar',
    pressEnter: 'Presione Enter para cerrar esta ventana.',
  },
  zh: {
    title: 'PS 查询刷新',
    running: '正在刷新查询… ({index}/{total})',
    patience: '在较慢的电脑上可能需要几分钟。请不要关闭此窗口。',
    doneTitle: '刷新完成',
    someFailedTitle: '有些文件未能刷新',
    ok: '{file}：已刷新 {tables} 个查询表，重新计算公式并保存。',
    okStale:
      '{file}：已刷新并保存 {tables} 个查询表。部分公式 PS Office 无法计算，保留之前的结果，用新版 Excel 打开时会重新计算。({detail})',
    limit:
      '{file}：未刷新。查询 {query} 中有 PS Office 无法与 Excel 完全一致计算的部分。({detail})',
    failed: '{file}：未能刷新。({detail})',
    locked: '{file}：文件已在其他程序（Excel 等）中打开，未能保存。请关闭后重试。',
    timeout: '{file}：超过 15 分钟已停止。文件保持原样。',
    notFile: '{file}：不是 Excel 工作簿（.xlsx、.xlsm）。',
    noFiles:
      '用法：关闭此窗口，然后把要刷新的 Excel 文件拖到桌面上的“PS Query Refresh”图标或 PSQueryRefresh.exe 文件上松开。也可以右键单击 Excel 文件并选择“发送到” › PS Query Refresh。',
    button: '确定',
    pressEnter: '按 Enter 键关闭此窗口。',
  },
  'zh-TW': {
    title: 'PS 查詢重新整理',
    running: '正在重新整理查詢… ({index}/{total})',
    patience: '在較慢的電腦上可能需要幾分鐘。請不要關閉此視窗。',
    doneTitle: '重新整理完成',
    someFailedTitle: '有些檔案未能重新整理',
    ok: '{file}：已重新整理 {tables} 個查詢表，重新計算公式並儲存。',
    okStale:
      '{file}：已重新整理並儲存 {tables} 個查詢表。部分公式 PS Office 無法計算，保留先前的結果，以新版 Excel 開啟時會重新計算。({detail})',
    limit:
      '{file}：未重新整理。查詢 {query} 中有 PS Office 無法與 Excel 完全一致計算的部分。({detail})',
    failed: '{file}：未能重新整理。({detail})',
    locked: '{file}：檔案已在其他程式（Excel 等）中開啟，未能儲存。請關閉後再試一次。',
    timeout: '{file}：超過 15 分鐘已停止。檔案保持原樣。',
    notFile: '{file}：不是 Excel 活頁簿（.xlsx、.xlsm）。',
    noFiles:
      '用法：關閉此視窗，然後把要重新整理的 Excel 檔案拖曳到桌面上的「PS Query Refresh」圖示或 PSQueryRefresh.exe 檔案上放開。也可以在 Excel 檔案上按右鍵並選擇「傳送到」 › PS Query Refresh。',
    button: '確定',
    pressEnter: '按 Enter 鍵關閉此視窗。',
  },
  ja: {
    title: 'PS クエリ更新',
    running: 'クエリを更新しています… ({index}/{total})',
    patience: '遅い PC では数分かかることがあります。このウィンドウを閉じないでください。',
    doneTitle: '更新が終わりました',
    someFailedTitle: '更新できなかったファイルがあります',
    ok: '{file}: クエリ テーブル {tables} 個を更新し、数式を再計算して保存しました。',
    okStale:
      '{file}: クエリ テーブル {tables} 個を更新して保存しました。一部の数式は PS Office で計算できず以前の結果のままです。最新の Excel で開くと再計算されます。({detail})',
    limit:
      '{file}: 更新していません。クエリ {query} に、PS Office が Excel と同じに計算できない部分があります。({detail})',
    failed: '{file}: 更新できませんでした。({detail})',
    locked:
      '{file}: 他のプログラム (Excel など) で開かれているため保存できませんでした。閉じてからもう一度お試しください。',
    timeout: '{file}: 15 分を超えたため止めました。ファイルは元のままです。',
    notFile: '{file}: Excel ブック (.xlsx、.xlsm) ではありません。',
    noFiles:
      '使い方: このウィンドウを閉じてから、更新する Excel ファイルをデスクトップの「PS Query Refresh」アイコンか PSQueryRefresh.exe ファイルの上にドラッグして離してください。Excel ファイルを右クリックして [送る] › [PS Query Refresh] を選んでもかまいません。',
    button: 'OK',
    pressEnter: 'Enter キーを押すとこのウィンドウが閉じます。',
  },
  fr: {
    title: 'PS Query Refresh',
    running: 'Actualisation des requêtes… ({index}/{total})',
    patience: 'Sur un PC lent, cela peut prendre quelques minutes. Laissez cette fenêtre ouverte.',
    doneTitle: 'Actualisation terminée',
    someFailedTitle: 'Certains fichiers n’ont pas été actualisés',
    ok: '{file} : {tables} tables de requête actualisées, formules recalculées, enregistré.',
    okStale:
      '{file} : {tables} tables de requête actualisées et enregistrées. Certaines formules n’ont pas pu être calculées par PS Office et gardent leur résultat précédent jusqu’à l’ouverture dans un Excel récent. ({detail})',
    limit:
      '{file} : non actualisé. La requête {query} utilise quelque chose que PS Office ne peut pas calculer exactement comme Excel. ({detail})',
    failed: '{file} : actualisation impossible. ({detail})',
    locked:
      '{file} : non enregistré, car il est ouvert dans un autre programme (Excel ?). Fermez-le et réessayez.',
    timeout: '{file} : arrêté après 15 minutes. Le fichier n’a pas été modifié.',
    notFile: '{file} : ce n’est pas un classeur Excel (.xlsx, .xlsm).',
    noFiles:
      'Utilisation : fermez cette fenêtre, puis faites glisser le fichier Excel à actualiser sur l’icône « PS Query Refresh » du bureau ou sur le fichier PSQueryRefresh.exe. Vous pouvez aussi cliquer avec le bouton droit sur le fichier Excel et choisir Envoyer vers › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Appuyez sur Entrée pour fermer cette fenêtre.',
  },
  de: {
    title: 'PS Query Refresh',
    running: 'Abfragen werden aktualisiert… ({index}/{total})',
    patience:
      'Auf einem langsamen PC kann das einige Minuten dauern. Bitte lassen Sie dieses Fenster offen.',
    doneTitle: 'Aktualisierung abgeschlossen',
    someFailedTitle: 'Einige Dateien wurden nicht aktualisiert',
    ok: '{file}: {tables} Abfragetabellen aktualisiert, Formeln neu berechnet, gespeichert.',
    okStale:
      '{file}: {tables} Abfragetabellen aktualisiert und gespeichert. Einige Formeln konnte PS Office nicht berechnen; sie behalten ihr bisheriges Ergebnis, bis die Datei in einem aktuellen Excel geöffnet wird. ({detail})',
    limit:
      '{file}: nicht aktualisiert. Die Abfrage {query} verwendet etwas, das PS Office nicht genau wie Excel berechnen kann. ({detail})',
    failed: '{file}: Aktualisierung fehlgeschlagen. ({detail})',
    locked:
      '{file}: nicht gespeichert, weil die Datei in einem anderen Programm (Excel?) geöffnet ist. Schließen Sie sie und versuchen Sie es erneut.',
    timeout: '{file}: nach 15 Minuten abgebrochen. Die Datei bleibt unverändert.',
    notFile: '{file}: keine Excel-Arbeitsmappe (.xlsx, .xlsm).',
    noFiles:
      'So geht es: Schließen Sie dieses Fenster und ziehen Sie die zu aktualisierende Excel-Datei auf das Symbol „PS Query Refresh“ auf dem Desktop oder auf die Datei PSQueryRefresh.exe. Oder klicken Sie mit der rechten Maustaste auf die Excel-Datei und wählen Sie Senden an › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Drücken Sie die Eingabetaste, um dieses Fenster zu schließen.',
  },
  pt: {
    title: 'PS Query Refresh',
    running: 'Atualizando as consultas… ({index}/{total})',
    patience: 'Em um PC lento, isso pode levar alguns minutos. Não feche esta janela.',
    doneTitle: 'Atualização concluída',
    someFailedTitle: 'Alguns arquivos não foram atualizados',
    ok: '{file}: {tables} tabelas de consulta atualizadas, fórmulas recalculadas, salvo.',
    okStale:
      '{file}: {tables} tabelas de consulta atualizadas e salvas. Algumas fórmulas não puderam ser calculadas pelo PS Office e mantêm o resultado anterior até o arquivo ser aberto em um Excel atual. ({detail})',
    limit:
      '{file}: não atualizado. A consulta {query} usa algo que o PS Office não consegue calcular exatamente como o Excel. ({detail})',
    failed: '{file}: não foi possível atualizar. ({detail})',
    locked:
      '{file}: não foi salvo porque está aberto em outro programa (Excel?). Feche-o e tente novamente.',
    timeout: '{file}: interrompido após 15 minutos. O arquivo ficou como estava.',
    notFile: '{file}: não é uma pasta de trabalho do Excel (.xlsx, .xlsm).',
    noFiles:
      'Como usar: feche esta janela e arraste o arquivo do Excel que deseja atualizar até o ícone "PS Query Refresh" na área de trabalho ou até o arquivo PSQueryRefresh.exe. Ou clique com o botão direito no arquivo do Excel e escolha Enviar para › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Pressione Enter para fechar esta janela.',
  },
  it: {
    title: 'PS Query Refresh',
    running: 'Aggiornamento delle query… ({index}/{total})',
    patience: 'Su un PC lento può richiedere alcuni minuti. Non chiudere questa finestra.',
    doneTitle: 'Aggiornamento completato',
    someFailedTitle: 'Alcuni file non sono stati aggiornati',
    ok: '{file}: {tables} tabelle di query aggiornate, formule ricalcolate, salvato.',
    okStale:
      '{file}: {tables} tabelle di query aggiornate e salvate. Alcune formule non sono state calcolate da PS Office e mantengono il risultato precedente finché il file non viene aperto in un Excel recente. ({detail})',
    limit:
      '{file}: non aggiornato. La query {query} usa qualcosa che PS Office non può calcolare esattamente come Excel. ({detail})',
    failed: '{file}: aggiornamento non riuscito. ({detail})',
    locked:
      '{file}: non salvato perché è aperto in un altro programma (Excel?). Chiudilo e riprova.',
    timeout: '{file}: interrotto dopo 15 minuti. Il file è rimasto com’era.',
    notFile: '{file}: non è una cartella di lavoro di Excel (.xlsx, .xlsm).',
    noFiles:
      'Come si usa: chiudi questa finestra, poi trascina il file Excel da aggiornare sull’icona "PS Query Refresh" del desktop o sul file PSQueryRefresh.exe. Oppure fai clic con il pulsante destro sul file Excel e scegli Invia a › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Premi Invio per chiudere questa finestra.',
  },
  ru: {
    title: 'PS Query Refresh',
    running: 'Обновление запросов… ({index}/{total})',
    patience: 'На медленном компьютере это может занять несколько минут. Не закрывайте это окно.',
    doneTitle: 'Обновление завершено',
    someFailedTitle: 'Некоторые файлы не обновлены',
    ok: '{file}: обновлено таблиц запросов: {tables}, формулы пересчитаны, файл сохранён.',
    okStale:
      '{file}: обновлено и сохранено таблиц запросов: {tables}. Некоторые формулы PS Office вычислить не смог; они сохраняют прежний результат, пока файл не откроют в современной версии Excel. ({detail})',
    limit:
      '{file}: не обновлён. Запрос {query} использует то, что PS Office не может вычислить точно так же, как Excel. ({detail})',
    failed: '{file}: обновить не удалось. ({detail})',
    locked:
      '{file}: не сохранён, потому что открыт в другой программе (Excel?). Закройте его и повторите попытку.',
    timeout: '{file}: остановлено через 15 минут. Файл не изменён.',
    notFile: '{file}: это не книга Excel (.xlsx, .xlsm).',
    noFiles:
      'Как пользоваться: закройте это окно и перетащите файл Excel, который нужно обновить, на значок «PS Query Refresh» на рабочем столе или на файл PSQueryRefresh.exe. Или щёлкните файл Excel правой кнопкой мыши и выберите «Отправить» › PS Query Refresh.',
    button: 'ОК',
    pressEnter: 'Нажмите Enter, чтобы закрыть это окно.',
  },
  pl: {
    title: 'PS Query Refresh',
    running: 'Odświeżanie zapytań… ({index}/{total})',
    patience: 'Na wolnym komputerze może to potrwać kilka minut. Nie zamykaj tego okna.',
    doneTitle: 'Odświeżanie zakończone',
    someFailedTitle: 'Niektóre pliki nie zostały odświeżone',
    ok: '{file}: odświeżono tabele zapytań ({tables}), przeliczono formuły, zapisano.',
    okStale:
      '{file}: odświeżono i zapisano tabele zapytań ({tables}). Niektórych formuł PS Office nie mógł obliczyć; zachowują poprzedni wynik do czasu otwarcia pliku w aktualnym Excelu. ({detail})',
    limit:
      '{file}: nie odświeżono. Zapytanie {query} używa czegoś, czego PS Office nie potrafi obliczyć dokładnie tak jak Excel. ({detail})',
    failed: '{file}: nie udało się odświeżyć. ({detail})',
    locked:
      '{file}: nie zapisano, ponieważ plik jest otwarty w innym programie (Excel?). Zamknij go i spróbuj ponownie.',
    timeout: '{file}: zatrzymano po 15 minutach. Plik pozostał bez zmian.',
    notFile: '{file}: to nie jest skoroszyt programu Excel (.xlsx, .xlsm).',
    noFiles:
      'Jak używać: zamknij to okno, a następnie przeciągnij plik Excel do odświeżenia na ikonę „PS Query Refresh” na pulpicie lub na plik PSQueryRefresh.exe. Możesz też kliknąć plik Excel prawym przyciskiem myszy i wybrać Wyślij do › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Naciśnij Enter, aby zamknąć to okno.',
  },
  cs: {
    title: 'PS Query Refresh',
    running: 'Aktualizují se dotazy… ({index}/{total})',
    patience: 'Na pomalém počítači to může trvat několik minut. Nezavírejte toto okno.',
    doneTitle: 'Aktualizace dokončena',
    someFailedTitle: 'Některé soubory se nepodařilo aktualizovat',
    ok: '{file}: aktualizováno tabulek dotazů: {tables}, vzorce přepočítány, uloženo.',
    okStale:
      '{file}: aktualizováno a uloženo tabulek dotazů: {tables}. Některé vzorce PS Office nedokázal spočítat; ponechávají předchozí výsledek, dokud soubor neotevřete v aktuálním Excelu. ({detail})',
    limit:
      '{file}: neaktualizováno. Dotaz {query} používá něco, co PS Office nedokáže spočítat přesně jako Excel. ({detail})',
    failed: '{file}: aktualizace se nezdařila. ({detail})',
    locked:
      '{file}: neuloženo, protože je otevřený v jiném programu (Excel?). Zavřete ho a zkuste to znovu.',
    timeout: '{file}: zastaveno po 15 minutách. Soubor zůstal beze změny.',
    notFile: '{file}: není sešit Excelu (.xlsx, .xlsm).',
    noFiles:
      'Jak na to: zavřete toto okno a přetáhněte soubor Excelu, který chcete aktualizovat, na ikonu „PS Query Refresh“ na ploše nebo na soubor PSQueryRefresh.exe. Nebo na soubor Excelu klikněte pravým tlačítkem a zvolte Odeslat › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Stisknutím klávesy Enter toto okno zavřete.',
  },
  nl: {
    title: 'PS Query Refresh',
    running: 'Query’s vernieuwen… ({index}/{total})',
    patience: 'Op een trage pc kan dit enkele minuten duren. Laat dit venster open.',
    doneTitle: 'Vernieuwen voltooid',
    someFailedTitle: 'Sommige bestanden zijn niet vernieuwd',
    ok: '{file}: {tables} querytabellen vernieuwd, formules opnieuw berekend, opgeslagen.',
    okStale:
      '{file}: {tables} querytabellen vernieuwd en opgeslagen. Sommige formules kon PS Office niet berekenen; ze houden hun vorige resultaat tot het bestand in een recente Excel wordt geopend. ({detail})',
    limit:
      '{file}: niet vernieuwd. De query {query} gebruikt iets wat PS Office niet precies zoals Excel kan berekenen. ({detail})',
    failed: '{file}: vernieuwen is mislukt. ({detail})',
    locked:
      '{file}: niet opgeslagen omdat het in een ander programma (Excel?) is geopend. Sluit het en probeer het opnieuw.',
    timeout: '{file}: gestopt na 15 minuten. Het bestand is ongewijzigd.',
    notFile: '{file}: geen Excel-werkmap (.xlsx, .xlsm).',
    noFiles:
      'Zo werkt het: sluit dit venster en sleep het Excel-bestand dat u wilt vernieuwen naar het pictogram "PS Query Refresh" op het bureaublad of naar het bestand PSQueryRefresh.exe. Of klik met de rechtermuisknop op het Excel-bestand en kies Kopiëren naar › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Druk op Enter om dit venster te sluiten.',
  },
  tr: {
    title: 'PS Query Refresh',
    running: 'Sorgular yenileniyor… ({index}/{total})',
    patience: 'Yavaş bir bilgisayarda birkaç dakika sürebilir. Lütfen bu pencereyi kapatmayın.',
    doneTitle: 'Yenileme tamamlandı',
    someFailedTitle: 'Bazı dosyalar yenilenmedi',
    ok: '{file}: {tables} sorgu tablosu yenilendi, formüller yeniden hesaplandı, kaydedildi.',
    okStale:
      '{file}: {tables} sorgu tablosu yenilendi ve kaydedildi. Bazı formülleri PS Office hesaplayamadı; dosya güncel bir Excel’de açılana kadar önceki sonuçlarını korurlar. ({detail})',
    limit:
      '{file}: yenilenmedi. {query} sorgusu, PS Office’in Excel ile tam olarak aynı şekilde hesaplayamadığı bir şey kullanıyor. ({detail})',
    failed: '{file}: yenilenemedi. ({detail})',
    locked:
      '{file}: başka bir programda (Excel?) açık olduğu için kaydedilmedi. Kapatıp yeniden deneyin.',
    timeout: '{file}: 15 dakikadan sonra durduruldu. Dosya olduğu gibi bırakıldı.',
    notFile: '{file}: bir Excel çalışma kitabı (.xlsx, .xlsm) değil.',
    noFiles:
      'Kullanım: Bu pencereyi kapatın, sonra yenilenecek Excel dosyasını masaüstündeki "PS Query Refresh" simgesinin ya da PSQueryRefresh.exe dosyasının üzerine sürükleyip bırakın. Ya da Excel dosyasına sağ tıklayıp Gönder › PS Query Refresh’i seçin.',
    button: 'Tamam',
    pressEnter: 'Bu pencereyi kapatmak için Enter tuşuna basın.',
  },
  th: {
    title: 'PS Query Refresh',
    running: 'กำลังรีเฟรชคิวรี… ({index}/{total})',
    patience: 'บนเครื่องที่ช้าอาจใช้เวลาหลายนาที โปรดอย่าปิดหน้าต่างนี้',
    doneTitle: 'รีเฟรชเสร็จแล้ว',
    someFailedTitle: 'มีบางไฟล์ที่รีเฟรชไม่สำเร็จ',
    ok: '{file}: รีเฟรชตารางคิวรี {tables} ตาราง คำนวณสูตรใหม่ และบันทึกแล้ว',
    okStale:
      '{file}: รีเฟรชและบันทึกตารางคิวรี {tables} ตารางแล้ว สูตรบางสูตร PS Office คำนวณไม่ได้ จึงยังเป็นผลลัพธ์เดิมจนกว่าจะเปิดไฟล์ใน Excel รุ่นปัจจุบัน ({detail})',
    limit:
      '{file}: ไม่ได้รีเฟรช คิวรี {query} ใช้สิ่งที่ PS Office ไม่สามารถคำนวณให้ตรงกับ Excel ได้ ({detail})',
    failed: '{file}: รีเฟรชไม่สำเร็จ ({detail})',
    locked:
      '{file}: บันทึกไม่ได้เพราะไฟล์เปิดอยู่ในโปรแกรมอื่น (Excel?) โปรดปิดไฟล์แล้วลองอีกครั้ง',
    timeout: '{file}: หยุดหลังจาก 15 นาที ไฟล์ยังคงเหมือนเดิม',
    notFile: '{file}: ไม่ใช่เวิร์กบุ๊ก Excel (.xlsx, .xlsm)',
    noFiles:
      'วิธีใช้: ปิดหน้าต่างนี้ แล้วลากไฟล์ Excel ที่ต้องการรีเฟรชไปวางบนไอคอน "PS Query Refresh" บนเดสก์ท็อป หรือบนไฟล์ PSQueryRefresh.exe หรือคลิกขวาที่ไฟล์ Excel แล้วเลือก ส่งไปยัง › PS Query Refresh',
    button: 'ตกลง',
    pressEnter: 'กด Enter เพื่อปิดหน้าต่างนี้',
  },
  id: {
    title: 'PS Query Refresh',
    running: 'Menyegarkan kueri… ({index}/{total})',
    patience: 'Di PC yang lambat ini bisa memakan waktu beberapa menit. Jangan tutup jendela ini.',
    doneTitle: 'Penyegaran selesai',
    someFailedTitle: 'Beberapa file tidak disegarkan',
    ok: '{file}: {tables} tabel kueri disegarkan, rumus dihitung ulang, disimpan.',
    okStale:
      '{file}: {tables} tabel kueri disegarkan dan disimpan. Beberapa rumus tidak dapat dihitung oleh PS Office dan tetap memakai hasil sebelumnya sampai file dibuka di Excel versi terbaru. ({detail})',
    limit:
      '{file}: tidak disegarkan. Kueri {query} memakai sesuatu yang tidak dapat dihitung PS Office persis seperti Excel. ({detail})',
    failed: '{file}: gagal disegarkan. ({detail})',
    locked:
      '{file}: tidak disimpan karena sedang dibuka di program lain (Excel?). Tutup lalu coba lagi.',
    timeout: '{file}: dihentikan setelah 15 menit. File tetap seperti semula.',
    notFile: '{file}: bukan buku kerja Excel (.xlsx, .xlsm).',
    noFiles:
      'Cara pakai: tutup jendela ini, lalu seret file Excel yang ingin disegarkan ke ikon "PS Query Refresh" di desktop atau ke file PSQueryRefresh.exe. Atau klik kanan file Excel lalu pilih Kirim ke › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Tekan Enter untuk menutup jendela ini.',
  },
  ms: {
    title: 'PS Query Refresh',
    running: 'Menyegar semula pertanyaan… ({index}/{total})',
    patience:
      'Pada PC yang perlahan, ini boleh mengambil masa beberapa minit. Jangan tutup tetingkap ini.',
    doneTitle: 'Segar semula selesai',
    someFailedTitle: 'Sesetengah fail tidak disegar semula',
    ok: '{file}: {tables} jadual pertanyaan disegar semula, formula dikira semula, disimpan.',
    okStale:
      '{file}: {tables} jadual pertanyaan disegar semula dan disimpan. Sesetengah formula tidak dapat dikira oleh PS Office dan mengekalkan hasil sebelumnya sehingga fail dibuka dalam Excel terkini. ({detail})',
    limit:
      '{file}: tidak disegar semula. Pertanyaan {query} menggunakan sesuatu yang tidak dapat dikira oleh PS Office sama seperti Excel. ({detail})',
    failed: '{file}: gagal disegar semula. ({detail})',
    locked:
      '{file}: tidak disimpan kerana sedang dibuka dalam program lain (Excel?). Tutup dan cuba lagi.',
    timeout: '{file}: dihentikan selepas 15 minit. Fail kekal seperti asal.',
    notFile: '{file}: bukan buku kerja Excel (.xlsx, .xlsm).',
    noFiles:
      'Cara guna: tutup tetingkap ini, kemudian seret fail Excel yang hendak disegar semula ke ikon "PS Query Refresh" di desktop atau ke fail PSQueryRefresh.exe. Atau klik kanan fail Excel dan pilih Hantar ke › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Tekan Enter untuk menutup tetingkap ini.',
  },
  vi: {
    title: 'PS Query Refresh',
    running: 'Đang làm mới truy vấn… ({index}/{total})',
    patience: 'Trên máy chậm, việc này có thể mất vài phút. Vui lòng không đóng cửa sổ này.',
    doneTitle: 'Đã làm mới xong',
    someFailedTitle: 'Một số tệp chưa được làm mới',
    ok: '{file}: đã làm mới {tables} bảng truy vấn, tính lại công thức và lưu.',
    okStale:
      '{file}: đã làm mới và lưu {tables} bảng truy vấn. Một số công thức PS Office không tính được nên giữ kết quả cũ cho đến khi mở tệp bằng Excel bản mới. ({detail})',
    limit:
      '{file}: chưa làm mới. Truy vấn {query} dùng thứ mà PS Office không thể tính giống hệt Excel. ({detail})',
    failed: '{file}: không làm mới được. ({detail})',
    locked:
      '{file}: chưa lưu vì tệp đang mở trong chương trình khác (Excel?). Hãy đóng tệp rồi thử lại.',
    timeout: '{file}: đã dừng sau 15 phút. Tệp vẫn giữ nguyên.',
    notFile: '{file}: không phải sổ làm việc Excel (.xlsx, .xlsm).',
    noFiles:
      'Cách dùng: đóng cửa sổ này, rồi kéo tệp Excel cần làm mới thả vào biểu tượng "PS Query Refresh" trên màn hình nền hoặc vào tệp PSQueryRefresh.exe. Hoặc nhấp chuột phải vào tệp Excel và chọn Gửi tới › PS Query Refresh.',
    button: 'OK',
    pressEnter: 'Nhấn Enter để đóng cửa sổ này.',
  },
  ar: {
    title: 'PS Query Refresh',
    running: 'جارٍ تحديث الاستعلامات… ({index}/{total})',
    patience: 'قد يستغرق ذلك بضع دقائق على جهاز بطيء. يُرجى عدم إغلاق هذه النافذة.',
    doneTitle: 'اكتمل التحديث',
    someFailedTitle: 'لم يتم تحديث بعض الملفات',
    ok: '{file}: تم تحديث {tables} من جداول الاستعلام وإعادة حساب الصيغ والحفظ.',
    okStale:
      '{file}: تم تحديث {tables} من جداول الاستعلام وحفظها. تعذّر على PS Office حساب بعض الصيغ فاحتفظت بنتائجها السابقة إلى أن يُفتح الملف في إصدار حديث من Excel. ({detail})',
    limit:
      '{file}: لم يتم التحديث. يستخدم الاستعلام {query} شيئًا لا يستطيع PS Office حسابه تمامًا كما يفعل Excel. ({detail})',
    failed: '{file}: تعذّر التحديث. ({detail})',
    locked: '{file}: لم يُحفظ لأنه مفتوح في برنامج آخر (Excel؟). أغلقه ثم حاول مرة أخرى.',
    timeout: '{file}: توقف بعد 15 دقيقة. بقي الملف كما كان.',
    notFile: '{file}: ليس مصنف Excel (.xlsx أو .xlsm).',
    noFiles:
      'طريقة الاستخدام: أغلق هذه النافذة، ثم اسحب ملف Excel المراد تحديثه وأفلته على أيقونة "PS Query Refresh" على سطح المكتب أو على الملف PSQueryRefresh.exe. أو انقر بزر الماوس الأيمن على ملف Excel واختر إرسال إلى › PS Query Refresh.',
    button: 'موافق',
    pressEnter: 'اضغط Enter لإغلاق هذه النافذة.',
  },
  he: {
    title: 'PS Query Refresh',
    running: 'מרענן שאילתות… ({index}/{total})',
    patience: 'במחשב איטי זה עשוי להימשך כמה דקות. נא לא לסגור חלון זה.',
    doneTitle: 'הרענון הסתיים',
    someFailedTitle: 'חלק מהקבצים לא רועננו',
    ok: '{file}: רועננו {tables} טבלאות שאילתה, הנוסחאות חושבו מחדש והקובץ נשמר.',
    okStale:
      '{file}: רועננו ונשמרו {tables} טבלאות שאילתה. חלק מהנוסחאות PS Office לא הצליח לחשב, והן שומרות על התוצאה הקודמת עד שהקובץ ייפתח ב-Excel עדכני. ({detail})',
    limit:
      '{file}: לא רוענן. השאילתה {query} משתמשת במשהו ש-PS Office לא יכול לחשב בדיוק כמו Excel. ({detail})',
    failed: '{file}: הרענון נכשל. ({detail})',
    locked: '{file}: לא נשמר כי הוא פתוח בתוכנה אחרת (Excel?). סגרו אותו ונסו שוב.',
    timeout: '{file}: נעצר אחרי 15 דקות. הקובץ נשאר כפי שהיה.',
    notFile: '{file}: אינו חוברת עבודה של Excel ‏(‎.xlsx, .xlsm‏).',
    noFiles:
      'אופן השימוש: סגרו חלון זה, ואז גררו את קובץ ה-Excel לרענון ושחררו אותו על הסמל "PS Query Refresh" בשולחן העבודה או על הקובץ PSQueryRefresh.exe. אפשר גם ללחוץ לחיצה ימנית על קובץ ה-Excel ולבחור שלח אל › PS Query Refresh.',
    button: 'אישור',
    pressEnter: 'הקישו Enter כדי לסגור חלון זה.',
  },
  hi: {
    title: 'PS Query Refresh',
    running: 'क्वेरी रीफ़्रेश हो रही हैं… ({index}/{total})',
    patience: 'धीमे पीसी पर इसमें कुछ मिनट लग सकते हैं। कृपया यह विंडो बंद न करें।',
    doneTitle: 'रीफ़्रेश पूरा हुआ',
    someFailedTitle: 'कुछ फ़ाइलें रीफ़्रेश नहीं हुईं',
    ok: '{file}: {tables} क्वेरी तालिकाएँ रीफ़्रेश हुईं, सूत्र फिर से गणना किए गए, सहेजा गया।',
    okStale:
      '{file}: {tables} क्वेरी तालिकाएँ रीफ़्रेश करके सहेजी गईं। कुछ सूत्रों की गणना PS Office नहीं कर सका; नए Excel में फ़ाइल खुलने तक उनके पिछले परिणाम बने रहेंगे। ({detail})',
    limit:
      '{file}: रीफ़्रेश नहीं हुई। क्वेरी {query} में कुछ ऐसा है जिसकी गणना PS Office ठीक Excel की तरह नहीं कर सकता। ({detail})',
    failed: '{file}: रीफ़्रेश नहीं हो सकी। ({detail})',
    locked:
      '{file}: सहेजी नहीं गई, क्योंकि यह किसी दूसरे प्रोग्राम (Excel?) में खुली है। इसे बंद करके फिर कोशिश करें।',
    timeout: '{file}: 15 मिनट बाद रोक दिया गया। फ़ाइल पहले जैसी ही है।',
    notFile: '{file}: Excel वर्कबुक (.xlsx, .xlsm) नहीं है।',
    noFiles:
      'उपयोग का तरीका: यह विंडो बंद करें, फिर रीफ़्रेश करने वाली Excel फ़ाइल को डेस्कटॉप पर "PS Query Refresh" आइकन या PSQueryRefresh.exe फ़ाइल पर खींचकर छोड़ें। या Excel फ़ाइल पर राइट-क्लिक करके भेजें › PS Query Refresh चुनें।',
    button: 'ठीक है',
    pressEnter: 'यह विंडो बंद करने के लिए Enter दबाएँ।',
  },
}

export function refreshRunLabels(lang: string): RefreshRunLabels {
  return LABELS[lang] ?? LABELS[lang.split('-')[0] ?? ''] ?? EN
}

function fill(template: string, params: Readonly<Record<string, string | number>>): string {
  return template.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in params ? String(params[key]) : match,
  )
}

/** one line per workbook: what happened to it */
export function describeRefreshResult(
  fileName: string,
  result: HeadlessRefreshResult,
  labels: RefreshRunLabels,
): string {
  if (result.ok) {
    const params = { file: fileName, tables: result.tables.length }
    if (result.formulas === 'recalculated') return fill(labels.ok, params)
    return fill(labels.okStale, { ...params, detail: result.formulasDetail ?? '' }).replace(
      / ?\(\)$/,
      '',
    )
  }
  const detail = result.detail ?? ''
  const strip = (text: string): string => text.replace(/ ?\(\)$/, '')
  switch (result.code) {
    case 'limit':
      return strip(fill(labels.limit, { file: fileName, query: result.query ?? '', detail }))
    case 'locked':
      return fill(labels.locked, { file: fileName })
    case 'timeout':
      return fill(labels.timeout, { file: fileName })
    case 'not-file':
      return fill(labels.notFile, { file: fileName })
    default:
      return strip(fill(labels.failed, { file: fileName, detail }))
  }
}

export interface RefreshRunSummary {
  readonly type: 'info' | 'warning' | 'error'
  readonly title: string
  readonly message: string
  readonly exitCode: 0 | 1 | 3
}

/** the closing message: every file's line, the worst outcome deciding its look */
export function summarizeRefreshRun(
  outcomes: readonly { readonly fileName: string; readonly result: HeadlessRefreshResult }[],
  labels: RefreshRunLabels,
): RefreshRunSummary {
  if (outcomes.length === 0) {
    return { type: 'info', title: labels.title, message: labels.noFiles, exitCode: 1 }
  }
  const lines = outcomes.map((outcome) =>
    describeRefreshResult(outcome.fileName, outcome.result, labels),
  )
  const failed = outcomes.some((outcome) => !outcome.result.ok)
  const stale = outcomes.some((outcome) => outcome.result.ok && outcome.result.formulas === 'stale')
  return {
    type: failed ? 'error' : stale ? 'warning' : 'info',
    title: failed ? labels.someFailedTitle : labels.doneTitle,
    message: lines.join('\n\n'),
    exitCode: failed ? 3 : 0,
  }
}

/** the run as one JSON line on stdout: each file's result as the engine gave it (no table list) */
export function refreshRunReport(
  outcomes: readonly { readonly fileName: string; readonly result: HeadlessRefreshResult }[],
): {
  readonly status: 'ok' | 'error'
  readonly files: readonly Record<string, unknown>[]
} {
  return {
    status: outcomes.length > 0 && outcomes.every((outcome) => outcome.result.ok) ? 'ok' : 'error',
    files: outcomes.map(({ fileName, result }) =>
      result.ok
        ? {
            file: fileName,
            ok: true,
            tables: result.tables.length,
            formulas: result.formulas,
            ...(result.formulasDetail ? { formulasDetail: result.formulasDetail } : {}),
            recalcOnOpen: result.recalcOnOpen,
          }
        : { file: fileName, ...result },
    ),
  }
}

/** the progress window's page for one file: which one, a moving bar, in the system's light or dark colors */
export function progressPageHtml(
  labels: RefreshRunLabels,
  progress: { readonly index: number; readonly total: number; readonly fileName: string },
  rtl: boolean,
): string {
  const escape = (text: string): string =>
    text.replace(
      /[&<>"]/g,
      (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]!,
    )
  const status = fill(labels.running, { index: progress.index, total: progress.total })
  return `<!doctype html><html dir="${rtl ? 'rtl' : 'ltr'}"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">
<title>${escape(labels.title)}</title>
<style>
:root{color-scheme:light dark;--bg:#ffffff;--fg:#1f1f1f;--muted:#5f6368;--track:#e3e3e3;--bar:#1a73e8}
@media (prefers-color-scheme:dark){:root{--bg:#202124;--fg:#e8eaed;--muted:#9aa0a6;--track:#3c4043;--bar:#8ab4f8}}
html,body{margin:0;height:100%;background:var(--bg);color:var(--fg);font:14px "Segoe UI","Malgun Gothic",system-ui,sans-serif}
main{padding:18px 20px;display:flex;flex-direction:column;gap:10px}
.status{font-weight:600}
.file{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hint{color:var(--muted);font-size:12px}
.track{height:4px;border-radius:2px;background:var(--track);overflow:hidden;position:relative}
.bar{position:absolute;top:0;bottom:0;width:30%;background:var(--bar);border-radius:2px;animation:move 1.4s ease-in-out infinite}
@keyframes move{0%{left:-30%}100%{left:100%}}
@media (prefers-reduced-motion:reduce){.bar{animation:none;left:0;width:100%;opacity:.5}}
</style></head><body><main>
<div class="status">${escape(status)}</div>
<div class="file">${escape(progress.fileName)}</div>
<div class="track"><div class="bar"></div></div>
<div class="hint">${escape(labels.patience)}</div>
</main></body></html>`
}
