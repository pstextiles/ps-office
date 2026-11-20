/** The build bundles the query refresh worker as its own chunk and returns its runtime path. */
declare module '*/query-refresh-worker?modulePath' {
  const path: string
  export default path
}
