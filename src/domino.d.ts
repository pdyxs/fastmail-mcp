// @mixmark-io/domino ships typings for the bare 'domino' module name only.
declare module '@mixmark-io/domino' {
  const domino: { createDocument(html?: string, force?: boolean): Document };
  export default domino;
}
