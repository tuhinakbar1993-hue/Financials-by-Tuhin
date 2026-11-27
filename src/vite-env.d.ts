/// <reference types="vite/client" />

declare module 'xlsx-js-style' {
  import type * as XLSX from 'xlsx';
  const XLSXStyle: typeof XLSX;
  export default XLSXStyle;
}
