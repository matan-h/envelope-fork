## Envelope
Lightweight office document-to-HTML parser written in pure JavaScript. Compatible with both Microsoft and OpenDocument formats. Designed to run in a browser context out of the box.

Currently supports:
- [x] DOCX
- [x] PPTX _(partially)_
- [x] XLSX _(partially)_
- [x] ODT
- [x] ODP _(partially)_
- [x] ODS

Not planned:
- DOC
- PPT
- XLS

## Usage
```js
import parseDOCX from "./parseDOCX.js";
// OR: import { parseODT } from "../parseODF.js";

const getHTML = async (bytes) => {
  return await parseDOCX(bytes);
}

const bytes = u8array; // Uint8Array containing DOCX file data
getHTML(bytes); // Returns Promise to HTML string
```
See `test/test.js` for a practical example.

## Contributing
Not yet.
