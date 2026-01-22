import parseDOCX from "../parseDOCX.js";
import parsePPTX from "../parsePPTX.js";
import { parseODT, parseODP, parseODS } from "../parseODF.js";

import { readFile } from "node:fs/promises";

const handlers = {
  "docx": parseDOCX,
  "pptx": parsePPTX,
  "odp": parseODP,
  "odt": parseODT,
  "ods": parseODS
};

const path = process.argv[2] || "";
const extension = path.split(".").at(-1);
const bytes = await readFile(path);
const html = await handlers[extension](bytes);

const output = `<div style="width: 50%; margin-left: 25%">${html}</div>`;
await Bun.write("output.html", output);
