import parseDOCX from "../parseDOCX.js";
import { parseODT, parseODP } from "../parseODF.js";

import { readFile } from "node:fs/promises";

const handlers = {
  "docx": parseDOCX,
  "odp": parseODP,
  "odt": parseODT
};

const path = process.argv[2] || "";
const extension = path.split(".").at(-1);
const bytes = await readFile(path);
const html = await handlers[extension](bytes);

const output = `<div style="width: 50%; margin-left: 25%">${html}</div>`;
await Bun.write("output.html", output);
