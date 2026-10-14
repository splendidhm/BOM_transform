// Spreadsheet values are authored with the bundled Artifact Tool runtime.
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
const [jobPath] = process.argv.slice(2);
const job = JSON.parse(await fs.readFile(jobPath,'utf8'));
const entry = process.env.ARTIFACT_TOOL_ENTRY || path.join(process.env.USERPROFILE,
  '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs');
const {FileBlob,SpreadsheetFile}=await import(pathToFileURL(entry).href);
const book=await SpreadsheetFile.importXlsx(await FileBlob.load(job.template));
const sheet=book.worksheets.getItemAt(0);
sheet.getUsedRange().clear({applyTo:'contents'});
sheet.getRange(`A5:T${job.values.length+4}`).values=job.values;
book.recalculate();
await (await SpreadsheetFile.exportXlsx(book)).save(job.stage);
