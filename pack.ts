import * as coda from "@codahq/packs-sdk";

import XlsxTemplate from "xlsx-template";

// Jednoduchý polyfill TextDecoder pro sandboxované prostředí (Coda Pack)
import { TextDecoder as NodeTextDecoder } from "util";
if (typeof (globalThis as any).TextDecoder === "undefined") {
  (globalThis as any).TextDecoder = NodeTextDecoder;
}

export const pack = coda.newPack();

// Povolit stahování příloh (Coda ukládá na coda.io domains)
pack.addNetworkDomain("coda.io");


pack.addFormula({
  name: "GenerateXlsx",
  description: "Vytvoří XLSX soubor ze šablony (Attachment) a JSON dat.",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.File,
      name: "template",
      description: "Šablona XLSX jako příloha (Attachment ve sloupci).",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "dataJson",
      description: "JSON objekt s daty k dosazení do šablony.",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "fileName",
      description: "Název výsledného souboru (volitelné, default: output.xlsx).",
      optional: true,
    }),
  ],
  resultType: coda.ValueType.String,
  codaType: coda.ValueHintType.Attachment,

  execute: async function ([template, dataJson, fileName], context) {
    // 1. Stáhnout šablonu jako binární data
    const response = await context.fetcher.fetch({
      method: "GET",
      url: template,
      isBinaryResponse: true,
      disableAuthentication: true,
    });
    const templateBuf = Buffer.from(response.body);

    // 2. Parse JSON
    let data: any;
    try {
      data = JSON.parse(dataJson);
    } catch (e) {
      throw new coda.UserVisibleError("Neplatný JSON: " + e.message);
    }


  // 3. Vytvořit šablonu a dosadit data pomocí optilude/xlsx-template
  const xlsxTemplate = new XlsxTemplate(templateBuf);
  xlsxTemplate.substitute(1, data); // 1 = první list

  // 4. Vygenerovat výsledek jako Buffer
  const output: Buffer = Buffer.from(xlsxTemplate.generate(), 'base64');

    // 5. Vrátit jako Attachment
    let temporaryOutputUrl =
      await context.temporaryBlobStorage.storeBlob(output, "base64", {downloadFilename : fileName});
    return temporaryOutputUrl;
  },
});
