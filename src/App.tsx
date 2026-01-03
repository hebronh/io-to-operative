// src/App.tsx
import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

import { TEMPLATE_CONFIG } from "./templates/templateConfig";
import type { OrderType } from "./templates/templateConfig";
//import { fillOperativeTemplateWorkbook } from "./converters/fillTemplate";
//import { fillOperativeTemplate } from "./converters/fillTemplate";
import { fillOperativeTemplate } from "./converters/fillTemplate";
import { parseSourceIo } from "./converters/parseSourceIo"; // you already have this

export default function App() {
  const [status, setStatus] = useState("");
  const [templateBufs, setTemplateBufs] = useState<
    Partial<Record<OrderType, ArrayBuffer>>
  >({});
  const [error, setError] = useState<string | null>(null); 
  

  // cache templates
  useEffect(() => {
    (async () => {
      try {
        setError(null);
        setStatus("Loading templates...");
        const entries = await Promise.all(
          (Object.keys(TEMPLATE_CONFIG) as OrderType[]).map(async (type) => {
            const url = TEMPLATE_CONFIG[type].templateUrl;
            const res = await fetch(url);
            if (!res.ok) throw new Error(`Failed to load ${type} template (${res.status})`);
            const buf = await res.arrayBuffer(); 
            //const head = new TextDecoder().decode(buf.slice(0, 200)); 
           // console.log(`${type} template HEAD:`, head); 
            const head = new TextDecoder().decode(buf.slice(0, 200)); 
            console.log(`[${type}] template HEAD:`, head);

            return [type, buf] as const;
          })
        );

        const next: Partial<Record<OrderType, ArrayBuffer>> = {};
        for (const [type, buf] of entries) next[type] = buf;

        setTemplateBufs(next);
        setStatus("Templates loaded. Upload an IO.");
      } catch (e: any) {
        console.error(e);
        setError(e?.message ?? "Failed to load templates.");
        setStatus("");
      }
    })();
  }, []);

  const handleFile = async (file: File) => {
    try {
      setError(null);
      setStatus(`Parsing IO: ${file.name}`);

      const lines = await parseSourceIo(file);  

      console.log("first line:", lines[0]); 
      console.log("order types:", Array.from(new Set(lines.map(l => l.orderType))));

      if (!lines.length) {
        setStatus("No lines found to export.");
        return;
      }

      // group lines by order type found in the IO
      const byType = new Map<OrderType, typeof lines>();
      for (const line of lines) {
        const t = line.orderType;
        if (!byType.has(t)) byType.set(t, []);
        byType.get(t)!.push(line);
      }

      const baseName = file.name.replace(/\.(xlsx|xls)$/i, "");

      setStatus(`Generating ${byType.size} output file(s)...`);

      for (const [type, typeLines] of byType.entries()) {
        const cfg = TEMPLATE_CONFIG[type];
        const buf = templateBufs[type];
        if (!buf) throw new Error(`Template buffer missing for ${type}`);

        // IMPORTANT: set header defaults per template/type
        // you can change these rules per Spectrum vs Effectv
        const header = {
          orderId: "117224",                 // ideally pulled from IO or asked in UI
          mediaPlanName: "Default Media Plan",
          sectionName: "Default Section",
          productName: type === "Spectrum" ? "Spectrum" : "Effectv",
          costMethod: "CPM",
          unitType: "Impressions",
          canOutput: true,
          canInvoice: true,
        };

       // const wb = fillOperativeTemplateWorkbook({
       //   templateBuf: buf,
      //    sheetName: cfg.sheetName,
      //    headerRowIndex: cfg.headerRowIndex,
      //    header,
      //    lines: typeLines,
      //  });    

        const wb = fillOperativeTemplate(buf, header, typeLines);

        const head = new TextDecoder().decode(buf.slice(0, 200));  
        console.log("TEMPLATE HEAD:", head);

        // write fast as xlsx
        const out = XLSX.write(wb, { bookType: "xls", type: "array" }); 
        saveAs( 
          new Blob([out], { type: "application/vnd.ms-excel" }), 
          "Operative_Output.xls" 
        );

      }

      setStatus("Done — downloaded all template outputs.");
    } catch (e: any) {
      console.error(e);
      setError(e?.message ?? "Conversion failed.");
      setStatus("Error during conversion.");
    }
  };

  const ready =
    templateBufs.Spectrum && templateBufs.Effectv; 
  

  return (
    <div style={{ padding: 24, maxWidth: 900, margin: "0 auto" }}>
      <h2>IO → Multi-Template Operative Converter</h2>

      {error && (
        <div style={{ color: "crimson", marginBottom: 12 }}>
          <strong>Error:</strong> {error}
        </div>
      )}

      <div style={{ marginBottom: 12 }}>
        <small>
          Spectrum template:{" "}
          {templateBufs.Spectrum ? "✅ loaded" : "⏳ loading..."} <br />
          Effectv template:{" "}
          {templateBufs.Effectv ? "✅ loaded" : "⏳ loading..."}
        </small>
      </div>

      <input
        type="file"
        accept=".xlsx,.xls"
        disabled={!ready}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleFile(f);
        }}
      />

      <p style={{ marginTop: 12 }}>{status}</p>
    </div>
  );
}