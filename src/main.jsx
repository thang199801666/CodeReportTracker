import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { invoke, isTauri as isTauriRuntime } from "@tauri-apps/api/core";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import "./styles.css";

const pdfjsPromise = import("pdfjs-dist").then((pdfjs) => {
  pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
  return pdfjs;
});

const sources = {
  IAPMO: {
    type: "ER",
    link: "https://forms.iapmo.org/ues_reports/EvaluationReports.aspx",
    pdfFolder: "https://forms.iapmo.org/ues_reports/reports/",
  },
  "ICC-ES": {
    type: "ESR",
    link: "https://icc-es.org/evaluation-report-program/reports-directory/",
    pdfFolder: "https://cdn-v2.icc-es.org/wp-content/uploads/report-directory/",
  },
  "LADBS RR": {
    type: "Other",
    link: "https://www.drjcertification.org/ter-directory",
    pdfFolder: "",
  },
};

const defaultColumnWidths = [105, 130, 220, 100, 105, 105, 105, 135, 155];
const isTauri = isTauriRuntime();
const isAbsolutePath = (value) => /^(?:[A-Za-z]:[\\/]|\\\\)/.test(value || "");
const MAX_WORKERS = 8;

function Icon({ name }) {
  const shapes = {
    open: (
      <>
        <path d="M3 7.5h7l2 2h9v9.5H3z" />
        <path d="M3 7.5V5h7l2 2" />
      </>
    ),
    save: (
      <>
        <path d="M4 3h13l3 3v15H4z" />
        <path d="M8 3v6h8V3M8 21v-7h8v7" />
      </>
    ),
    saveAs: (
      <>
        <path d="M4 3h13l3 3v15H4z" />
        <path d="M8 3v6h8V3M8 21v-7h8v7M17 14v5M14.5 16.5H19.5" />
      </>
    ),
    excel: (
      <>
        <path d="M4 3h16v18H4z" />
        <path d="m8 8 4 4-4 4M13.5 16H17" />
        <path d="M7 3v18" />
      </>
    ),
    checkLink: (
      <>
        <path d="M10 13a5 5 0 0 0 7.5.5l1.5-1.5a5 5 0 0 0-7-7l-.8.8" />
        <path d="M14 11a5 5 0 0 0-7.5-.5L5 12a5 5 0 0 0 7 7l.8-.8" />
        <path d="m8 12 2 2 4-4" />
      </>
    ),
    refresh: (
      <>
        <path d="M20 11a8 8 0 0 0-14.8-4L3 9" />
        <path d="M3 4v5h5M4 13a8 8 0 0 0 14.8 4L21 15" />
        <path d="M21 20v-5h-5" />
      </>
    ),
    download: (
      <>
        <path d="M12 3v12M7 11l5 5 5-5M4 20h16" />
      </>
    ),
    stop: <path d="M6 6h12v12H6z" />,
    trash: (
      <>
        <path d="M4 7h16M10 11v6M14 11v6M6 7l1 14h10l1-14M9 7V4h6v3" />
      </>
    ),
    settings: (
      <>
        <path d="M12 15.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7Z" />
        <path d="m19.4 15 .7 1.2-1.8 1.8-1.2-.7-1.4.6-.3 1.4h-2.8l-.3-1.4-1.4-.6-1.2.7-1.8-1.8.7-1.2-.6-1.4-1.4-.3v-2.6l1.4-.3.6-1.4-.7-1.2 1.8-1.8 1.2.7 1.4-.6.3-1.4h2.8l.3 1.4 1.4.6 1.2-.7 1.8 1.8-.7 1.2.6 1.4 1.4.3v2.6l-1.4.3Z" />
      </>
    ),
    help: (
      <>
        <circle cx="12" cy="12" r="9" />
        <path d="M9.7 9a2.4 2.4 0 1 1 4 1.8c-1.1.8-1.7 1.2-1.7 2.7M12 17h.01" />
      </>
    ),
    plus: (
      <>
        <path d="M12 5v14M5 12h14" />
      </>
    ),
    rename: (
      <>
        <path d="m4 16-.8 4.8L8 20l11.5-11.5a2.1 2.1 0 0 0-3-3L5 17" />
        <path d="m14.5 7.5 3 3" />
      </>
    ),
    edit: (
      <>
        <path d="m4 16-.8 4.8L8 20 19.5 8.5a2.1 2.1 0 0 0-3-3L5 17" />
        <path d="m14.5 7.5 3 3" />
      </>
    ),
    close: (
      <>
        <path d="m7 7 10 10M17 7 7 17" />
      </>
    ),
    terminal: (
      <>
        <path d="m4 6 6 6-6 6M12 18h8" />
      </>
    ),
    github: (
      <>
        <path d="M9 19c-4 .9-4-2-5-2m10 4v-3.9c0-1 .1-1.4-.5-2.1 2.1-.2 4.3-1 4.3-4.5 0-1-.4-1.9-1-2.6.1-.2.4-1.2-.1-2.5 0 0-.8-.3-2.7 1a9.4 9.4 0 0 0-5 0c-1.9-1.3-2.7-1-2.7-1-.5 1.3-.2 2.3-.1 2.5-.6.7-1 1.6-1 2.6 0 3.5 2.2 4.3 4.3 4.5-.6.5-.6 1.1-.6 2.1V21" />
      </>
    ),
    folder: (
      <>
        <path d="M3 6h7l2 2h9v11H3z" />
        <path d="M3 6v-1h7l2 2" />
      </>
    ),
  };
  return (
    <svg
      className="icon"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.7"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {shapes[name]}
    </svg>
  );
}

function parseCodeInfo(text) {
  const code = text.match(/following[\s\S]{0,180}?(\d{4})/i)?.[1] || "n/a";
  const issueMatches = [
    ...text.matchAll(
      /(?:issue|revised)[\s\S]{0,40}?((?:\d{1,2}[/-]){1,2}\d{2,4}|[A-Za-z]{3,9}\s+\d{4})/gi,
    ),
  ];
  const valid = text.match(
    /(?:renewal|valid|expiration|expires)[\s\S]{0,80}?((?:\d{1,2}[/-]){1,2}\d{2,4}|[A-Za-z]{3,9}\s+\d{4}|\d{4})/i,
  )?.[1];
  const monthYear = (value) => {
    if (!value) return "n/a";
    const date = new Date(value);
    return Number.isNaN(date.getTime())
      ? value
      : date.toLocaleDateString("en-US", { month: "short", year: "numeric" });
  };
  return {
    latest: code,
    issue: monthYear(issueMatches.at(-1)?.[1]),
    expiration: monthYear(valid),
  };
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const externalSignal = options.signal;
  const abort = () => controller.abort(externalSignal?.reason);
  if (externalSignal) {
    if (externalSignal.aborted) abort();
    else externalSignal.addEventListener("abort", abort, { once: true });
  }
  const timer = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    window.clearTimeout(timer);
    externalSignal?.removeEventListener("abort", abort);
  }
}

async function readFirstPage(url, signal) {
  const requestUrl = /^https?:\/\//i.test(url)
    ? `/api/pdf-download?url=${encodeURIComponent(url)}`
    : url;
  const response = await fetchWithTimeout(requestUrl, { signal });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const contentType =
    response.headers.get("content-type") || "unknown content type";
  const bytes = await response.arrayBuffer();
  const signature = new TextDecoder().decode(new Uint8Array(bytes.slice(0, 5)));
  if (signature !== "%PDF-")
    throw new Error(`Response is not a PDF (${contentType}).`);
  const blob = new Blob([bytes], { type: "application/pdf" });
  const { getDocument } = await pdfjsPromise;
  const loadingTask = getDocument({ data: bytes });
  loadingTask.onPassword = (callback) => callback("");
  const pdf = await loadingTask.promise;
  try {
    const page = await pdf.getPage(1);
    const content = await page.getTextContent();
    return { text: content.items.map((item) => item.str).join(" "), blob };
  } finally {
    await pdf.destroy();
  }
}

async function readLocalPdfFirstPage(file) {
  const bytes = await file.arrayBuffer();
  const { getDocument } = await pdfjsPromise;
  const loadingTask = getDocument({ data: bytes });
  loadingTask.onPassword = (callback) => callback("");
  const pdf = await loadingTask.promise;
  try {
    const page = await pdf.getPage(1);
    const content = await page.getTextContent();
    return content.items.map((item) => item.str).join(" ");
  } finally {
    await pdf.destroy();
  }
}

async function resolveEmbeddedPdf(url, signal) {
  const response = await fetchWithTimeout(
    `/api/pdf-resolve?url=${encodeURIComponent(url)}`,
    { signal },
  );
  if (!response.ok) return null;
  return (await response.json()).url || null;
}

async function headPdf(url, signal) {
  const requestUrl = `/api/pdf-head?url=${encodeURIComponent(url)}`;
  const response = await fetchWithTimeout(requestUrl, { signal });
  if (!response.ok) return false;
  const contentType = response.headers.get("content-type")?.toLowerCase() || "";
  return (
    contentType.includes("application/pdf") ||
    new URL(url).pathname.toLowerCase().endsWith(".pdf")
  );
}

function parseCodeNumber(value) {
  const trimmed = value.trim();
  const match = trimmed.match(/^([A-Za-z]+)[\s-]*0*(\d+)$/);
  if (match) return { prefix: match[1], numeric: match[2] };
  const parts = trimmed.split(/[\s-]+/).filter(Boolean);
  if (parts.length >= 2 && /^[A-Za-z]+$/.test(parts[0]))
    return {
      prefix: parts[0],
      numeric: parts.slice(1).join("").replace(/^0+/, "") || parts.at(-1),
    };
  return { prefix: "", numeric: "" };
}

function buildPdfCandidates(number, source) {
  const { prefix, numeric } = parseCodeNumber(number);
  const setting =
    Object.values(sources).find(
      (item) => item.type.toLowerCase() === prefix.toLowerCase(),
    ) ||
    Object.values(sources).find((item) =>
      number.toLowerCase().includes(item.type.toLowerCase()),
    );
  if (!setting) return { setting: null, candidates: [] };
  const cleaned = number.replace(/[\s-]+/g, "");
  const stems = [
    ...new Set(
      [
        number.trim(),
        `${prefix}-${numeric}`,
        cleaned,
        numeric,
        numeric.length === 3 ? `0${numeric}` : "",
      ].filter(Boolean),
    ),
  ];
  const bases = [setting.link, setting.pdfFolder].filter(Boolean);
  const candidates = bases.flatMap((base) =>
    stems.map((stem) => `${base.replace(/\/$/, "")}/${stem}.pdf`),
  );
  if (
    (source?.link || "").toLowerCase().includes("icc-es.org") ||
    setting.type.toUpperCase() === "ESR"
  )
    candidates.push(
      `https://icc-es.org/wp-content/uploads/report-directory/${encodeURIComponent(number)}.pdf`,
    );
  return { setting, candidates: [...new Set(candidates)] };
}

async function checkCodeLink(row, signal) {
  if (
    row.link &&
    /^https?:\/\//i.test(row.link) &&
    (await headPdf(row.link, signal))
  )
    return { exists: true, link: row.link };
  const { candidates } = buildPdfCandidates(row.number || "", row);
  const checkedCandidates = await Promise.all(
    candidates.map(async (candidate) => {
      try {
        return (await headPdf(candidate, signal)) ? candidate : null;
      } catch (error) {
        if (error.name === "AbortError") throw error;
        return null;
      }
    }),
  );
  const foundCandidate = checkedCandidates.find(Boolean);
  if (foundCandidate) return { exists: true, link: foundCandidate };
  return { exists: false, link: row.link };
}

async function readCodeFirstPage(row, signal) {
  try {
    return { ...(await readFirstPage(row.link, signal)), sourceUrl: row.link };
  } catch (error) {
    if (!String(error?.message || "").startsWith("Response is not a PDF"))
      throw error;
    const embeddedUrl = await resolveEmbeddedPdf(row.link, signal);
    if (embeddedUrl) {
      try {
        return {
          ...(await readFirstPage(embeddedUrl, signal)),
          sourceUrl: embeddedUrl,
        };
      } catch (embeddedError) {
        if (embeddedError.name === "AbortError") throw embeddedError;
      }
    }
    const { candidates } = buildPdfCandidates(row.number || "", row);
    let lastCandidateError = error;
    const readableCandidates = await Promise.all(
      candidates.map(async (candidate) => {
        try {
          return (await headPdf(candidate, signal)) ? candidate : null;
        } catch (candidateError) {
          if (candidateError.name === "AbortError") throw candidateError;
          return null;
        }
      }),
    );
    for (const candidate of readableCandidates.filter(Boolean)) {
      try {
        return {
          ...(await readFirstPage(candidate, signal)),
          sourceUrl: candidate,
        };
      } catch (candidateError) {
        if (candidateError.name === "AbortError") throw candidateError;
        lastCandidateError = candidateError;
      }
    }
    throw lastCandidateError;
  }
}

function formatOperationError(error) {
  if (error?.name === "TypeError" || error?.message === "Failed to fetch")
    return "Failed to fetch. Check the URL, network access, or CORS settings.";
  return error?.message || "Unknown operation error.";
}

function safeDirectoryName(value) {
  return (
    String(value || "New Tab")
      .replace(/[<>:"/\\|?*]/g, "_")
      .trim() || "New Tab"
  );
}

async function runConcurrent(count, worker, signal, concurrency = 4) {
  const workerCount = Math.min(concurrency, count);
  const assignments = Array.from({ length: workerCount }, (_, workerIndex) =>
    Array.from({ length: count }, (_, index) => index).filter(
      (index) => index % workerCount === workerIndex,
    ),
  );
  await Promise.all(
    assignments.map(async (indexes, workerIndex) => {
      for (const index of indexes) {
        if (signal.aborted) return;
        await worker(index, workerIndex);
      }
    }),
  );
}

function readCrpFile(buffer) {
  const view = new DataView(buffer);
  const decoder = new TextDecoder("utf-8");
  let offset = 0;
  const readInt = () => {
    if (offset + 4 > view.byteLength)
      throw new Error("Unexpected end of CRP file.");
    const value = view.getInt32(offset, true);
    offset += 4;
    return value;
  };
  const readByte = () => {
    if (offset >= view.byteLength)
      throw new Error("Unexpected end of CRP file.");
    return view.getUint8(offset++);
  };
  const readString = () => {
    const length = readInt();
    if (length < 0) return "";
    if (offset + length > view.byteLength)
      throw new Error("Invalid CRP string length.");
    const value = decoder.decode(new Uint8Array(buffer, offset, length));
    offset += length;
    return value;
  };
  const magic = decoder.decode(
    new Uint8Array(buffer, 0, Math.min(4, view.byteLength)),
  );
  if (magic !== "CRPB")
    throw new Error("This file is not a valid Code Report file.");
  offset = 4;
  const version = readByte();
  if (version !== 1 && version !== 2)
    throw new Error(`Unsupported CRP version: ${version}`);
  const tabs = [];
  const tabCount = readInt();
  if (tabCount < 0 || tabCount > 1000)
    throw new Error("Invalid tab count in CRP file.");
  for (let tabIndex = 0; tabIndex < tabCount; tabIndex += 1) {
    const header = readString() || "New Tab";
    const itemCount = readInt();
    if (itemCount < 0 || itemCount > 100000)
      throw new Error("Invalid item count in CRP file.");
    const items = [];
    for (let itemIndex = 0; itemIndex < itemCount; itemIndex += 1) {
      const [
        number,
        link,
        webType,
        category,
        description,
        products,
        latest,
        oldLatest,
        issue,
        oldIssue,
        expiration,
        oldExpiration,
      ] = Array.from({ length: 12 }, readString);
      const progress = readInt();
      const lastCheck = readString();
      // Check results are session state and must not be restored from CRP files.
      readByte();
      readByte();
      const exists = version === 2 ? readByte() !== 0 : false;
      items.push({
        number,
        link,
        webType,
        category,
        description,
        products,
        latest,
        oldLatest,
        issue,
        oldIssue,
        expiration,
        oldExpiration,
        progress,
        lastCheck,
        checked: false,
        updated: false,
        exists,
      });
    }
    tabs.push({ id: Date.now() + tabIndex, header, items });
  }
  return tabs;
}

function writeCrpFile(tabs) {
  const encoder = new TextEncoder();
  const chunks = [new TextEncoder().encode("CRPB"), Uint8Array.of(2)];
  const appendInt = (value) => {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setInt32(0, value, true);
    chunks.push(bytes);
  };
  const appendByte = (value) => chunks.push(Uint8Array.of(value ? 1 : 0));
  const appendString = (value) => {
    const bytes = value == null ? null : encoder.encode(String(value));
    appendInt(bytes ? bytes.length : -1);
    if (bytes) chunks.push(bytes);
  };
  appendInt(tabs.length);
  tabs.forEach((tab) => {
    appendString(tab.header);
    appendInt(tab.items.length);
    tab.items.forEach((row) => {
      [
        row.number,
        row.link,
        row.webType,
        row.category,
        row.description,
        row.products,
        row.latest,
        row.oldLatest,
        row.issue,
        row.oldIssue,
        row.expiration,
        row.oldExpiration,
      ].forEach(appendString);
      appendInt(0);
      appendString(row.lastCheck);
      appendByte(false);
      appendByte(false);
      appendByte(row.exists);
    });
  });
  const output = new Uint8Array(
    chunks.reduce((size, chunk) => size + chunk.length, 0),
  );
  let position = 0;
  chunks.forEach((chunk) => {
    output.set(chunk, position);
    position += chunk.length;
  });
  return output;
}

function snapshotKey(data) {
  return Array.from(data).join(",");
}

function importExcelFile(buffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true });
  const aliases = {
    number: [
      "code report no",
      "code report number",
      "report no",
      "report number",
      "number",
    ],
    link: ["link", "url", "report link", "code report link"],
    webType: ["web type", "type"],
    category: ["product category", "category"],
    description: ["description"],
    products: ["products listed", "number of products", "products"],
    latest: ["latest code", "latest"],
    issue: ["issue/rev date", "issue date", "rev date", "revision date"],
    expiration: ["expiration date", "expiration", "expires"],
    progress: ["download process", "download progress"],
    status: ["status"],
  };
  const normalize = (value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[#/&]+/g, " ")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ");
  const normalizedAliases = Object.fromEntries(
    Object.entries(aliases).map(([field, names]) => [
      field,
      names.map(normalize),
    ]),
  );
  const findHeader = (headers, field) => {
    const names = normalizedAliases[field];
    return (
      headers.find(({ text }) => names.includes(normalize(text)))?.column ?? -1
    );
  };
  const readCell = (sheet, row, column) => {
    if (column < 0) return { text: "", link: "" };
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
    return { text: cell?.w ?? cell?.v ?? "", link: cell?.l?.Target ?? "" };
  };
  return workbook.SheetNames.map((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    let headerRow = -1;
    let indexes = null;
    let bestScore = 0;
    const lastHeaderRow = Math.min(range.e.r, range.s.r + 20);
    for (let row = range.s.r; row <= lastHeaderRow; row += 1) {
      const headers = [];
      for (let column = range.s.c; column <= range.e.c; column += 1)
        headers.push({ column, text: readCell(sheet, row, column).text });
      const candidate = Object.fromEntries(
        Object.keys(aliases).map((field) => [
          field,
          findHeader(headers, field),
        ]),
      );
      const score = Object.entries(candidate).filter(
        ([field, column]) => column >= 0 && field !== "status",
      ).length;
      if (candidate.number >= 0 && score > bestScore) {
        bestScore = score;
        headerRow = row;
        indexes = candidate;
      }
    }
    if (headerRow < 0 || !indexes) return null;
    const items = [];
    for (let row = headerRow + 1; row <= range.e.r; row += 1) {
      const values = Object.fromEntries(
        Object.entries(indexes).map(([field, column]) => [
          field,
          readCell(sheet, row, column),
        ]),
      );
      const text = (field) => String(values[field]?.text ?? "").trim();
      if (!text("number")) continue;
      if (
        !Object.values(values).some((value) => value.text !== "" || value.link)
      )
        continue;
      const status = text("status").toLowerCase();
      items.push({
        number: text("number"),
        link: values.number.link || text("link"),
        webType: text("webType"),
        category: text("category"),
        description: text("description"),
        products: text("products"),
        latest: text("latest"),
        oldLatest: "",
        issue: text("issue"),
        oldIssue: "",
        expiration: text("expiration"),
        oldExpiration: "",
        progress: Number.parseInt(text("progress"), 10) || 0,
        lastCheck: "",
        checked: false,
        updated: status.includes("update"),
        exists: false,
      });
    }
    return {
      id: Date.now() + sheetIndex,
      header: sheetName || "New Tab",
      items,
    };
  }).filter(Boolean);
}

function updateExcelWorkbook(buffer, tabs, log) {
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true });
  const aliases = {
    number: [
      "code report no",
      "code report number",
      "report no",
      "report number",
      "number",
    ],
    latest: ["latest code", "latest"],
    issue: ["issue/rev date", "issue date", "rev date", "revision date"],
    expiration: ["expiration date", "expiration", "expires"],
  };
  const normalize = (value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[#/&]+/g, " ")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ");
  const codeKey = (value) =>
    String(value ?? "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  const excelNumber = (value) => {
    const text = String(value ?? "").trim();
    if (/^\d+(?:\.\d+)?$/.test(text)) return text;
    const match = text.match(
      /^(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s-]+(\d{4})$/i,
    );
    if (!match) return null;
    const date = new Date(
      Date.UTC(
        Number(match[2]),
        new Date(`${match[1]} 1, 2000`).getUTCMonth(),
        1,
      ),
    );
    return String(
      Math.floor((date.getTime() - Date.UTC(1899, 11, 30)) / 86400000),
    );
  };
  const readText = (sheet, row, column) => {
    if (column < 0) return "";
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
    return String(cell?.w ?? cell?.v ?? "").trim();
  };
  const findColumn = (headers, names) =>
    headers.find(({ text }) => names.includes(normalize(text)))?.column ?? -1;
  const setCell = (sheet, row, column, value) => {
    if (column < 0) return;
    const address = XLSX.utils.encode_cell({ r: row, c: column });
    const cell = sheet[address] || {};
    cell.v = value || "";
    cell.w = value || "";
    cell.t = "s";
    sheet[address] = cell;
  };
  const sheetByName = new Map(
    workbook.SheetNames.map((name) => [normalize(name), name]),
  );
  let updatedCount = 0;
  for (const tab of tabs) {
    const sheetName = sheetByName.get(normalize(tab.header));
    if (!sheetName) {
      log(`Export skipped tab '${tab.header}': matching sheet not found.`);
      continue;
    }
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    let layout = null;
    for (
      let row = range.s.r;
      row <= Math.min(range.e.r, range.s.r + 20);
      row += 1
    ) {
      const headers = [];
      for (let column = range.s.c; column <= range.e.c; column += 1)
        headers.push({ column, text: readText(sheet, row, column) });
      const candidate = Object.fromEntries(
        Object.entries(aliases).map(([field, names]) => [
          field,
          findColumn(headers, names.map(normalize)),
        ]),
      );
      if (candidate.number >= 0 && candidate.latest >= 0) {
        layout = { row, ...candidate };
        break;
      }
    }
    if (!layout) {
      log(`Export skipped tab '${tab.header}': required headers not found.`);
      continue;
    }
    const rows = new Map(
      tab.items
        .map((item) => [codeKey(item.number), item])
        .filter(([key]) => key),
    );
    let sheetUpdated = 0;
    for (let row = layout.row + 1; row <= range.e.r; row += 1) {
      const item = rows.get(codeKey(readText(sheet, row, layout.number)));
      if (!item) continue;
      setCell(sheet, row, layout.latest, item.latest);
      setCell(sheet, row, layout.issue, item.issue);
      setCell(sheet, row, layout.expiration, item.expiration);
      sheetUpdated += 1;
    }
    updatedCount += sheetUpdated;
    log(`Export updated ${sheetUpdated} row(s) in '${sheetName}'.`);
  }
  return {
    data: XLSX.write(workbook, { type: "array", bookType: "xlsx" }),
    updatedCount,
  };
}

async function updateExcelTemplate(buffer, tabs, log) {
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true });
  const normalize = (value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[#/&]+/g, " ")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ");
  const codeKey = (value) =>
    String(value ?? "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  const excelNumber = (value) => {
    const text = String(value ?? "").trim();
    if (/^\d+(?:\.\d+)?$/.test(text)) return text;
    const match = text.match(
      /^(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s-]+(\d{4})$/i,
    );
    if (!match) return null;
    const date = new Date(
      Date.UTC(
        Number(match[2]),
        new Date(`${match[1]} 1, 2000`).getUTCMonth(),
        1,
      ),
    );
    return String(
      Math.floor((date.getTime() - Date.UTC(1899, 11, 30)) / 86400000),
    );
  };
  const aliases = {
    number: [
      "code report no",
      "code report number",
      "report no",
      "report number",
      "number",
    ],
    latest: ["latest code", "latest"],
    issue: ["issue/rev date", "issue date", "rev date", "revision date"],
    expiration: ["expiration date", "expiration", "expires"],
  };
  const readText = (sheet, row, column) => {
    if (column < 0) return "";
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
    return String(cell?.w ?? cell?.v ?? "").trim();
  };
  const updatesBySheet = new Map();
  const sheetByName = new Map(
    workbook.SheetNames.map((name) => [normalize(name), name]),
  );
  let updatedCount = 0;
  for (const tab of tabs) {
    const sheetName = sheetByName.get(normalize(tab.header));
    if (!sheetName) {
      log(`Export skipped tab '${tab.header}': matching sheet not found.`);
      continue;
    }
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    let layout = null;
    for (
      let row = range.s.r;
      row <= Math.min(range.e.r, range.s.r + 20);
      row += 1
    ) {
      const headers = [];
      for (let column = range.s.c; column <= range.e.c; column += 1)
        headers.push({ column, text: readText(sheet, row, column) });
      const candidate = Object.fromEntries(
        Object.entries(aliases).map(([field, names]) => [
          field,
          headers.find(({ text }) =>
            names.map(normalize).includes(normalize(text)),
          )?.column ?? -1,
        ]),
      );
      if (candidate.number >= 0 && candidate.latest >= 0) {
        layout = { row, ...candidate };
        break;
      }
    }
    if (!layout) {
      log(`Export skipped tab '${tab.header}': required headers not found.`);
      continue;
    }
    const rows = new Map(
      tab.items
        .map((item) => [codeKey(item.number), item])
        .filter(([key]) => key),
    );
    const updates = [];
    let changedRows = 0;
    for (let row = layout.row + 1; row <= range.e.r; row += 1) {
      const item = rows.get(codeKey(readText(sheet, row, layout.number)));
      if (!item) continue;
      const rowUpdates = [
        [layout.latest, item.latest],
        [layout.issue, item.issue],
        [layout.expiration, item.expiration],
      ]
        .filter(
          ([column, value]) =>
            column >= 0 && readText(sheet, row, column) !== String(value || ""),
        )
        .map(([column, value]) => ({
          address: XLSX.utils.encode_cell({ r: row, c: column }),
          value: value || "",
        }));
      updates.push(...rowUpdates);
      if (rowUpdates.length) {
        updatedCount += 1;
        changedRows += 1;
      }
    }
    updatesBySheet.set(sheetName, updates);
    log(`Export prepared ${changedRows} row(s) in '${sheetName}'.`);
  }

  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = new DOMParser().parseFromString(
    await zip.file("xl/workbook.xml").async("string"),
    "application/xml",
  );
  const relationshipsXml = new DOMParser().parseFromString(
    await zip.file("xl/_rels/workbook.xml.rels").async("string"),
    "application/xml",
  );
  const relationships = new Map(
    Array.from(relationshipsXml.getElementsByTagName("Relationship")).map(
      (relationship) => [
        relationship.getAttribute("Id"),
        relationship.getAttribute("Target"),
      ],
    ),
  );
  const styleMap = new Map();
  const sheets = Array.from(workbookXml.getElementsByTagNameNS("*", "sheet"));
  let modifiedCells = 0;
  for (const sheetNode of sheets) {
    const sheetName = sheetNode.getAttribute("name");
    const updates = updatesBySheet.get(sheetName);
    if (!updates?.length) continue;
    const relationshipId =
      sheetNode.getAttribute("r:id") ||
      sheetNode.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      );
    const target = relationships.get(relationshipId);
    if (!target) continue;
    const xmlPath = target.startsWith("/")
      ? target.slice(1)
      : `xl/${target.replace(/^\.\//, "")}`;
    let sheetXml = await zip.file(xmlPath).async("string");
    for (const update of updates) {
      const address = update.address.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const cellPattern = new RegExp(
        `<c(\\s[^>]*\\br="${address}"[^>]*)>([\\s\\S]*?)</c>|<c(\\s[^>]*\\br="${address}"[^>]*)\\s*/>`,
      );
      const match = sheetXml.match(cellPattern);
      if (!match) continue;
      const attributes = match[1] || match[3];
      const originalStyle = Number.parseInt(
        attributes.match(/\bs="(\d+)"/)?.[1] || "0",
        10,
      );
      const orangeStyle = styleMap.get("orangeStyles")?.get(originalStyle);
      let nextAttributes = attributes.replace(/\s+t="[^"]*"/g, "");
      if (orangeStyle != null) {
        nextAttributes = /\bs="\d+"/.test(nextAttributes)
          ? nextAttributes.replace(/\bs="\d+"/, `s="${orangeStyle}"`)
          : `${nextAttributes} s="${orangeStyle}"`;
      }
      const escapedValue = String(update.value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      const originalType = attributes.match(/\bt="([^"]+)"/)?.[1];
      const numericValue =
        !originalType || originalType === "n"
          ? excelNumber(update.value)
          : null;
      if (numericValue !== null) {
        nextAttributes = nextAttributes.replace(/\s+t="[^"]*"/g, "");
        sheetXml = sheetXml.replace(
          cellPattern,
          `<c${nextAttributes}><v>${numericValue}</v></c>`,
        );
        modifiedCells += 1;
        continue;
      }
      const preserveSpace = /^\s|\s$/.test(String(update.value || ""))
        ? ' xml:space="preserve"'
        : "";
      const replacement = `<c${nextAttributes} t="inlineStr"><is><t${preserveSpace}>${escapedValue}</t></is></c>`;
      sheetXml = sheetXml.replace(cellPattern, replacement);
      modifiedCells += 1;
    }
    zip.file(xmlPath, sheetXml);
  }
  return {
    data: await zip.generateAsync({ type: "uint8array" }),
    updatedCount,
    modifiedCells,
  };
}

function buildTabsWorkbook(tabs) {
  const workbook = XLSX.utils.book_new();
  const headers = [
    "Code Report No",
    "Product Category",
    "Description",
    "Products Listed",
    "Latest Code",
    "Issue/Rev Date",
    "Expiration Date",
  ];
  const usedNames = new Set();
  const sheetName = (value, index) => {
    const base =
      String(value || `Tab ${index + 1}`)
        .replace(/[\\/*?:[\]]/g, " ")
        .trim()
        .slice(0, 31) || `Tab ${index + 1}`;
    let name = base;
    let suffix = 2;
    while (usedNames.has(name)) {
      const tail = ` (${suffix})`;
      name = `${base.slice(0, 31 - tail.length)}${tail}`;
      suffix += 1;
    }
    usedNames.add(name);
    return name;
  };
  const headerStyle = {
    fill: { fgColor: { rgb: "FF1F4E78" } },
    font: { bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
  };
  const orangeStyle = { fill: { fgColor: { rgb: "FFFFC000" } } };
  const linkStyle = {
    font: { color: { rgb: "FF2E8DEF" }, underline: "single" },
  };
  const missingLinkStyle = {
    font: { color: { rgb: "FFC43D3D" }, underline: "single" },
  };
  tabs.forEach((tab, tabIndex) => {
    const rows = [
      headers,
      ...tab.items.map((row) => [
        row.number || "",
        row.category || "",
        row.description || "",
        row.products || "",
        row.latest || "",
        row.issue || "",
        row.expiration || "",
      ]),
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [
      { wch: 18 },
      { wch: 22 },
      { wch: 38 },
      { wch: 18 },
      { wch: 14 },
      { wch: 18 },
      { wch: 18 },
    ];
    sheet["!rows"] = [{ hpt: 28 }];
    headers.forEach((_, column) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: column })];
      if (cell) cell.s = headerStyle;
    });
    tab.items.forEach((row, rowIndex) => {
      const excelRow = rowIndex + 1;
      const linkCell = sheet[XLSX.utils.encode_cell({ r: excelRow, c: 0 })];
      if (linkCell) {
        linkCell.s = row.checked && !row.exists ? missingLinkStyle : linkStyle;
        if (row.link) linkCell.l = { Target: row.link };
      }
      [
        [4, row.oldLatest, row.latest],
        [5, row.oldIssue, row.issue],
        [6, row.oldExpiration, row.expiration],
      ].forEach(([column, oldValue, newValue]) => {
        const cell = sheet[XLSX.utils.encode_cell({ r: excelRow, c: column })];
        if (cell && oldValue && oldValue !== newValue) cell.s = orangeStyle;
      });
    });
    XLSX.utils.book_append_sheet(
      workbook,
      sheet,
      sheetName(tab.header, tabIndex),
    );
  });
  return XLSX.write(workbook, {
    type: "array",
    bookType: "xlsx",
    cellStyles: true,
  });
}

function Button({ icon, children, onClick, disabled, split }) {
  return (
    <button
      className={`ribbon-button icon-${icon}`}
      onClick={onClick}
      disabled={disabled}
    >
      <Icon name={icon} />
      <span>{children}</span>
      {split && <b className="split-arrow">⌄</b>}
    </button>
  );
}

function SplitButton({ icon, children, onClick, disabled, items }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  useEffect(() => {
    if (!open) return undefined;
    const closeWhenOutside = (event) => {
      if (!rootRef.current?.contains(event.target)) setOpen(false);
    };
    document.addEventListener("pointerdown", closeWhenOutside);
    return () => document.removeEventListener("pointerdown", closeWhenOutside);
  }, [open]);
  const choose = (action) => {
    setOpen(false);
    action?.();
  };
  return (
    <div className={`split-button-wrap icon-${icon}`} ref={rootRef}>
      <button
        className="ribbon-button split-main"
        onClick={onClick}
        disabled={disabled}
      >
        <Icon name={icon} />
        <span>{children}</span>
      </button>
      <button
        className="split-toggle"
        onClick={() => setOpen((value) => !value)}
        disabled={disabled}
        aria-label={`${children} menu`}
      >
        ⌄
      </button>
      {open && (
        <div className="split-menu">
          {items.map((item) => (
            <button key={item.label} onClick={() => choose(item.action)}>
              <Icon name={item.icon || icon} />
              <span>{item.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function App() {
  const [tabs, setTabs] = useState([{ id: 1, header: "New Tab", items: [] }]);
  const [selectedId, setSelectedId] = useState(1);
  const [busy, setBusy] = useState(false);
  const [consoleText, setConsoleText] = useState(
    "**************Initialize**************\n",
  );
  const [showSettings, setShowSettings] = useState(false);
  const [activeRibbonTab, setActiveRibbonTab] = useState("Home");
  const [editingTabId, setEditingTabId] = useState(null);
  const [draggedTabId, setDraggedTabId] = useState(null);
  const [dragOverTabId, setDragOverTabId] = useState(null);
  const [dragPreview, setDragPreview] = useState(null);
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [lastSelectedRow, setLastSelectedRow] = useState(null);
  const [columnWidths, setColumnWidths] = useState(defaultColumnWidths);
  const [columnResize, setColumnResize] = useState(null);
  const [draftRow, setDraftRow] = useState({
    number: "",
    link: "",
    category: "",
    description: "",
    products: "",
    latest: "",
    issue: "",
    expiration: "",
    checked: false,
    updated: false,
  });
  const [contextMenu, setContextMenu] = useState(null);
  const [editingRowIndex, setEditingRowIndex] = useState(null);
  const [isDirty, setIsDirty] = useState(false);
  const [consoleHeight, setConsoleHeight] = useState(142);
  const [resizingConsole, setResizingConsole] = useState(false);
  const downloadDirectoryRef = useRef(null);
  const consoleRef = useRef(null);
  const [settings, setSettings] = useState(() => {
    try {
      return {
        workerCount: 4,
        downloadDirectory: "Downloads",
        ...JSON.parse(localStorage.getItem("code-report-settings") || "{}"),
      };
    } catch {
      return { workerCount: 4, downloadDirectory: "Downloads" };
    }
  });
  const crpInputRef = useRef(null);
  const exportInputRef = useRef(null);
  const exportModeRef = useRef("new");
  const operationControllerRef = useRef(null);
  const savedSnapshotRef = useRef(snapshotKey(writeCrpFile(tabs)));
  const loadedSnapshotRef = useRef(null);
  const snapshotInitializedRef = useRef(false);
  const current = tabs.find((tab) => tab.id === selectedId) || tabs[0];
  useEffect(() => {
    if (isTauri && isAbsolutePath(settings.downloadDirectory)) {
      downloadDirectoryRef.current = settings.downloadDirectory;
    }
  }, [settings.downloadDirectory]);
  useEffect(() => {
    const snapshot = snapshotKey(writeCrpFile(tabs));
    if (!snapshotInitializedRef.current) {
      savedSnapshotRef.current = snapshot;
      snapshotInitializedRef.current = true;
      setIsDirty(false);
      return;
    }
    if (loadedSnapshotRef.current !== null) {
      savedSnapshotRef.current = loadedSnapshotRef.current;
      loadedSnapshotRef.current = null;
      setIsDirty(false);
      return;
    }
    setIsDirty(snapshot !== savedSnapshotRef.current);
  }, [tabs]);
  useEffect(() => {
    const warnBeforeClose = (event) => {
      if (!isDirty) return;
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", warnBeforeClose);
    return () => window.removeEventListener("beforeunload", warnBeforeClose);
  }, [isDirty]);
  useEffect(() => {
    if (consoleRef.current)
      consoleRef.current.scrollTop = consoleRef.current.scrollHeight;
  }, [consoleText]);
  useEffect(() => {
    const showContextMenu = (event) => {
      if (event.defaultPrevented) return;
      event.preventDefault();
      if (!event.target.closest(".table-wrap")) return;
      setContextMenu({ x: event.clientX, y: event.clientY });
    };
    const closeContextMenu = () => setContextMenu(null);
    document.addEventListener("contextmenu", showContextMenu);
    document.addEventListener("click", closeContextMenu);
    document.addEventListener("scroll", closeContextMenu, true);
    return () => {
      document.removeEventListener("contextmenu", showContextMenu);
      document.removeEventListener("click", closeContextMenu);
      document.removeEventListener("scroll", closeContextMenu, true);
    };
  }, []);
  const log = (text) =>
    setConsoleText((value) => {
      const lines = `${value}-[${new Date().toLocaleTimeString()}] : ${text}`
        .split("\n")
        .filter(Boolean);
      return `${lines.slice(-500).join("\n")}\n`;
    });
  useEffect(() => {
    const reportError = (message) => log(`Unexpected error: ${message}`);
    const handleError = (event) => {
      console.error(
        "Unhandled application error",
        event.error || event.message,
      );
      reportError(event.error?.message || event.message || "Unknown error");
    };
    const handleRejection = (event) => {
      console.error("Unhandled promise rejection", event.reason);
      reportError(event.reason?.message || String(event.reason));
    };
    window.addEventListener("error", handleError);
    window.addEventListener("unhandledrejection", handleRejection);
    return () => {
      window.removeEventListener("error", handleError);
      window.removeEventListener("unhandledrejection", handleRejection);
    };
  }, []);
  const updateRow = (index, patch) =>
    setTabs((all) =>
      all.map((tab) =>
        tab.id === selectedId
          ? {
              ...tab,
              items: tab.items.map((row, i) =>
                i === index ? { ...row, ...patch } : row,
              ),
            }
          : tab,
      ),
    );
  const selectTab = (id) => {
    setSelectedId(id);
    setSelectedRows(new Set());
    setLastSelectedRow(null);
  };
  const addTab = () => {
    const id = Date.now();
    setTabs((all) => [...all, { id, header: "New Tab", items: [] }]);
    selectTab(id);
    log("Created new tab.");
  };
  const closeTab = (id) => {
    const tab = tabs.find((item) => item.id === id);
    if (
      !window.confirm(
        `Delete tab '${tab?.header || "New Tab"}'? All rows in this tab will be removed.`,
      )
    )
      return;
    const next = tabs.filter((tab) => tab.id !== id);
    if (!next.length) {
      const replacement = { id: Date.now(), header: "New Tab", items: [] };
      setTabs([replacement]);
      setSelectedId(replacement.id);
      return;
    }
    setTabs(next);
    if (id === selectedId) setSelectedId(next[0].id);
  };
  const renameTab = (id, header) => {
    const nextHeader = header.trim() || "New Tab";
    setTabs((all) =>
      all.map((tab) => (tab.id === id ? { ...tab, header: nextHeader } : tab)),
    );
    setEditingTabId(null);
  };
  const reorderTabs = (targetId) => {
    if (draggedTabId == null || draggedTabId === targetId) return;
    setTabs((all) => {
      const fromIndex = all.findIndex((tab) => tab.id === draggedTabId);
      const targetIndex = all.findIndex((tab) => tab.id === targetId);
      if (fromIndex < 0 || targetIndex < 0) return all;
      const next = [...all];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(targetIndex, 0, moved);
      return next;
    });
  };
  const moveTabOver = (targetId) => {
    if (dragOverTabId === targetId) return;
    reorderTabs(targetId);
    setDragOverTabId(targetId);
  };
  const startTabDrag = (event, tab) => {
    setDraggedTabId(tab.id);
    setDragPreview({ header: tab.header, x: event.clientX, y: event.clientY });
    event.dataTransfer.effectAllowed = "move";
  };
  const moveTabDrag = (event) =>
    setDragPreview((preview) =>
      preview ? { ...preview, x: event.clientX, y: event.clientY } : preview,
    );
  const endTabDrag = () => {
    setDraggedTabId(null);
    setDragOverTabId(null);
    setDragPreview(null);
  };
  const updateDraftRow = (field, value) =>
    setDraftRow((row) => ({ ...row, [field]: value }));
  const commitDraftRow = () => {
    setTabs((all) =>
      all.map((tab) =>
        tab.id === selectedId
          ? {
              ...tab,
              items: [
                ...tab.items,
                {
                  ...draftRow,
                  oldLatest: "",
                  oldIssue: "",
                  oldExpiration: "",
                  progress: 0,
                  checked: false,
                  updated: false,
                  exists: true,
                },
              ],
            }
          : tab,
      ),
    );
    setDraftRow({
      number: "",
      link: "",
      category: "",
      description: "",
      products: "",
      latest: "",
      issue: "",
      expiration: "",
      checked: false,
      updated: false,
    });
  };
  const selectRow = (index, event) => {
    const next = new Set(selectedRows);
    if (event.shiftKey && lastSelectedRow !== null) {
      const start = Math.min(lastSelectedRow, index);
      const end = Math.max(lastSelectedRow, index);
      for (let rowIndex = start; rowIndex <= end; rowIndex += 1)
        next.add(rowIndex);
    } else if (event.ctrlKey || event.metaKey) {
      if (next.has(index)) next.delete(index);
      else next.add(index);
      setLastSelectedRow(index);
    } else {
      next.clear();
      next.add(index);
      setLastSelectedRow(index);
    }
    setSelectedRows(next);
  };
  const showRowContextMenu = (event, index) => {
    event.preventDefault();
    event.stopPropagation();
    if (!selectedRows.has(index)) {
      setSelectedRows(new Set([index]));
      setLastSelectedRow(index);
    }
    setContextMenu({ x: event.clientX, y: event.clientY, row: true });
  };
  const showTabContextMenu = (event, tabId) => {
    event.preventDefault();
    event.stopPropagation();
    setContextMenu({ x: event.clientX, y: event.clientY, tabId });
  };
  const deleteSelectedRows = () => {
    const count = selectedRows.size;
    if (!count) return;
    setTabs((all) =>
      all.map((tab) =>
        tab.id === selectedId
          ? {
              ...tab,
              items: tab.items.filter((_, index) => !selectedRows.has(index)),
            }
          : tab,
      ),
    );
    setSelectedRows(new Set());
    setLastSelectedRow(null);
    setContextMenu(null);
    log(`Deleted ${count} selected row(s).`);
  };
  const editSelectedRow = () => {
    if (selectedRows.size !== 1) return;
    setEditingRowIndex([...selectedRows][0]);
    setContextMenu(null);
  };
  const runCheck = async (mode = "update") => {
    if (!current.items.length) {
      log(`${mode} aborted: no rows to process.`);
      return;
    }
    let localDirectory = null;
    if (mode === "local") {
      const rootDirectory = await ensureDownloadDirectory("Update Local");
      if (!rootDirectory) return;
      try {
        if (isTauri) {
          localDirectory = rootDirectory;
        } else {
          const permission = await rootDirectory.queryPermission({
            mode: "read",
          });
          if (permission !== "granted")
            await rootDirectory.requestPermission({ mode: "read" });
          localDirectory = await rootDirectory.getDirectoryHandle(
            safeDirectoryName(current.header),
            { create: false },
          );
        }
      } catch (error) {
        log(
          error.name === "NotFoundError"
            ? `Update Local skipped: folder '${current.header}' was not found. Download PDFs for this tab first.`
            : `Update Local failed: ${formatOperationError(error)}`,
        );
        return;
      }
    }
    setBusy(true);
    const controller = new AbortController();
    operationControllerRef.current = controller;
    log(
      mode === "check"
        ? "Check Link started: checking PDF existence..."
        : `${mode === "local" ? "Update Local" : "Search"} started: reading first page of PDFs with 4 workers...`,
    );
    await runConcurrent(
      current.items.length,
      async (i) => {
        const row = current.items[i];
        if (controller.signal.aborted) return;
        if (mode !== "check" && mode !== "local" && !row.link) {
          log(`${row.number || "(unnamed)"} skipped: no link.`);
          return;
        }
        try {
          if (mode === "check") {
            const result = await checkCodeLink(row, controller.signal);
            updateRow(i, {
              checked: true,
              exists: result.exists,
              link: result.link,
            });
            log(
              `Checked ${row.number || "(unnamed)"}: PDF ${result.exists ? "found" : "missing"}`,
            );
          } else {
            let text;
            if (mode === "local") {
              const fileName = `${row.number || "report"}.pdf`;
              if (isTauri) {
                const bytes = await invoke("read_pdf", {
                  root: localDirectory,
                  tab: current.header,
                  fileName,
                });
                text = await readLocalPdfFirstPage(
                  new Blob([new Uint8Array(bytes)]),
                );
              } else {
                const fileHandle = await localDirectory.getFileHandle(
                  fileName,
                  { create: false },
                );
                text = await readLocalPdfFirstPage(await fileHandle.getFile());
              }
            } else {
              ({ text } = await readCodeFirstPage(row, controller.signal));
            }
            const info = parseCodeInfo(text);
            const update = {
              oldLatest: row.latest,
              oldIssue: row.issue,
              oldExpiration: row.expiration,
              ...info,
              updated:
                info.latest !== row.latest ||
                info.issue !== row.issue ||
                info.expiration !== row.expiration,
            };
            updateRow(i, update);
            log(`${row.number || "PDF"}: completed.`);
          }
        } catch (error) {
          if (controller.signal.aborted || error.name === "AbortError") return;
          if (mode === "local" && error.name === "NotFoundError") {
            log(`${row.number || "PDF"} skipped: local PDF was not found.`);
            return;
          }
          log(`${row.number || "PDF"} failed: ${formatOperationError(error)}`);
        }
      },
      controller.signal,
      Math.max(1, Math.min(MAX_WORKERS, Number(settings.workerCount) || 4)),
    );
    log(
      controller.signal.aborted
        ? "Operation cancelled."
        : `${mode === "check" ? "Check Link" : mode === "local" ? "Update Local" : "Search"} finished.`,
    );
    operationControllerRef.current = null;
    setBusy(false);
  };
  const ensureDownloadDirectory = async (operation) => {
    if (isTauri) {
      if (isAbsolutePath(downloadDirectoryRef.current))
        return downloadDirectoryRef.current;
      try {
        const directory = await invoke("pick_directory");
        if (directory) {
          downloadDirectoryRef.current = directory;
          return directory;
        }
      } catch (error) {
        log(`${operation} folder selection failed: ${error}`);
      }
      return null;
    }
    if (downloadDirectoryRef.current?.getFileHandle)
      return downloadDirectoryRef.current;
    if (!window.showDirectoryPicker) {
      log(
        `${operation} unavailable: folder access is not supported by this browser.`,
      );
      return null;
    }
    try {
      downloadDirectoryRef.current = await window.showDirectoryPicker({
        mode: "readwrite",
      });
      return downloadDirectoryRef.current;
    } catch (error) {
      if (error.name !== "AbortError")
        log(`${operation} folder selection failed: ${error.message}`);
      return null;
    }
  };
  const downloadPdfs = async () => {
    if (!(await ensureDownloadDirectory("Download PDFs"))) return;
    setBusy(true);
    const controller = new AbortController();
    operationControllerRef.current = controller;
    log("Download started: saving PDFs from selected tab.");
    await runConcurrent(
      current.items.length,
      async (i) => {
        const row = current.items[i];
        if (controller.signal.aborted || !row.link) return;
        try {
          const { blob } = await readCodeFirstPage(row, controller.signal);
          const fileName = `${row.number || "report"}.pdf`;
          const rootDirectory = downloadDirectoryRef.current;
          if (isTauri) {
            await invoke("write_pdf", {
              root: rootDirectory,
              tab: current.header,
              fileName,
              data: Array.from(new Uint8Array(await blob.arrayBuffer())),
            });
            updateRow(i, { progress: 100 });
            log(`${row.number || "PDF"} downloaded.`);
            return;
          }
          const permission = await rootDirectory.queryPermission({
            mode: "readwrite",
          });
          if (permission !== "granted")
            await rootDirectory.requestPermission({ mode: "readwrite" });
          const directory = await rootDirectory.getDirectoryHandle(
            safeDirectoryName(current.header),
            { create: true },
          );
          const fileHandle = await directory.getFileHandle(fileName, {
            create: true,
          });
          const writable = await fileHandle.createWritable();
          await writable.write(blob);
          await writable.close();
          updateRow(i, { progress: 100 });
          log(`${row.number || "PDF"} downloaded.`);
        } catch (error) {
          if (controller.signal.aborted || error.name === "AbortError") return;
          log(
            `${row.number || "PDF"} download failed: ${formatOperationError(error)}`,
          );
        }
      },
      controller.signal,
      Math.max(1, Math.min(MAX_WORKERS, Number(settings.workerCount) || 4)),
    );
    operationControllerRef.current = null;
    setBusy(false);
    log(
      controller.signal.aborted ? "Download cancelled." : "Download finished.",
    );
  };
  const openTabFolder = async (tabId) => {
    const targetTab = tabs.find((tab) => tab.id === tabId) || current;
    if (!targetTab) return;
    const rootDirectory = await ensureDownloadDirectory("Open Folder");
    if (!rootDirectory) return;
    try {
      if (isTauri) {
        await invoke("open_folder", {
          root: rootDirectory,
          tab: targetTab.header,
        });
        return;
      }
      const directory = await rootDirectory.getDirectoryHandle(
        safeDirectoryName(targetTab.header),
        { create: false },
      );
      if (window.showDirectoryPicker) {
        await window.showDirectoryPicker({
          mode: "readwrite",
          startIn: directory,
        });
      }
    } catch (error) {
      log(`Open Folder failed: ${formatOperationError(error)}`);
    }
  };
  const deletePdfs = async (tabId = selectedId) => {
    const targetTab = tabs.find((tab) => tab.id === tabId) || current;
    if (!targetTab) return;
    if (!(await ensureDownloadDirectory("Delete PDFs"))) return;
    if (
      !window.confirm(
        `Delete all PDF files from the '${targetTab.header}' folder? This action cannot be undone.`,
      )
    )
      return;
    setBusy(true);
    try {
      const rootDirectory = downloadDirectoryRef.current;
      if (isTauri) {
        const deleted = await invoke("delete_pdfs", {
          root: rootDirectory,
          tab: targetTab.header,
        });
        log(`Deleted ${deleted} PDF file(s) from '${targetTab.header}'.`);
        return;
      }
      const permission = await rootDirectory.queryPermission({
        mode: "readwrite",
      });
      if (permission !== "granted")
        await rootDirectory.requestPermission({ mode: "readwrite" });
      const directory = await rootDirectory.getDirectoryHandle(
        safeDirectoryName(targetTab.header),
        { create: false },
      );
      let deleted = 0;
      for await (const [name] of directory.entries()) {
        if (name.toLowerCase().endsWith(".pdf")) {
          await directory.removeEntry(name);
          deleted += 1;
        }
      }
      log(`Deleted ${deleted} PDF file(s) from '${targetTab.header}'.`);
    } catch (error) {
      log(`Delete PDFs failed: ${formatOperationError(error)}`);
    } finally {
      setBusy(false);
    }
  };
  const openDataFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const loadedTabs = file.name.toLowerCase().endsWith(".crp")
        ? readCrpFile(buffer)
        : importExcelFile(buffer);
      const nextTabs = loadedTabs.length
        ? loadedTabs
        : [{ id: Date.now(), header: "New Tab", items: [] }];
      setTabs(nextTabs);
      loadedSnapshotRef.current = snapshotKey(writeCrpFile(nextTabs));
      setSelectedId(nextTabs[0]?.id || null);
      log(`Imported file: ${file.name}`);
    } catch (error) {
      log(`Failed to load CRP file: ${error.message}`);
    }
    event.target.value = "";
  };
  const exportExcelBuffer = async (buffer, fileName, overwritePath = null) => {
    try {
      const result = await updateExcelTemplate(buffer, tabs, log);
      if (overwritePath) {
        await invoke("overwrite_file", {
          path: overwritePath,
          data: Array.from(result.data),
        });
        log(
          `Overwrote ${result.updatedCount} row(s), ${result.modifiedCells} cell(s) in ${overwritePath}.`,
        );
      } else if (isTauri) {
        const path = await invoke("save_excel", {
          data: Array.from(result.data),
        });
        if (path)
          log(
            `Exported ${result.updatedCount} row(s), ${result.modifiedCells} cell(s) to ${path}.`,
          );
      } else {
        const outputName = `${fileName.replace(/\.[^.]+$/, "")}-updated.xlsx`;
        const blob = new Blob([result.data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const objectUrl = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = objectUrl;
        anchor.download = outputName;
        anchor.click();
        window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
        log(
          `Exported ${result.updatedCount} row(s), ${result.modifiedCells} cell(s) to ${outputName}.`,
        );
      }
    } catch (error) {
      log(`Excel export failed: ${error.message || error}`);
    }
  };
  const exportExcelFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (exportModeRef.current === "overwrite") {
      window.alert(
        "Overwriting an existing file requires the Tauri desktop app.",
      );
      event.target.value = "";
      return;
    }
    await exportExcelBuffer(await file.arrayBuffer(), file.name);
    event.target.value = "";
  };
  const startExcelExport = async (mode) => {
    exportModeRef.current = mode;
    if (!isTauri) {
      if (mode === "overwrite") {
        if (!window.showOpenFilePicker) {
          window.alert(
            "Overwriting requires a browser with File System Access API support, such as Chrome or Edge.",
          );
          return;
        }
        try {
          const [handle] = await window.showOpenFilePicker({
            types: [
              {
                description: "Excel Workbook",
                accept: {
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    [".xlsx"],
                  "application/vnd.ms-excel": [".xls"],
                },
              },
            ],
            multiple: false,
          });
          const file = await handle.getFile();
          const result = await updateExcelTemplate(
            await file.arrayBuffer(),
            tabs,
            log,
          );
          const permission = await handle.queryPermission({
            mode: "readwrite",
          });
          if (permission !== "granted")
            await handle.requestPermission({ mode: "readwrite" });
          const writable = await handle.createWritable();
          await writable.write(result.data);
          await writable.close();
          log(
            `Overwrote ${result.updatedCount} row(s), ${result.modifiedCells} cell(s) in ${file.name}.`,
          );
        } catch (error) {
          if (error.name !== "AbortError")
            log(`Excel overwrite failed: ${error.message || error}`);
        }
        return;
      }
      exportInputRef.current?.click();
      return;
    }
    try {
      const selected = await invoke("pick_excel_file");
      if (!selected) return;
      await exportExcelBuffer(
        new Uint8Array(selected.data).buffer,
        selected.path,
        mode === "overwrite" ? selected.path : null,
      );
    } catch (error) {
      log(`Excel file selection failed: ${error.message || error}`);
    }
  };
  const exportTabsToFile = async () => {
    try {
      const data = buildTabsWorkbook(tabs);
      if (isTauri) {
        const path = await invoke("save_excel", { data: Array.from(data) });
        if (path) log(`Exported ${tabs.length} tab(s) to ${path}.`);
        return;
      }
      const blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const objectUrl = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = "code-report-tracker.xlsx";
      anchor.click();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
      log(`Exported ${tabs.length} tab(s) to code-report-tracker.xlsx.`);
    } catch (error) {
      log(`Excel export failed: ${error.message || error}`);
    }
  };
  const saveReport = async () => {
    const data = writeCrpFile(tabs);
    if (isTauri) {
      try {
        const path = await invoke("save_crp", { data: Array.from(data) });
        if (!path) return;
        savedSnapshotRef.current = snapshotKey(data);
        setIsDirty(false);
        log(`Saved current Code Report to ${path}.`);
      } catch (error) {
        log(`Save failed: ${error}`);
      }
      return;
    }
    const blob = new Blob([data], { type: "application/octet-stream" });
    const a = document.createElement("a");
    const objectUrl = URL.createObjectURL(blob);
    a.href = objectUrl;
    a.download = "report.crp";
    a.target = "_self";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
    savedSnapshotRef.current = snapshotKey(data);
    setIsDirty(false);
    log("Saved current Code Report.");
  };
  const cancelWork = () => {
    if (operationControllerRef.current) {
      operationControllerRef.current.abort();
      return;
    }
    setBusy(false);
    log("Cancel");
  };
  const saveSettings = (nextSettings) => {
    const normalized = {
      workerCount: Math.max(
        1,
        Math.min(MAX_WORKERS, Number(nextSettings.workerCount) || 4),
      ),
      downloadDirectory: nextSettings.downloadDirectory || "Downloads",
    };
    setSettings(normalized);
    if (isTauri)
      downloadDirectoryRef.current = isAbsolutePath(
        normalized.downloadDirectory,
      )
        ? normalized.downloadDirectory
        : null;
    localStorage.setItem("code-report-settings", JSON.stringify(normalized));
  };
  const setDownloadDirectory = (handle) => {
    downloadDirectoryRef.current = handle;
  };
  const closeTabFromButton = (event, id) => {
    event.stopPropagation();
    closeTab(id);
  };
  const startConsoleResize = (event) => {
    event.preventDefault();
    setResizingConsole(true);
    event.currentTarget.setPointerCapture(event.pointerId);
  };
  const resizeConsole = (event) => {
    if (!resizingConsole) return;
    const nextHeight = window.innerHeight - event.clientY - 25;
    setConsoleHeight(
      Math.min(Math.max(nextHeight, 86), Math.floor(window.innerHeight * 0.65)),
    );
  };
  const stopConsoleResize = () => setResizingConsole(false);
  const startColumnResize = (event, index) => {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    const neighborIndex =
      index === columnWidths.length - 1 ? index - 1 : index + 1;
    setColumnResize({
      index,
      neighborIndex,
      startX: event.clientX,
      startWidth: columnWidths[index],
      neighborWidth: columnWidths[neighborIndex],
    });
  };
  const moveColumnResize = (event) => {
    if (!columnResize) return;
    const delta = event.clientX - columnResize.startX;
    const nextWidth = Math.max(
      55,
      Math.min(
        columnResize.startWidth + delta,
        columnResize.startWidth + columnResize.neighborWidth - 55,
      ),
    );
    const appliedDelta = nextWidth - columnResize.startWidth;
    setColumnWidths((widths) =>
      widths.map((width, index) =>
        index === columnResize.index
          ? nextWidth
          : index === columnResize.neighborIndex
            ? columnResize.neighborWidth - appliedDelta
            : width,
      ),
    );
  };
  const stopColumnResize = () => setColumnResize(null);

  const tableRows = current.items.length
    ? current.items.map((row, index) => (
        <ReportRow
          key={`${row.number}-${index}`}
          row={row}
          index={index}
          selected={selectedRows.has(index)}
          update={(patch) => updateRow(index, patch)}
          onSelect={(event) => selectRow(index, event)}
          onContextMenu={(event) => showRowContextMenu(event, index)}
        />
      ))
    : null;

  const headers = [
    "Code Report No",
    "Product Category",
    "Description",
    "Products Listed",
    "Latest Code",
    "Issue/Rev Date",
    "Expiration Date",
    "Download Process",
    "Status",
  ];
  const columnTotal = columnWidths.reduce((total, width) => total + width, 0);
  const columnStyle = (index) => ({
    width: `${((columnWidths[index] / columnTotal) * 100).toFixed(4)}%`,
  });
  return (
    <div className="desktop-window">
      <input
        ref={crpInputRef}
        className="hidden-file-input"
        type="file"
        accept=".crp,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        onChange={openDataFile}
      />
      <input
        ref={exportInputRef}
        className="hidden-file-input"
        type="file"
        accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        onChange={exportExcelFile}
      />
      <div className="titlebar">
        <div className="app-title">
          <span className="app-logo">CR</span>
          <b>Code Report Tracker</b>
        </div>
        <span className="title-caption">PDF report management workspace</span>
      </div>
      <div className="ribbon-tabs">
        <button
          className={`ribbon-tab ${activeRibbonTab === "Home" ? "active" : ""}`}
          onClick={() => setActiveRibbonTab("Home")}
        >
          Home
        </button>
        <button
          className={`ribbon-tab ${activeRibbonTab === "Help" ? "active" : ""}`}
          onClick={() => setActiveRibbonTab("Help")}
        >
          Help
        </button>
      </div>
      <div className="ribbon-body">
        {activeRibbonTab === "Home" ? (
          <>
            <div className="ribbon-group">
              <div className="ribbon-actions">
                <Button
                  icon="open"
                  onClick={() => crpInputRef.current?.click()}
                >
                  Open
                </Button>
                <Button icon="save" onClick={saveReport}>
                  Save
                </Button>
                <Button
                  icon="open"
                  onClick={() => crpInputRef.current?.click()}
                >
                  Import
                </Button>
                <Button icon="excel" onClick={exportTabsToFile} disabled={busy}>
                  Export to File
                </Button>
              </div>
              <small>File</small>
            </div>
            <div className="ribbon-group">
              <div className="ribbon-actions">
                <Button
                  icon="checkLink"
                  onClick={() => runCheck("check")}
                  disabled={busy}
                >
                  Check Link
                </Button>
                <SplitButton
                  icon="refresh"
                  onClick={() => runCheck("update")}
                  disabled={busy}
                  items={[
                    { label: "Update", action: () => runCheck("update") },
                    { label: "Update Local", action: () => runCheck("local") },
                  ]}
                >
                  Update
                </SplitButton>
                <SplitButton
                  icon="download"
                  onClick={downloadPdfs}
                  disabled={busy}
                  items={[
                    { label: "Download PDFs", action: downloadPdfs },
                    {
                      label: "Delete PDFs",
                      icon: "trash",
                      action: deletePdfs,
                    },
                  ]}
                >
                  Download
                </SplitButton>
                <Button icon="stop" onClick={cancelWork} disabled={!busy}>
                  Stop
                </Button>
              </div>
              <small>PDF Operations</small>
            </div>
            <div className="ribbon-group compact-group">
              <div className="ribbon-actions">
                <Button icon="settings" onClick={() => setShowSettings(true)}>
                  Settings
                </Button>
              </div>
              <small>Settings</small>
            </div>
            <div className="ribbon-spacer" />
          </>
        ) : (
          <div className="ribbon-group compact-group">
            <div className="ribbon-actions">
              <Button
                icon="github"
                onClick={() => {
                  const githubWindow = window.open(
                    "https://github.com/thang199801666/CodeReportTracker",
                    "_blank",
                    "noopener",
                  );
                  if (!githubWindow)
                    window.location.assign(
                      "https://github.com/thang199801666/CodeReportTracker",
                    );
                }}
              >
                GitHub
              </Button>
            </div>
            <small>Help</small>
          </div>
        )}
      </div>
      <main className="main-area">
        <div className="tabbar">
          {tabs.map((tab) => (
            <div
              className={`tab ${tab.id === selectedId ? "selected" : ""} ${draggedTabId === tab.id ? "dragging" : ""} ${dragOverTabId === tab.id ? "drag-over" : ""}`}
              key={tab.id}
              draggable
              onDragStart={(event) => startTabDrag(event, tab)}
              onDrag={(event) => moveTabDrag(event)}
              onDragOver={(event) => {
                event.preventDefault();
                moveTabOver(tab.id);
              }}
              onDrop={(event) => {
                event.preventDefault();
                endTabDrag();
              }}
              onDragEnd={endTabDrag}
              onClick={() => selectTab(tab.id)}
              onDoubleClick={() => setEditingTabId(tab.id)}
              onContextMenu={(event) => showTabContextMenu(event, tab.id)}
            >
              {editingTabId === tab.id ? (
                <input
                  className="tab-rename"
                  autoFocus
                  defaultValue={tab.header}
                  onBlur={(event) => renameTab(tab.id, event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter")
                      renameTab(tab.id, event.currentTarget.value);
                    if (event.key === "Escape") setEditingTabId(null);
                  }}
                />
              ) : (
                <span>{tab.header}</span>
              )}
              <button
                onClick={(event) => closeTabFromButton(event, tab.id)}
                aria-label={`Close ${tab.header}`}
              >
                <Icon name="close" />
              </button>
            </div>
          ))}
          <button className="new-tab" onClick={addTab}>
            <Icon name="plus" />
          </button>
        </div>
        <div className="table-wrap">
          <table style={{ width: "100%", minWidth: "0" }}>
            <colgroup>
              {columnWidths.map((_, index) => (
                <col key={index} style={columnStyle(index)} />
              ))}
            </colgroup>
            <thead>
              <tr>
                {headers.map((header, index) => (
                  <th key={header} style={columnStyle(index)}>
                    {header}
                    <span
                      className="column-grip"
                      onPointerDown={(event) => startColumnResize(event, index)}
                      onPointerMove={moveColumnResize}
                      onPointerUp={stopColumnResize}
                      onPointerCancel={stopColumnResize}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tableRows}
              <DraftRow
                row={draftRow}
                widths={columnWidths}
                total={columnTotal}
                update={updateDraftRow}
                onCommit={commitDraftRow}
              />
            </tbody>
          </table>
        </div>
      </main>
      <div
        className={`resize-handle ${resizingConsole ? "dragging" : ""}`}
        onPointerDown={startConsoleResize}
        onPointerMove={resizeConsole}
        onPointerUp={stopConsoleResize}
        onPointerCancel={stopConsoleResize}
        role="separator"
        aria-label="Resize Console"
        aria-orientation="horizontal"
      >
        <span />
      </div>
      <section className="console" style={{ height: consoleHeight }}>
        <div className="console-head">
          <span>
            <Icon name="terminal" /> Console
          </span>
          <span className="console-status">
            {busy ? "Running..." : "Ready"} · Drag the bar above to resize
          </span>
        </div>
        <div ref={consoleRef} className="console-output">
          {consoleText.split("\n").map((line, index) => {
            if (!line) return null;
            const tone =
              /failed|error|missing|skipped|unavailable|aborted|cancelled/i.test(
                line,
              )
                ? "error"
                : /downloaded|completed|finished|saved|exported|updated|created|deleted|found/i.test(
                      line,
                    )
                  ? "success"
                  : "normal";
            return (
              <div className={`console-line ${tone}`} key={`${index}-${line}`}>
                {line}
              </div>
            );
          })}
        </div>
      </section>
      <footer>
        <span>Code Report Tracker</span>
        <span>
          {current.items.length} row(s) · {tabs.length} tab(s)
        </span>
      </footer>
      {showSettings && (
        <SettingsModal
          settings={settings}
          onSave={saveSettings}
          onDirectorySelected={setDownloadDirectory}
          onClose={() => setShowSettings(false)}
        />
      )}
      {editingRowIndex !== null && current.items[editingRowIndex] && (
        <RowEditModal
          row={current.items[editingRowIndex]}
          onSave={(changes) => {
            updateRow(editingRowIndex, changes);
            setEditingRowIndex(null);
            log(`Edited ${current.items[editingRowIndex].number || "row"}.`);
          }}
          onClose={() => setEditingRowIndex(null)}
        />
      )}
      {contextMenu && (
        <div
          className="context-menu"
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onClick={(event) => event.stopPropagation()}
        >
          {contextMenu.tabId ? (
            <>
              <button
                onClick={() => {
                  setEditingTabId(contextMenu.tabId);
                  setContextMenu(null);
                }}
              >
                <Icon name="rename" /> Rename
              </button>
              <button
                onClick={() => {
                  const tabId = contextMenu.tabId;
                  setContextMenu(null);
                  deletePdfs(tabId);
                }}
              >
                <Icon name="trash" /> Delete PDF
              </button>
              <button
                onClick={() => {
                  const tabId = contextMenu.tabId;
                  setContextMenu(null);
                  openTabFolder(tabId);
                }}
              >
                <Icon name="folder" /> Open Folder
              </button>
            </>
          ) : (
            <>
              <button
                onClick={() => {
                  setContextMenu(null);
                  crpInputRef.current?.click();
                }}
              >
                <Icon name="open" /> Import
              </button>
              <button
                onClick={editSelectedRow}
                disabled={!contextMenu.row || selectedRows.size !== 1}
              >
                <Icon name="edit" /> Edit
              </button>
              <button
                onClick={deleteSelectedRows}
                disabled={!contextMenu.row || !selectedRows.size}
              >
                <Icon name="trash" /> Delete
              </button>
            </>
          )}
        </div>
      )}
      {dragPreview && (
        <div
          className="tab-drag-preview"
          style={{ left: dragPreview.x + 12, top: dragPreview.y + 12 }}
        >
          {dragPreview.header}
        </div>
      )}
    </div>
  );
}

function ReportRow({ row, index, selected, update, onSelect, onContextMenu }) {
  const displayValue = (field) => <span>{row[field] || "-"}</span>;
  const validLink = /^https?:\/\/\S+$/i.test(row.link || "");
  return (
    <tr
      className={selected ? "selected-row" : ""}
      onClick={onSelect}
      onContextMenu={onContextMenu}
    >
      <td
        className={`code-link ${row.checked ? (row.exists ? "checked exists" : "checked missing") : "unchecked"}`}
        title={row.link || "-"}
      >
        <a
          href={validLink ? row.link : undefined}
          target="_blank"
          rel="noreferrer"
          onClick={(event) => {
            if (!validLink) event.preventDefault();
          }}
          title={row.link || "-"}
        >
          {displayValue("number")}
        </a>
      </td>
      <td title={row.category || "-"}>{displayValue("category")}</td>
      <td title={row.description || "-"}>{displayValue("description")}</td>
      <td className="center" title={row.products || "-"}>
        {displayValue("products")}
      </td>
      <td
        className={`center changed ${row.oldLatest && row.oldLatest !== row.latest ? "yellow" : ""}`}
        title={`Old Data: ${row.oldLatest || "-"}\nNew Data: ${row.latest || "-"}`}
      >
        {displayValue("latest")}
      </td>
      <td
        className={`right changed ${row.oldIssue && row.oldIssue !== row.issue ? "yellow" : ""}`}
        title={`Old Data: ${row.oldIssue || "-"}\nNew Data: ${row.issue || "-"}`}
      >
        {displayValue("issue")}
      </td>
      <td
        className={`right changed ${row.oldExpiration && row.oldExpiration !== row.expiration ? "yellow" : ""}`}
        title={`Old Data: ${row.oldExpiration || "-"}\nNew Data: ${row.expiration || "-"}`}
      >
        {displayValue("expiration")}
      </td>
      <td title={`${row.progress}%`}>
        <div className="progress">
          <span style={{ width: `${row.progress}%` }} />
          <b>{row.progress}%</b>
        </div>
      </td>
      <td
        title={`Checked: ${row.checked ? "Yes" : "No"}; Updated: ${row.updated ? "Yes" : "No"}`}
      >
        <div className="statuses">
          <label className={row.checked ? "checked" : ""}>
            <i /> Checked
          </label>
          <label className={row.updated ? "updated" : ""}>
            <i /> Updated
          </label>
        </div>
      </td>
    </tr>
  );
}

function DraftRow({ row, widths, total, update, onCommit }) {
  const input = (field, placeholder = "") => (
    <input
      value={row[field]}
      placeholder={placeholder}
      onChange={(event) => update(field, event.target.value)}
      onKeyDown={(event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          onCommit();
        }
      }}
    />
  );
  const checkbox = (field, label) => (
    <label>
      <input
        type="checkbox"
        checked={row[field]}
        onChange={(event) => update(field, event.target.checked)}
      />{" "}
      {label}
    </label>
  );
  const style = (index) => ({
    width: `${((widths[index] / total) * 100).toFixed(4)}%`,
  });
  return (
    <tr className="draft-row">
      <td title={row.number || "New report"} style={style(0)}>
        {input("number", "New report")}
      </td>
      <td title={row.category || "-"} style={style(1)}>
        {input("category")}
      </td>
      <td title={row.description || "-"} style={style(2)}>
        {input("description")}
      </td>
      <td className="center" title={row.products || "-"} style={style(3)}>
        {input("products")}
      </td>
      <td className="center" title={row.latest || "-"} style={style(4)}>
        {input("latest")}
      </td>
      <td className="center" title={row.issue || "-"} style={style(5)}>
        {input("issue")}
      </td>
      <td className="center" title={row.expiration || "-"} style={style(6)}>
        {input("expiration")}
      </td>
      <td title="0%" style={style(7)}>
        <div className="progress">
          <span style={{ width: "0%" }} />
          <b>0%</b>
        </div>
      </td>
      <td
        title={`Checked: ${row.checked ? "Yes" : "No"}; Updated: ${row.updated ? "Yes" : "No"}`}
        style={style(8)}
      >
        <div className="draft-status">
          {checkbox("checked", "Checked")}
          {checkbox("updated", "Updated")}
        </div>
      </td>
    </tr>
  );
}

function RowEditModal({ row, onSave, onClose }) {
  const [draft, setDraft] = useState({ ...row });
  const fields = [
    ["number", "Code Report No"],
    ["link", "Link"],
    ["category", "Product Category"],
    ["description", "Description"],
    ["products", "Products Listed"],
    ["latest", "Latest Code"],
    ["issue", "Issue/Rev Date"],
    ["expiration", "Expiration Date"],
  ];
  return (
    <div className="modal-backdrop">
      <div className="modal row-edit-modal">
        <div className="modal-title">
          <b>Edit Row</b>
          <button onClick={onClose}>
            <Icon name="close" />
          </button>
        </div>
        <div className="row-edit-fields">
          {fields.map(([field, label]) => (
            <label key={field}>
              {label}
              <input
                value={draft[field] || ""}
                onChange={(event) =>
                  setDraft((value) => ({
                    ...value,
                    [field]: event.target.value,
                  }))
                }
              />
            </label>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={() => onSave(draft)}>
            Save
          </button>
        </div>
      </div>
    </div>
  );
}

function SettingsModal({ settings, onSave, onDirectorySelected, onClose }) {
  const [entries, setEntries] = useState(
    Object.entries(sources).map(([name, value]) => ({ name, ...value })),
  );
  const [options, setOptions] = useState(settings);
  const save = () => {
    onSave(options);
    onClose();
  };
  const chooseDirectory = async () => {
    if (isTauri) {
      try {
        const path = await invoke("pick_directory");
        if (path) {
          onDirectorySelected(path);
          setOptions((value) => ({ ...value, downloadDirectory: path }));
        }
      } catch (error) {
        window.alert(`Unable to select folder: ${error}`);
      }
      return;
    }
    if (!window.showDirectoryPicker) {
      window.alert("Folder selection is not supported by this browser.");
      return;
    }
    try {
      const handle = await window.showDirectoryPicker({ mode: "readwrite" });
      onDirectorySelected(handle);
      setOptions((value) => ({ ...value, downloadDirectory: handle.name }));
    } catch (error) {
      if (error.name !== "AbortError")
        window.alert(`Unable to select folder: ${error.message}`);
    }
  };
  return (
    <div className="modal-backdrop">
      <div className="modal">
        <div className="modal-title">
          <b>Settings</b>
          <button onClick={onClose}>
            <Icon name="close" />
          </button>
        </div>
        <div className="setting-options">
          <label>
            Worker Count
            <input
              type="number"
              min="1"
              max={MAX_WORKERS}
              value={options.workerCount}
              onChange={(event) =>
                setOptions((value) => ({
                  ...value,
                  workerCount: event.target.value,
                }))
              }
            />
          </label>
          <label>
            Download Directory
            <div className="directory-picker">
              <input
                value={options.downloadDirectory}
                placeholder="Enter full path or browse"
                onChange={(event) =>
                  setOptions((value) => ({
                    ...value,
                    downloadDirectory: event.target.value,
                  }))
                }
              />
              <button type="button" onClick={chooseDirectory}>
                Browse
              </button>
            </div>
            <small className="setting-note">
              {isTauri
                ? "Tauri uses the selected folder's full path."
                : "Browsers expose the selected folder name only. Enter an absolute path manually if required."}
            </small>
          </label>
        </div>
        <div className="settings-box">
          <div className="settings-heading">Link Settings (table)</div>
          <table className="settings-table">
            <thead>
              <tr>
                <th>Web Name</th>
                <th>Type</th>
                <th>Link</th>
                <th>PDF Folder Link</th>
              </tr>
            </thead>
            <tbody>
              {entries.map((entry, index) => (
                <tr key={entry.name}>
                  <td>
                    <input
                      value={entry.name}
                      onChange={(e) =>
                        setEntries((all) =>
                          all.map((item, i) =>
                            i === index
                              ? { ...item, name: e.target.value }
                              : item,
                          ),
                        )
                      }
                    />
                  </td>
                  <td>
                    <select
                      value={entry.type}
                      onChange={(e) =>
                        setEntries((all) =>
                          all.map((item, i) =>
                            i === index
                              ? { ...item, type: e.target.value }
                              : item,
                          ),
                        )
                      }
                    >
                      <option>ER</option>
                      <option>ESR</option>
                      <option>Other</option>
                      <option>Folder</option>
                    </select>
                  </td>
                  <td>
                    <input
                      value={entry.link}
                      onChange={(e) =>
                        setEntries((all) =>
                          all.map((item, i) =>
                            i === index
                              ? { ...item, link: e.target.value }
                              : item,
                          ),
                        )
                      }
                    />
                  </td>
                  <td>
                    <input
                      value={entry.pdfFolder}
                      onChange={(e) =>
                        setEntries((all) =>
                          all.map((item, i) =>
                            i === index
                              ? { ...item, pdfFolder: e.target.value }
                              : item,
                          ),
                        )
                      }
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="modal-actions">
          <button onClick={save}>Save</button>
          <button className="secondary" onClick={onClose}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

class AppErrorBoundary extends React.Component {
  state = { error: null };

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error("Code Report Tracker crashed while rendering.", error, info);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <main className="app-error">
        <h1>Code Report Tracker could not render this view</h1>
        <p>{this.state.error.message || "Unexpected application error."}</p>
        <button onClick={() => window.location.reload()}>Reload App</button>
      </main>
    );
  }
}

createRoot(document.getElementById("root")).render(
  <AppErrorBoundary>
    <App />
  </AppErrorBoundary>,
);
