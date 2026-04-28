import fs from "fs";
import crypto from "crypto";
import puppeteer from "puppeteer-core";
import chromium from "@sparticuz/chromium";

const isVercel = !!process.env.VERCEL;

let executablePathPromise = null;
let browserPromise = null;

const pdfCache = new Map();
const PDF_CACHE_TTL = 5 * 60 * 1000;
const PDF_CACHE_MAX_ITEMS = 20;

const findLocalChrome = () => {
  const candidates = [
    process.env.PUPPETEER_EXECUTABLE_PATH,
    process.env.PUPPETEER_BROWSER_PATH,
    process.env.USERPROFILE
      ? `${process.env.USERPROFILE}\\.cache\\puppeteer\\chrome\\win64-146.0.7680.153\\chrome-win64\\chrome.exe`
      : "",
    process.env.LOCALAPPDATA
      ? `${process.env.LOCALAPPDATA}\\Google\\Chrome\\Application\\chrome.exe`
      : "",
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  ].filter(Boolean);

  return candidates.find((filePath) => fs.existsSync(filePath)) || null;
};

const getExecutablePath = async () => {
  if (!executablePathPromise) {
    executablePathPromise = (async () => {
      if (isVercel) {
        return await chromium.executablePath();
      }

      const localChrome = findLocalChrome();
      if (localChrome) return localChrome;

      throw new Error(
        "Chrome executable not found locally. Set PUPPETEER_EXECUTABLE_PATH or install Chrome for Puppeteer."
      );
    })();
  }

  return executablePathPromise;
};

const getBrowser = async () => {
  if (!browserPromise) {
    browserPromise = (async () => {
      const executablePath = await getExecutablePath();

      return puppeteer.launch({
        executablePath,
        headless: true,
        args: isVercel ? chromium.args : [],
        defaultViewport: {
          width: 900,
          height: 1400,
          deviceScaleFactor: 1,
        },
      });
    })();
  }

  return browserPromise;
};

const safeTitleFrom = (title) =>
  String(title || "MPO")
    .replace(/[^a-z0-9\-_. ]/gi, "_")
    .slice(0, 80);

const makeCacheKey = (html, title) =>
  crypto.createHash("sha1").update(`${title || "MPO"}::${html}`).digest("hex");

const pruneCache = () => {
  const now = Date.now();

  for (const [key, entry] of pdfCache.entries()) {
    if (now - entry.createdAt > PDF_CACHE_TTL) {
      pdfCache.delete(key);
    }
  }

  if (pdfCache.size <= PDF_CACHE_MAX_ITEMS) return;

  const entries = Array.from(pdfCache.entries()).sort(
    (a, b) => a[1].createdAt - b[1].createdAt
  );

  while (entries.length > PDF_CACHE_MAX_ITEMS) {
    const [oldestKey] = entries.shift();
    pdfCache.delete(oldestKey);
  }
};

const getCachedPdf = (cacheKey) => {
  const entry = pdfCache.get(cacheKey);
  if (!entry) return null;

  if (Date.now() - entry.createdAt > PDF_CACHE_TTL) {
    pdfCache.delete(cacheKey);
    return null;
  }

  return entry.buffer;
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  let payload = req.body || {};

  if (typeof payload === "string") {
    try {
      payload = JSON.parse(payload);
    } catch {
      return res.status(400).json({ error: "Invalid JSON payload." });
    }
  }

  const { html, title } = payload;

  if (!html || typeof html !== "string") {
    return res.status(400).json({ error: "Missing HTML payload." });
  }

  const safeTitle = safeTitleFrom(title);
  const cacheKey = makeCacheKey(html, title);
  const cachedPdf = getCachedPdf(cacheKey);

  if (cachedPdf) {
    res.status(200);
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${safeTitle}.pdf"`);
    res.setHeader("Content-Length", String(cachedPdf.length));
    return res.end(cachedPdf);
  }

  let page;

  try {
    const browser = await getBrowser();
    page = await browser.newPage();

    page.setDefaultNavigationTimeout(0);
    page.setDefaultTimeout(0);

    await page.setRequestInterception(true);
    page.on("request", (request) => {
      const url = request.url();

      const allowed =
        url.startsWith("data:") ||
        url.startsWith("blob:") ||
        url.startsWith("file:") ||
        url.startsWith("about:blank");

      if (allowed) {
        request.continue();
        return;
      }

      request.abort();
    });

    await page.emulateMediaType("print");

    await page.setContent(html, {
      waitUntil: "domcontentloaded",
      timeout: 0,
    });

    await page.evaluate(async () => {
      try {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }
      } catch {}

      await new Promise((resolve) => setTimeout(resolve, 150));
    });

    const pdfBuffer = await page.pdf({
      printBackground: true,
      format: "A4",
      margin: {
        top: "0px",
        right: "0px",
        bottom: "0px",
        left: "0px",
      },
      preferCSSPageSize: true,
    });

    pdfCache.set(cacheKey, {
      buffer: pdfBuffer,
      createdAt: Date.now(),
    });
    pruneCache();

    res.status(200);
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${safeTitle}.pdf"`);
    res.setHeader("Content-Length", String(pdfBuffer.length));
    return res.end(pdfBuffer);
  } catch (error) {
    console.error("render-mpo-pdf failed:", error);
    return res.status(500).json({
      error: "Failed to render preview PDF.",
      details: error.message,
    });
  } finally {
    if (page) {
      await page.close().catch(() => {});
    }
  }
}
