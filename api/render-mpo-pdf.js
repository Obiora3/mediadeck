import fs from "fs";
import puppeteer from "puppeteer-core";
import chromium from "@sparticuz/chromium-min";

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
  if (process.env.VERCEL) {
    return await chromium.executablePath();
  }

  const localChrome = findLocalChrome();
  if (localChrome) return localChrome;

  throw new Error(
    "Chrome executable not found locally. Set PUPPETEER_EXECUTABLE_PATH or install Chrome for Puppeteer."
  );
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

  let browser;

  try {
    const executablePath = await getExecutablePath();

    browser = await puppeteer.launch({
      executablePath,
      headless: true,
      args: process.env.VERCEL ? chromium.args : [],
      defaultViewport: {
        width: 900,
        height: 1400,
        deviceScaleFactor: 1,
      },
    });

    const page = await browser.newPage();

    // Avoid hanging forever on slow/external assets.
    page.setDefaultNavigationTimeout(0);
    page.setDefaultTimeout(0);

    // Stop external requests from blocking PDF generation.
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

      // Abort external network requests so setContent doesn't hang.
      request.abort();
    });

    await page.emulateMediaType("screen");

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

      await new Promise((resolve) => setTimeout(resolve, 400));
    });

    const dimensions = await page.evaluate(() => {
      const body = document.body;
      const root = document.documentElement;

      return {
        width: Math.ceil(
          Math.max(
            body.scrollWidth,
            body.offsetWidth,
            body.clientWidth,
            root.scrollWidth,
            root.offsetWidth,
            root.clientWidth
          )
        ),
        height: Math.ceil(
          Math.max(
            body.scrollHeight,
            body.offsetHeight,
            body.clientHeight,
            root.scrollHeight,
            root.offsetHeight,
            root.clientHeight
          )
        ),
      };
    });

    const pdfWidth = Math.max(900, dimensions.width);
    const pdfHeight = Math.max(1200, dimensions.height + 24);

    const pdfBuffer = await page.pdf({
      printBackground: true,
      width: `${pdfWidth}px`,
      height: `${pdfHeight}px`,
      margin: {
        top: "0px",
        right: "0px",
        bottom: "0px",
        left: "0px",
      },
      preferCSSPageSize: false,
      pageRanges: "1",
    });

    const safeTitle = String(title || "MPO")
      .replace(/[^a-z0-9\-_. ]/gi, "_")
      .slice(0, 80);

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
    if (browser) {
      await browser.close().catch(() => {});
    }
  }
}