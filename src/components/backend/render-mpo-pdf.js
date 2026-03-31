import puppeteer from "puppeteer";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  const { html, title } = req.body || {};

  if (!html || typeof html !== "string") {
    return res.status(400).json({ error: "Missing HTML payload." });
  }

  let browser;

  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    const page = await browser.newPage();

    // Match the on-screen preview conditions as closely as possible.
    await page.setViewport({
      width: 900,
      height: 1400,
      deviceScaleFactor: 1,
    });

    await page.emulateMediaType("screen");
    await page.setContent(html, {
      waitUntil: ["domcontentloaded", "networkidle0"],
    });

    // Wait for fonts/images/layout to settle before measuring.
    await page.evaluate(async () => {
      try {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }
      } catch {}
      await new Promise((resolve) => setTimeout(resolve, 250));
    });

    const dimensions = await page.evaluate(() => {
      const body = document.body;
      const root = document.documentElement;

      const width = Math.max(
        body.scrollWidth,
        body.offsetWidth,
        body.clientWidth,
        root.scrollWidth,
        root.offsetWidth,
        root.clientWidth
      );

      const height = Math.max(
        body.scrollHeight,
        body.offsetHeight,
        body.clientHeight,
        root.scrollHeight,
        root.offsetHeight,
        root.clientHeight
      );

      return {
        width: Math.ceil(width),
        height: Math.ceil(height),
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
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${safeTitle}.pdf"`);
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
