import { useEffect, useRef, useState } from "react";

const getRenderPdfEndpoint = () => {
  const explicitBase =
    (import.meta.env.VITE_PDF_API_BASE_URL || import.meta.env.VITE_API_BASE_URL || "")
      .trim()
      .replace(/\/+$/, "");

  return explicitBase ? `${explicitBase}/api/render-mpo-pdf` : "/api/render-mpo-pdf";
};

const PrintPreview = ({ html, csv, title, onClose }) => {
  const [tab, setTab] = useState("preview");
  const [copied, setCopied] = useState(false);
  const [pdfBusy, setPdfBusy] = useState(false);
  const [pdfError, setPdfError] = useState("");
  const [pdfUrl, setPdfUrl] = useState("");
  const [pdfBlob, setPdfBlob] = useState(null);
  const iframeRef = useRef(null);

  const safeName = (t) => (t || "MPO").replace(/[^a-z0-9\-_. ]/gi, "_").slice(0, 80);

  const revokePdfUrl = () => {
    setPdfUrl((current) => {
      if (current) URL.revokeObjectURL(current);
      return "";
    });
  };

  const fetchPdfBlob = async () => {
    const endpoint = getRenderPdfEndpoint();

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ html, title }),
    });

    const contentType = response.headers.get("content-type") || "";

    if (!response.ok) {
      let message = "";

      try {
        const data = await response.json();
        message = data?.details || data?.error || "";
      } catch {
        try {
          message = await response.text();
        } catch {
          message = "";
        }
      }

      throw new Error(message || `PDF render failed (${response.status}).`);
    }

    if (!contentType.toLowerCase().includes("application/pdf")) {
      const badText = await response.text();
      throw new Error(`Server did not return a PDF. Response was: ${badText.slice(0, 300)}`);
    }

    const blob = await response.blob();

    if (blob.size < 1000) {
      throw new Error("Downloaded file is too small to be a valid PDF.");
    }

    return blob;
  };

  const ensurePdfReady = async () => {
    if (pdfBlob && pdfUrl) {
      return { blob: pdfBlob, url: pdfUrl };
    }

    setPdfBusy(true);
    setPdfError("");

    try {
      const blob = await fetchPdfBlob();
      revokePdfUrl();
      const url = URL.createObjectURL(blob);
      setPdfBlob(blob);
      setPdfUrl(url);
      return { blob, url };
    } catch (error) {
      const rawMessage = error?.message || "Failed to generate PDF preview.";
      const message =
        rawMessage.includes("Failed to fetch")
          ? "PDF API is not reachable. Deploy api/render-mpo-pdf.js or set VITE_PDF_API_BASE_URL."
          : rawMessage;

      setPdfError(message);
      throw new Error(message);
    } finally {
      setPdfBusy(false);
    }
  };

  useEffect(() => {
    setPdfBlob(null);
    setPdfError("");
    revokePdfUrl();

    if (tab === "preview") {
      ensurePdfReady().catch(() => {});
    }

    return () => {
      revokePdfUrl();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [html, title]);

  useEffect(() => {
    if (tab === "preview" && !pdfUrl && !pdfBusy) {
      ensurePdfReady().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab]);

  const handlePrint = async () => {
    try {
      const { url } = await ensurePdfReady();
      const printFrame = iframeRef.current;

      if (printFrame?.contentWindow) {
        printFrame.contentWindow.focus();
        printFrame.contentWindow.print();
        return;
      }

      window.open(url, "_blank", "noopener,noreferrer");
    } catch (error) {
      console.error("PDF print failed:", error);
    }
  };

  const handleDownloadPDF = async () => {
    try {
      const { blob } = await ensurePdfReady();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${safeName(title)}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
    } catch (error) {
      console.error("Server PDF download failed:", error);
      alert(`PDF download failed: ${error.message}`);
    }
  };

  const handleDownloadHTML = () => {
    const clean = html.replace(/<script>[\s\S]*?window\.addEventListener[\s\S]*?<\/script>/, "");
    const blob = new Blob([clean], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${safeName(title)}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  const handleDownloadCSV = () => {
    if (!csv) return;
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${safeName(title)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  const handleCopy = () => {
    if (!csv) return;
    navigator.clipboard
      .writeText(csv)
      .then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 2500);
      })
      .catch(() => {
        const ta = document.getElementById("csv-ta");
        if (ta) {
          ta.select();
          document.execCommand("copy");
          setCopied(true);
          setTimeout(() => setCopied(false), 2500);
        }
      });
  };

  const btnStyle = (bg, color, border) => ({
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "7px 16px",
    background: bg,
    color,
    border: `1px solid ${border}`,
    borderRadius: 8,
    fontFamily: "'Syne',sans-serif",
    fontWeight: 700,
    fontSize: 12,
    cursor: "pointer",
  });

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 2000,
        background: "rgba(0,0,0,.88)",
        display: "flex",
        flexDirection: "column",
        backdropFilter: "blur(6px)",
      }}
    >
      <div
        style={{
          background: "#0e1118",
          borderBottom: "1px solid rgba(255,255,255,.1)",
          padding: "11px 18px",
          display: "flex",
          alignItems: "center",
          gap: 10,
          flexShrink: 0,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            fontFamily: "'Syne',sans-serif",
            fontWeight: 700,
            fontSize: 14,
            flex: 1,
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
          }}
        >
          {title}
        </div>

        <div
          style={{
            display: "flex",
            gap: 0,
            background: "#141824",
            borderRadius: 7,
            overflow: "hidden",
            border: "1px solid rgba(255,255,255,.1)",
          }}
        >
          {[["preview", "📄 Preview / Print"], ["csv", "📊 CSV / Excel"]].map(([t, l]) => (
            <button
              key={t}
              onClick={() => setTab(t)}
              style={{
                padding: "7px 14px",
                border: "none",
                background: tab === t ? "#f0a500" : "transparent",
                color: tab === t ? "#000" : "#8b93a7",
                fontFamily: "'Syne',sans-serif",
                fontWeight: 600,
                fontSize: 12,
                cursor: "pointer",
              }}
            >
              {l}
            </button>
          ))}
        </div>

        {tab === "preview" && (
          <>
            <button
              onClick={handleDownloadPDF}
              disabled={pdfBusy}
              style={{
                ...btnStyle("rgba(34,197,94,.15)", "#22c55e", "rgba(34,197,94,.35)"),
                opacity: pdfBusy ? 0.65 : 1,
                cursor: pdfBusy ? "wait" : "pointer",
              }}
            >
              {pdfBusy ? "⏳ Building PDF..." : "⬇ Download PDF"}
            </button>

            <button
              onClick={handlePrint}
              disabled={pdfBusy || !pdfUrl}
              style={{
                ...btnStyle("#0A1F44", "#D4870A", "#D4870A"),
                opacity: pdfBusy ? 0.65 : 1,
                cursor: pdfBusy ? "wait" : "pointer",
              }}
            >
              🖨 Print
            </button>

            <button
              onClick={handleDownloadHTML}
              style={btnStyle("rgba(139,92,246,.15)", "#a78bfa", "rgba(139,92,246,.4)")}
            >
              ⬇ Download HTML
            </button>
          </>
        )}

        {tab === "csv" && csv && (
          <>
            <button
              onClick={handleDownloadCSV}
              style={btnStyle("rgba(34,197,94,.15)", "#22c55e", "rgba(34,197,94,.4)")}
            >
              ⬇ Download CSV
            </button>

            <button
              onClick={handleCopy}
              style={btnStyle(
                copied ? "rgba(34,197,94,.25)" : "rgba(255,255,255,.05)",
                copied ? "#22c55e" : "#8b93a7",
                copied ? "rgba(34,197,94,.4)" : "rgba(255,255,255,.12)"
              )}
            >
              {copied ? "✓ Copied!" : "📋 Copy to Clipboard"}
            </button>
          </>
        )}

        <button
          onClick={onClose}
          style={btnStyle("rgba(239,68,68,.15)", "#ef4444", "rgba(239,68,68,.3)")}
        >
          ✕ Close
        </button>
      </div>

      {tab === "preview" && (
        <div
          style={{
            flex: 1,
            overflow: "auto",
            background: "#ccc",
            display: "flex",
            justifyContent: "center",
            padding: 24,
          }}
        >
          {pdfBusy && !pdfUrl ? (
            <div
              style={{
                width: "100%",
                maxWidth: 900,
                minHeight: 700,
                background: "#fff",
                boxShadow: "0 8px 40px rgba(0,0,0,.5)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: "'Syne',sans-serif",
                fontWeight: 700,
                color: "#1f2937",
              }}
            >
              Building preview PDF...
            </div>
          ) : pdfError ? (
            <div
              style={{
                width: "100%",
                maxWidth: 900,
                minHeight: 700,
                background: "#fff",
                boxShadow: "0 8px 40px rgba(0,0,0,.5)",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                padding: 24,
                textAlign: "center",
                color: "#7f1d1d",
                gap: 12,
              }}
            >
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18 }}>
                Could not load PDF preview
              </div>
              <div style={{ maxWidth: 520, fontSize: 14 }}>{pdfError}</div>
              <button
                onClick={() => ensurePdfReady().catch(() => {})}
                style={btnStyle("rgba(34,197,94,.15)", "#22c55e", "rgba(34,197,94,.35)")}
              >
                Retry
              </button>
            </div>
          ) : pdfUrl ? (
            <iframe
              ref={iframeRef}
              src={pdfUrl}
              style={{
                width: "100%",
                maxWidth: 980,
                height: "calc(100vh - 150px)",
                border: "none",
                boxShadow: "0 8px 40px rgba(0,0,0,.5)",
                background: "#fff",
              }}
              title="MPO PDF Preview"
            />
          ) : (
            <div
              style={{
                width: "100%",
                maxWidth: 900,
                minHeight: 700,
                background: "#fff",
                boxShadow: "0 8px 40px rgba(0,0,0,.5)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: "'Syne',sans-serif",
                fontWeight: 700,
                color: "#1f2937",
              }}
            >
              Preparing preview...
            </div>
          )}
        </div>
      )}

      {tab === "csv" && csv && (
        <div
          style={{
            flex: 1,
            overflow: "hidden",
            display: "flex",
            flexDirection: "column",
            padding: 20,
            gap: 10,
          }}
        >
          <div
            style={{
              background: "rgba(34,197,94,.08)",
              border: "1px solid rgba(34,197,94,.2)",
              borderRadius: 9,
              padding: "11px 15px",
              fontSize: 13,
              color: "#22c55e",
            }}
          >
            💡 Click <strong>Download CSV</strong> to save, then open in Excel or Google Sheets. Or{" "}
            <strong>Copy to Clipboard</strong> and paste directly.
          </div>

          <textarea
            id="csv-ta"
            readOnly
            value={csv}
            style={{
              flex: 1,
              background: "#0e1118",
              border: "1px solid rgba(255,255,255,.1)",
              borderRadius: 9,
              padding: 14,
              color: "#8b93a7",
              fontFamily: "monospace",
              fontSize: 11,
              resize: "none",
              outline: "none",
              lineHeight: 1.6,
            }}
          />
        </div>
      )}
    </div>
  );
};

export default PrintPreview;
