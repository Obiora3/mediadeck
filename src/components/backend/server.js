import express from "express";
import cors from "cors";
import renderMpoPdf from "./render-mpo-pdf.js";

const app = express();
const PORT = 4000;

// Allow your frontend to call this backend
app.use(cors({
  origin: [
    "http://localhost:5173",
    "http://localhost:3000",
  ],
}));

// Accept large HTML payloads
app.use(express.json({ limit: "10mb" }));

// Wrap the handler so it works with Express
app.post("/api/render-mpo-pdf", async (req, res) => {
  return renderMpoPdf(req, res);
});

app.get("/", (_req, res) => {
  res.send("PDF backend is running.");
});

app.listen(PORT, () => {
  console.log(`PDF backend running on http://localhost:${PORT}`);
});