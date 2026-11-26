import 'dotenv/config';
import express from 'express';
import { createServer as createHttpServer } from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI, Type } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const httpServer = createHttpServer(app);
  const port = Number(process.env.PORT || 3000);

  app.use(express.json({ limit: '25mb' }));

  // Health check
  app.get('/api/health', (_req, res) => {
    res.json({ ok: true });
  });

  // Server-side Multimodal Scanned PDF / Image Repayment Schedule Extraction Fallback
  app.post('/api/repayment-schedule/parse', async (req, res) => {
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(503).json({
          error: 'GEMINI_API_KEY is not configured for OCR fallback on scanned PDFs/images. Use text-based PDF, Excel (.xlsx/.xls), CSV, or Paste Table fallback.',
        });
      }

      const { fileName, mimeType, base64Data } = req.body as {
        fileName?: string;
        mimeType?: string;
        base64Data?: string;
      };

      if (!base64Data || !mimeType) {
        return res.status(400).json({ error: 'Missing uploaded file content (base64Data / mimeType).' });
      }

      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          },
        },
      });

      const prompt = `Extract the complete bank/client Repayment Schedule from the attached document (${fileName || 'schedule'}).
Return every Disbursement and Repayment installment row across all pages without omitting or summarizing any installment.
Important rules:
1. Extract exact numeric amounts (strip Indian/Western commas like "24,30,778.97" -> 2430778.97). Preserve 2 decimal places accurately.
2. Keep the original EMI Date text verbatim (e.g., "01/02/2027").
3. Do NOT swap Inst. Principal and Inst. Profit. Include Gestation Payment, Over due Gest., and Overdue Profit when present.
4. For the Disbursement row, set isDisbursement=true, instPrincipal=0, instProfit=0, and balance=disbursed principal amount.
5. Extract stated header metadata (investmentMade, interval, numberOfInstallments, gestationPeriod, disbursementDate, firstInstallmentDate, expiryDate, lastInstallmentMinimized) and footer Total row (statedTotalPrincipal, statedTotalProfit) if present.`;

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: {
          parts: [
            {
              inlineData: {
                mimeType,
                data: base64Data,
              },
            },
            { text: prompt },
          ],
        },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              investmentMade: { type: Type.NUMBER },
              interval: { type: Type.STRING },
              numberOfInstallments: { type: Type.INTEGER },
              gestationPeriod: { type: Type.STRING },
              disbursementDate: { type: Type.STRING },
              firstInstallmentDate: { type: Type.STRING },
              expiryDate: { type: Type.STRING },
              lastInstallmentMinimized: { type: Type.STRING },
              statedTotalPrincipal: { type: Type.NUMBER },
              statedTotalProfit: { type: Type.NUMBER },
              rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    eventType: { type: Type.STRING },
                    emiDate: { type: Type.STRING },
                    instPrincipal: { type: Type.NUMBER },
                    instProfit: { type: Type.NUMBER },
                    gestationPayment: { type: Type.NUMBER },
                    overdueGestation: { type: Type.NUMBER },
                    overdueProfit: { type: Type.NUMBER },
                    balance: { type: Type.NUMBER },
                    total: { type: Type.NUMBER },
                    isDisbursement: { type: Type.BOOLEAN },
                  },
                  required: ['emiDate', 'instPrincipal', 'instProfit'],
                },
              },
            },
            required: ['rows'],
          },
        },
      });

      const rawJson = (response.text ?? '{}').trim();
      const parsed = JSON.parse(rawJson);
      return res.json(parsed);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to extract repayment schedule via server OCR.';
      return res.status(500).json({ error: message });
    }
  });

  const isProd = process.env.NODE_ENV === 'production';
  if (!isProd) {
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: false,
        ws: { server: httpServer },
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  httpServer.listen(port, '0.0.0.0', () => {
    console.log(`Financials with Tuhin server listening on http://0.0.0.0:${port}`);
  });
}

startServer().catch(err => {
  console.error('Server startup failed:', err);
  process.exit(1);
});
