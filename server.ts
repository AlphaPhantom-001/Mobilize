import express from 'express';
import path from 'path';
import dotenv from 'dotenv';
import { GoogleGenAI, ThinkingLevel, Type } from '@google/genai';
import { createServer as createViteServer } from 'vite';

dotenv.config();

const PORT = Number(process.env.PORT) || 3000;
const EXTRACT_MODELS = ['gemini-4-hydrogen', 'gemini-3.1-flash-lite'] as const;

const CARD_SYSTEM_INSTRUCTION = 'Return only what is printed. Use ? for any unreadable character. Never guess.';

const CARD_PROPERTIES = {
  card_type: {
    type: Type.STRING,
    enum: ['pvc', 'evoter_printout', 'text_card', 'old_card', 'not_a_card'],
  },
  name: { type: Type.STRING },
  code: { type: Type.STRING },
  delim_lga: { type: Type.STRING },
  delim_ward: { type: Type.STRING },
  vin: { type: Type.STRING },
  rotation_needed: { type: Type.INTEGER },
  legible: {
    type: Type.OBJECT,
    properties: {
      name: { type: Type.STRING, enum: ['full', 'partial', 'none'] },
      code: { type: Type.STRING, enum: ['full', 'partial', 'none'] },
    },
    required: ['name', 'code'],
  },
};

function getAiClient() {
  return new GoogleGenAI({
    apiKey: process.env.GEMINI_API_KEY,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

function isRateLimitError(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err || '');
  return msg.includes('429') || msg.includes('RESOURCE_EXHAUSTED') || msg.toLowerCase().includes('quota');
}

function extractRetrySeconds(err: unknown): number {
  const msg = err instanceof Error ? err.message : String(err || '');
  const m = msg.match(/"retryDelay"\s*:\s*"(\d+)s"/i) || msg.match(/retry in\s+(?:(\d+)h)?(?:(\d+)m)?(\d+)/i);
  if (!m) return 60;
  if (m.length === 2) return Number(m[1]) || 60;
  const hrs = Number(m[1] || 0);
  const mins = Number(m[2] || 0);
  const secs = Number(m[3] || 0);
  return hrs * 3600 + mins * 60 + secs;
}

async function generateWithFallback(
  ai: GoogleGenAI,
  parts: Array<{ text: string } | { inlineData: { mimeType: string; data: string } }>,
  responseSchema: Record<string, unknown>
) {
  let lastErr: unknown = null;

  for (const model of EXTRACT_MODELS) {
    try {
      const response = await ai.models.generateContent({
        model,
        contents: { parts },
        config: {
          systemInstruction: CARD_SYSTEM_INSTRUCTION,
          responseMimeType: 'application/json',
          responseSchema,
          ...(model === 'gemini-4-hydrogen'
            ? { thinkingConfig: { thinkingLevel: ThinkingLevel.LOW } }
            : {}),
        },
      });
      return { model, text: (response.text ?? '{}').trim() };
    } catch (err) {
      lastErr = err;
      if (isRateLimitError(err)) {
        console.warn(`Gemini model ${model} rate/quota limited; trying next fallback if available.`);
        continue;
      }
      throw err;
    }
  }

  throw lastErr;
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '35mb' }));

  app.post('/api/extract-card', async (req, res) => {
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(503).json({ ok: false, error: 'GEMINI_API_KEY is not configured' });
      }

      const { imageBase64, mimeType } = req.body || {};
      if (!imageBase64 || typeof imageBase64 !== 'string') {
        return res.status(400).json({ ok: false, error: 'imageBase64 is required' });
      }

      const cleanBase64 = imageBase64.replace(/^data:[^;]+;base64,/, '').trim();
      const ai = getAiClient();

      const { model, text } = await generateWithFallback(
        ai,
        [
          {
            inlineData: {
              mimeType: mimeType || 'image/jpeg',
              data: cleanBase64,
            },
          },
          {
            text: [
              'Extract the details printed on this card or e-voter printout image.',
              'Return only what is printed. Use ? for any unreadable character. Never guess.',
              'Fields:',
              '- card_type: one of "pvc", "evoter_printout", "text_card", "old_card", "not_a_card"',
              '- name: full printed name (empty string if unreadable or missing)',
              '- code: printed delimitation CODE (e.g. 08-13-09-008), empty string if not visible',
              '- delim_lga: printed LGA / Local Government text (empty string if not visible)',
              '- delim_ward: printed Ward / Registration Area text (empty string if not visible)',
              '- vin: printed VIN (Voter Identification Number), empty string if not visible',
              '- rotation_needed: 0, 90, 180, or 270 clockwise degrees needed to make the card upright',
              '- legible: { name: "full"|"partial"|"none", code: "full"|"partial"|"none" }',
            ].join('\n'),
          },
        ],
        {
          type: Type.OBJECT,
          properties: CARD_PROPERTIES,
          required: ['card_type', 'name', 'code', 'delim_lga', 'delim_ward', 'vin', 'rotation_needed', 'legible'],
        }
      );

      const parsed = JSON.parse(text);
      return res.json({ ok: true, model, data: parsed });
    } catch (err: unknown) {
      if (isRateLimitError(err)) {
        const retryAfterSec = extractRetrySeconds(err);
        return res.status(429).json({
          ok: false,
          code: 'QUOTA_EXHAUSTED',
          retryAfterSec,
          error: 'Gemini free-tier quota limit reached; switching to local OCR.',
        });
      }
      const message = err instanceof Error ? err.message : 'Extraction failed';
      console.error('/api/extract-card error:', message);
      return res.status(500).json({ ok: false, error: message });
    }
  });

  app.post('/api/extract-cards-batch', async (req, res) => {
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(503).json({ ok: false, error: 'GEMINI_API_KEY is not configured' });
      }

      const cards = Array.isArray(req.body?.cards) ? req.body.cards : [];
      if (cards.length === 0) {
        return res.status(400).json({ ok: false, error: 'cards array is required' });
      }

      const ai = getAiClient();
      const parts: Array<{ text: string } | { inlineData: { mimeType: string; data: string } }> = [];

      for (const card of cards.slice(0, 10)) {
        const idx = Number(card.index ?? 0);
        const cleanBase64 = String(card.imageBase64 || '').replace(/^data:[^;]+;base64,/, '').trim();
        if (!cleanBase64) continue;
        parts.push({ text: `Card index ${idx}:` });
        parts.push({
          inlineData: {
            mimeType: card.mimeType || 'image/jpeg',
            data: cleanBase64,
          },
        });
      }

      parts.push({
        text: [
          'Extract the printed details for EVERY labeled card image above, returning one JSON object per card index.',
          'Return only what is printed. Use ? for any unreadable character. Never guess.',
          'For each card return:',
          '- index: the integer Card index shown before the image',
          '- card_type: one of "pvc", "evoter_printout", "text_card", "old_card", "not_a_card"',
          '- name: full printed name (empty string if unreadable or missing)',
          '- code: printed delimitation CODE (e.g. 08-13-09-008), empty string if not visible',
          '- delim_lga: printed LGA / Local Government text (empty string if not visible)',
          '- delim_ward: printed Ward / Registration Area text (empty string if not visible)',
          '- vin: printed VIN (Voter Identification Number), empty string if not visible',
          '- rotation_needed: 0, 90, 180, or 270 clockwise degrees needed to make the card upright',
          '- legible: { name: "full"|"partial"|"none", code: "full"|"partial"|"none" }',
        ].join('\n'),
      });

      const { model, text } = await generateWithFallback(ai, parts, {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            index: { type: Type.INTEGER },
            ...CARD_PROPERTIES,
          },
          required: ['index', 'card_type', 'name', 'code', 'delim_lga', 'delim_ward', 'vin', 'rotation_needed', 'legible'],
        },
      });

      const parsed = JSON.parse(text || '[]');
      return res.json({ ok: true, model, data: Array.isArray(parsed) ? parsed : [] });
    } catch (err: unknown) {
      if (isRateLimitError(err)) {
        const retryAfterSec = extractRetrySeconds(err);
        return res.status(429).json({
          ok: false,
          code: 'QUOTA_EXHAUSTED',
          retryAfterSec,
          error: 'Gemini free-tier quota limit reached; switching to local OCR.',
        });
      }
      const message = err instanceof Error ? err.message : 'Batch extraction failed';
      console.error('/api/extract-cards-batch error:', message);
      return res.status(500).json({ ok: false, error: message });
    }
  });

  function isValidExternalSyncUrl(rawUrl: unknown): string | null {
    if (typeof rawUrl !== 'string') return null;
    const trimmed = rawUrl.trim();
    if (!trimmed) return null;
    try {
      const parsed = new URL(trimmed);
      if (parsed.protocol !== 'https:') return null;
      const host = parsed.hostname.toLowerCase();
      if (host === 'localhost' || host === '127.0.0.1' || host === '0.0.0.0' || host.startsWith('10.') || host.startsWith('192.168.') || host.startsWith('169.254.')) {
        return null;
      }
      return parsed.toString();
    } catch {
      return null;
    }
  }

  app.post('/api/sync-proxy', async (req, res) => {
    const syncUrl = isValidExternalSyncUrl(req.body?.syncUrl);
    const payload = req.body?.payload;
    if (!syncUrl) {
      return res.status(400).json({ ok: false, error: 'Invalid Sync URL. Must be a valid https:// Google Apps Script Web App URL.' });
    }
    if (!payload || typeof payload !== 'object') {
      return res.status(400).json({ ok: false, error: 'payload object is required' });
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000);
    try {
      const upstream = await fetch(syncUrl, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      const finalUrl = upstream.url || '';
      const txt = await upstream.text();

      if (finalUrl.includes('accounts.google.com') || /ServiceLogin|accounts\.google\.com/i.test(txt)) {
        return res.status(403).json({
          ok: false,
          error: 'Google Apps Script requires login. Redeploy with "Who has access" set to "Anyone".',
        });
      }

      if (!upstream.ok) {
        return res.status(502).json({
          ok: false,
          status: upstream.status,
          error: `Sync endpoint returned HTTP ${upstream.status}`,
        });
      }

      try {
        const parsed = JSON.parse(txt);
        return res.json({ ok: true, data: parsed });
      } catch {
        return res.json({ ok: true, data: { result: 'success', raw: txt } });
      }
    } catch (err: unknown) {
      clearTimeout(timeoutId);
      const message = err instanceof Error ? err.message : 'Upstream sync request failed';
      console.error('/api/sync-proxy error:', message);
      return res.status(502).json({ ok: false, error: message });
    }
  });

  app.post('/api/test-sync-url', async (req, res) => {
    const syncUrl = isValidExternalSyncUrl(req.body?.syncUrl);
    if (!syncUrl) {
      return res.status(400).json({ ok: false, error: 'Invalid Sync URL. Must start with https://' });
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 15000);
    try {
      const upstream = await fetch(syncUrl, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      const finalUrl = upstream.url || '';
      const txt = await upstream.text();
      if (finalUrl.includes('accounts.google.com') || /ServiceLogin|accounts\.google\.com/i.test(txt)) {
        return res.status(403).json({
          ok: false,
          status: upstream.status,
          error: 'Apps Script redirected to Google Sign-In. Set "Who has access" to "Anyone" and redeploy.',
        });
      }

      return res.json({
        ok: upstream.ok,
        status: upstream.status,
        error: upstream.ok ? undefined : `HTTP ${upstream.status}`,
      });
    } catch (err: unknown) {
      clearTimeout(timeoutId);
      const message = err instanceof Error ? err.message : 'Connection test failed';
      return res.status(502).json({ ok: false, error: message });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(process.cwd(), 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Mobilize server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
