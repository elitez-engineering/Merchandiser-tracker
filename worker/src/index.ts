/**
 * PMI SKU OCR — reads a merchandiser's dispenser/shelf photo and returns the
 * PMI/BAT/JTI/other cigarette SKUs + prices it can see, matched against the
 * canonical SKU list from data/pmi-sku-reference.json (sent by the caller).
 *
 * Modeled on elitezshelf-frontage's categorise-photo Supabase function
 * (Claude vision -> structured JSON), adapted to a standalone Cloudflare
 * Worker since Merchandiser-tracker (Firebase + GitHub Pages) has no backend
 * of its own.
 *
 * Every detection here is a SUGGESTION — pmi-intake.js always requires the
 * admin to confirm/edit before it is saved, so a bad OCR read never reaches
 * the exported master Excel unreviewed.
 */

export interface Env {
  ANTHROPIC_API_KEY: string;
}

const MODEL = "claude-haiku-4-5-20251001";
const ALLOWED_ORIGIN = "https://elitez-engineering.github.io";
const MAX_BASE64_LEN = 10_000_000; // ~7.3MB decoded — a phone photo comfortably fits

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Max-Age": "86400",
};

type SkuRef = { excelCol: number; name: string };
type Detection = { excelCol: number; name: string; price: string | null; confidence: number };

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "content-type": "application/json" },
  });
}

async function detectSkus(
  base64: string,
  mediaType: string,
  skuNames: SkuRef[],
  apiKey: string,
): Promise<Detection[]> {
  const namesList = skuNames.map((s) => s.name).join("\n- ");
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 1024,
      messages: [
        {
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
            {
              type: "text",
              text:
                `This is a photo of a cigarette retail display/dispenser taken by a merchandiser doing a PMI distribution audit. ` +
                `Identify every visible SKU and its shelf-edge price label, if legible.\n\n` +
                `Respond with ONLY a JSON array on a single line, no markdown, no commentary:\n` +
                `[{"name":"<exact SKU name>","price":"<price as a plain number string, or null if not legible>","confidence":<0.0-1.0>}]\n\n` +
                `The "name" field MUST be copied EXACTLY (character for character) from this canonical list — do not invent, translate, ` +
                `abbreviate, or paraphrase a name. If a visible pack does not clearly match one of these exact names, omit it rather than guessing:\n` +
                `- ${namesList}\n\n` +
                `Only include SKUs you can actually see a pack of. Return [] if nothing in the list is visible.`,
            },
          ],
        },
      ],
    }),
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`anthropic ${resp.status}: ${body.slice(0, 300)}`);
  }
  const data = (await resp.json()) as { content?: { text?: string }[] };
  const text = data?.content?.[0]?.text ?? "[]";
  let parsed: { name: string; price: string | null; confidence: number }[];
  try {
    parsed = JSON.parse(text);
  } catch {
    const m = text.match(/\[[\s\S]*\]/);
    if (!m) throw new Error(`unparseable response: ${text.slice(0, 300)}`);
    parsed = JSON.parse(m[0]);
  }
  if (!Array.isArray(parsed)) throw new Error(`expected a JSON array, got: ${text.slice(0, 300)}`);

  const byName = new Map(skuNames.map((s) => [s.name, s.excelCol]));
  const out: Detection[] = [];
  for (const d of parsed) {
    if (typeof d?.name !== "string") continue;
    const excelCol = byName.get(d.name);
    if (excelCol === undefined) continue; // model invented/paraphrased a name — drop it, don't guess
    out.push({
      excelCol,
      name: d.name,
      price: typeof d.price === "string" ? d.price.slice(0, 20) : null,
      confidence: typeof d.confidence === "number" ? Math.max(0, Math.min(1, d.confidence)) : 0,
    });
  }
  return out;
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
    if (request.method !== "POST") return jsonResponse({ error: "POST only" }, 405);

    let payload: { image?: string; mediaType?: string; skuNames?: SkuRef[] };
    try {
      payload = await request.json();
    } catch {
      return jsonResponse({ error: "invalid json" }, 400);
    }

    const { image, mediaType, skuNames } = payload;
    if (typeof image !== "string" || !image) return jsonResponse({ error: "missing image" }, 400);
    if (image.length > MAX_BASE64_LEN) return jsonResponse({ error: "image too large" }, 413);
    if (!Array.isArray(skuNames) || !skuNames.length) return jsonResponse({ error: "missing skuNames" }, 400);
    const safeMediaType = typeof mediaType === "string" && /^image\/(jpeg|png|webp)$/.test(mediaType)
      ? mediaType
      : "image/jpeg";

    if (!env.ANTHROPIC_API_KEY) return jsonResponse({ error: "ANTHROPIC_API_KEY not configured" }, 500);

    try {
      const detections = await detectSkus(image, safeMediaType, skuNames, env.ANTHROPIC_API_KEY);
      return jsonResponse({ ok: true, detections });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return jsonResponse({ error: msg }, 500);
    }
  },
};
