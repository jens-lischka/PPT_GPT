// supabase/functions/generate/index.ts
//
// Edge function: turn a natural-language brief (or an edit command on an
// existing slide) into a rendered OW deck via Claude + the Python renderer.
//
// Env (supabase secrets set):
//   ANTHROPIC_API_KEY            Claude API key
//   RENDERER_URL                 e.g. https://ow-renderer-xxxx.a.run.app
//   ALLOWED_ORIGINS              comma-separated; default https://localhost:3000
//   SUPABASE_URL                 (optional) enables persistence
//   SUPABASE_SERVICE_ROLE_KEY    (optional) enables persistence (bypasses RLS)
//
// Modes (POST body.mode):
//   "storyline"      { prompt, language?, audience? }
//        -> Claude drafts an outline; NO render. Pane shows it for confirmation.
//   "generate"       { prompt?, storyline?, language?, audience?, owner? }
//        -> Claude drafts a full deck_spec -> /render -> (persist) -> return.
//   "generate_slide" { prompt, language?, audience?, deck_id?, position? }
//        -> Claude drafts ONE slide -> /render-slide -> (persist) -> return.
//   "edit_slide"     { slide_spec, command, language?, audience?, deck_id?, slide_uuid? }
//        -> Claude edits the slide -> /render-slide -> (persist revision) -> return.
//
// The gate result travels back untouched; the pane blocks insertion when
// gate.passed is false (non-overridable invariant, carried over from the GPT).

import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import {
  SYSTEM_STORYLINE,
  SYSTEM_GENERATE,
  SYSTEM_GENERATE_SLIDE,
  SYSTEM_EDIT_SLIDE,
} from "../_shared/system_prompt.ts";

const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY")!;
const RENDERER_URL = Deno.env.get("RENDERER_URL")!;
const ALLOWED_ORIGINS = (Deno.env.get("ALLOWED_ORIGINS") ?? "https://localhost:3000")
  .split(",").map((s) => s.trim());
const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const PERSIST = Boolean(SUPABASE_URL && SERVICE_ROLE);

const MODEL = "claude-opus-4-8"; // per Anthropic guidance: default to Opus 4.8

// ---- CORS -----------------------------------------------------------------
function corsHeaders(origin: string | null): HeadersInit {
  const allow = origin && ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Headers": "authorization, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Vary": "Origin",
  };
}

// ---- Claude ---------------------------------------------------------------
async function claude(system: string, userText: string): Promise<unknown> {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 16000,
      thinking: { type: "adaptive" },          // Opus 4.8: opt in explicitly
      output_config: { effort: "medium" },     // balance quality vs latency
      system: [{ type: "text", text: system, cache_control: { type: "ephemeral" } }],
      messages: [{ role: "user", content: userText }],
    }),
  });
  if (!res.ok) throw new Error(`Claude ${res.status}: ${await res.text()}`);
  const data = await res.json();
  const text = (data.content ?? [])
    .filter((b: { type: string }) => b.type === "text")
    .map((b: { text: string }) => b.text)
    .join("\n");
  return parseJson(text);
}

// Robust extraction: strip fences, then take the first balanced {...} or [...].
function parseJson(raw: string): unknown {
  let t = raw.trim().replace(/^```(?:json)?/i, "").replace(/```$/,"").trim();
  try { return JSON.parse(t); } catch { /* fall through */ }
  const start = t.search(/[{[]/);
  if (start < 0) throw new Error(`no JSON in model output: ${t.slice(0, 200)}`);
  const open = t[start], close = open === "{" ? "}" : "]";
  let depth = 0, inStr = false, esc = false;
  for (let i = start; i < t.length; i++) {
    const c = t[i];
    if (inStr) { if (esc) esc = false; else if (c === "\\") esc = true; else if (c === '"') inStr = false; continue; }
    if (c === '"') inStr = true;
    else if (c === open) depth++;
    else if (c === close && --depth === 0) return JSON.parse(t.slice(start, i + 1));
  }
  throw new Error(`unbalanced JSON in model output: ${t.slice(0, 200)}`);
}

// ---- Renderer -------------------------------------------------------------
async function renderDeck(deck_spec: unknown) {
  const res = await fetch(`${RENDERER_URL}/render`, {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ deck_spec }),
  });
  if (!res.ok) throw new Error(`renderer /render ${res.status}: ${await res.text()}`);
  return res.json();
}
async function renderSlide(slide_spec: unknown, language?: string, audience?: string) {
  const res = await fetch(`${RENDERER_URL}/render-slide`, {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ slide_spec, language, audience }),
  });
  if (!res.ok) throw new Error(`renderer /render-slide ${res.status}: ${await res.text()}`);
  return res.json();
}

// ---- Supabase persistence (optional; service role, bypasses RLS) ----------
async function sb(path: string, init: RequestInit) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      "apikey": SERVICE_ROLE!, "authorization": `Bearer ${SERVICE_ROLE}`,
      "content-type": "application/json", "prefer": "return=representation",
      ...(init.headers ?? {}),
    },
  });
  if (!res.ok) throw new Error(`supabase ${path} ${res.status}: ${await res.text()}`);
  return res.json();
}

interface DeckSpec { language?: string; audience?: string; title?: string; slides: any[]; }

async function persistDeck(spec: DeckSpec, owner: string) {
  const [deck] = await sb("decks", {
    method: "POST",
    body: JSON.stringify({ owner, title: spec.title ?? "Untitled deck",
      language: spec.language ?? null, audience: spec.audience ?? null }),
  });
  const slideRows = spec.slides.map((s, i) => ({ deck_id: deck.id, position: i, spec: s }));
  const slides = await sb("slides", { method: "POST", body: JSON.stringify(slideRows) });
  // revisions: one per slide, source 'generate'
  await sb("revisions", { method: "POST", body: JSON.stringify(
    slides.map((row: any) => ({ slide_id: row.id, spec: row.spec, source: "generate" }))) });
  return { deck_id: deck.id,
    slides: slides.map((r: any) => ({ position: r.position, slide_uuid: r.id })) };
}

async function persistEdit(deck_id: string, slide_uuid: string, spec: unknown,
                           gate: unknown, command: string) {
  await sb(`slides?id=eq.${slide_uuid}`, { method: "PATCH",
    body: JSON.stringify({ spec, gate_result: gate }) });
  await sb("revisions", { method: "POST",
    body: JSON.stringify({ slide_id: slide_uuid, spec, source: "chat_edit", prompt: command }) });
}

// ---- HTTP -----------------------------------------------------------------
Deno.serve(async (req) => {
  const origin = req.headers.get("origin");
  const cors = corsHeaders(origin);
  if (req.method === "OPTIONS") return new Response(null, { headers: cors });

  const json = (body: unknown, status = 200) =>
    new Response(JSON.stringify(body), { status, headers: { ...cors, "content-type": "application/json" } });

  try {
    const body = await req.json();
    const { mode } = body;

    if (mode === "storyline") {
      const storyline = await claude(SYSTEM_STORYLINE,
        `Brief: ${body.prompt}\n` +
        (body.language ? `language: ${body.language}\n` : "") +
        (body.audience ? `audience: ${body.audience}\n` : ""));
      return json({ storyline });
    }

    if (mode === "generate") {
      const spec = await claude(SYSTEM_GENERATE,
        (body.storyline
          ? `Confirmed storyline (build the full deck from it):\n${JSON.stringify(body.storyline)}\n`
          : `Brief: ${body.prompt}\n`) +
        (body.language ? `language: ${body.language}\n` : "") +
        (body.audience ? `audience: ${body.audience}\n` : "")) as DeckSpec;
      const rendered = await renderDeck(spec);
      let persisted = null;
      if (PERSIST && body.owner && rendered.gate_result?.passed) {
        persisted = await persistDeck(spec, body.owner);
      }
      return json({ spec, ...rendered, persisted });
    }

    if (mode === "generate_slide") {
      const slide = await claude(SYSTEM_GENERATE_SLIDE,
        `Request: ${body.prompt}\n` +
        (body.language ? `language: ${body.language}\n` : "") +
        (body.audience ? `audience: ${body.audience}\n` : ""));
      const rendered = await renderSlide(slide, body.language, body.audience);
      return json({ spec: slide, ...rendered });
    }

    if (mode === "edit_slide") {
      if (!body.slide_spec || !body.command)
        return json({ error: "edit_slide requires slide_spec and command" }, 400);
      const slide = await claude(SYSTEM_EDIT_SLIDE,
        `Existing slide spec:\n${JSON.stringify(body.slide_spec)}\n\n` +
        `Edit command: ${body.command}\n` +
        (body.language ? `language: ${body.language}\n` : "") +
        (body.audience ? `audience: ${body.audience}\n` : ""));
      const rendered = await renderSlide(slide, body.language, body.audience);
      if (PERSIST && body.deck_id && body.slide_uuid && rendered.gate_result?.passed) {
        await persistEdit(body.deck_id, body.slide_uuid, slide, rendered.gate_result, body.command);
      }
      return json({ spec: slide, ...rendered });
    }

    return json({ error: `unknown mode: ${mode}` }, 400);
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
});
