// supabase/functions/generate/index.ts
// Edge function: turn a natural-language prompt (or an edit command on an
// existing slide spec) into a rendered deck via Claude + the Python renderer.
//
// Env (set via `supabase secrets set`):
//   ANTHROPIC_API_KEY   Claude API key
//   RENDERER_URL        e.g. https://ow-renderer-xxxx.a.run.app
//
// Modes:
//   { mode: "generate", prompt, language?, audience? }
//       -> Claude drafts a full deck_spec -> POST renderer /render
//   { mode: "edit_slide", slide_spec, command, language? }
//       -> Claude edits the fragment -> POST renderer /render-slide
//
// This is deliberately a skeleton: the system prompt below is a placeholder.
// The real one is distilled from the GPT package's 00_gpt_instructions.md +
// the intent schema in 01_readme_deck_system.md — port it in Phase 1 and keep
// it in sync with the runtime version pinned in CLAUDE.md.

import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY")!;
const RENDERER_URL = Deno.env.get("RENDERER_URL")!;

const SYSTEM_PROMPT = `You produce deck_spec JSON for the OW deck runtime.
Respond ONLY with valid JSON, no prose, no markdown fences.
(Placeholder — port the distilled instructions + intent schema here.)`;

async function claude(messages: { role: string; content: string }[]) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 8000,
      system: SYSTEM_PROMPT,
      messages,
    }),
  });
  if (!res.ok) throw new Error(`Claude ${res.status}: ${await res.text()}`);
  const data = await res.json();
  const text = data.content
    .filter((b: { type: string }) => b.type === "text")
    .map((b: { text: string }) => b.text)
    .join("\n");
  return JSON.parse(text.replace(/```json|```/g, "").trim());
}

async function renderDeck(deck_spec: unknown) {
  const res = await fetch(`${RENDERER_URL}/render`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ deck_spec }),
  });
  if (!res.ok) throw new Error(`renderer ${res.status}: ${await res.text()}`);
  return res.json();
}

async function renderSlide(slide_spec: unknown, language?: string) {
  const res = await fetch(`${RENDERER_URL}/render-slide`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ slide_spec, language }),
  });
  if (!res.ok) throw new Error(`renderer ${res.status}: ${await res.text()}`);
  return res.json();
}

Deno.serve(async (req) => {
  try {
    const body = await req.json();

    if (body.mode === "generate") {
      const spec = await claude([{
        role: "user",
        content:
          `Create a deck_spec for: ${body.prompt}\n` +
          (body.language ? `language: ${body.language}\n` : "") +
          (body.audience ? `audience: ${body.audience}\n` : ""),
      }]);
      const rendered = await renderDeck(spec);
      return Response.json({ spec, ...rendered });
    }

    if (body.mode === "edit_slide") {
      const spec = await claude([{
        role: "user",
        content:
          `Here is a slide spec fragment:\n${JSON.stringify(body.slide_spec)}\n` +
          `Apply this edit and return the full updated fragment:\n${body.command}`,
      }]);
      const rendered = await renderSlide(spec, body.language);
      return Response.json({ spec, ...rendered });
    }

    return Response.json({ error: "unknown mode" }, { status: 400 });
  } catch (e) {
    return Response.json({ error: String(e) }, { status: 500 });
  }
});
