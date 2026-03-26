import type { EditorView } from '@milkdown/kit/prose/view';

// ── Config ─────────────────────────────────────────────────────────────────────

let agentModel = 'claude-sonnet-4-6';
let agentBy = 'proof:ai';
let docSlug = '';
let shareToken = '';

// ── @proof mention detection ───────────────────────────────────────────────────

const PROOF_MENTION_RE = /@proof\s+(.+)/i;
let lastSeenText = '';
let debounceTimer: ReturnType<typeof setTimeout> | null = null;
let processing = false;

function extractDocSlugFromUrl(): string {
  const match = window.location.pathname.match(/\/d\/([^/?#]+)/);
  return match ? match[1] : '';
}

function extractShareTokenFromUrl(): string {
  const params = new URLSearchParams(window.location.search);
  return params.get('token') || '';
}

async function handleProofMention(view: EditorView, mention: string): Promise<void> {
  if (processing) return;
  processing = true;

  const slug = docSlug || extractDocSlugFromUrl();
  const token = shareToken || extractShareTokenFromUrl();
  if (!slug) { processing = false; return; }

  // Find which block contains the @proof mention to get blockRef
  let blockRef: string | null = null;
  const doc = view.state.doc;
  let blockIndex = 0;
  doc.forEach((node) => {
    blockIndex++;
    if (node.textContent.match(PROOF_MENTION_RE)) {
      blockRef = `b${blockIndex}`;
    }
  });

  try {
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (token) headers['x-share-token'] = token;

    await fetch(`/api/agent/${slug}/proof-ask`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        mention,
        blockRef,
        model: agentModel,
        by: agentBy,
      }),
    });
  } catch (err) {
    console.warn('[proof] @proof request failed:', err);
  } finally {
    processing = false;
  }
}

function checkForProofMention(view: EditorView): void {
  const text = view.state.doc.textContent;
  if (text === lastSeenText) return;
  lastSeenText = text;

  const match = text.match(PROOF_MENTION_RE);
  if (!match) return;

  const mention = match[1].trim();
  if (!mention) return;

  if (debounceTimer) clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    void handleProofMention(view, mention);
  }, 800);
}

// ── Public API ─────────────────────────────────────────────────────────────────

export function initAgentIntegration(view: EditorView): void {
  docSlug = extractDocSlugFromUrl();
  shareToken = extractShareTokenFromUrl();

  // Poll for @proof mentions every 2 seconds
  const interval = setInterval(() => {
    checkForProofMention(view);
  }, 2000);

  // Clean up on view destroy
  const originalDestroy = view.destroy.bind(view);
  view.destroy = () => {
    clearInterval(interval);
    if (debounceTimer) clearTimeout(debounceTimer);
    originalDestroy();
  };
}

export function handleMarksChange(_marks: unknown[], _view: EditorView): void {
  // Reserved for future use
}

export function sweepForActionableItems(_triggerOnFirstSweep = false): void {
  // Reserved for always-on mode
}

export function setAlwaysOnEnabled(_enabled: boolean): void {
  // Reserved for always-on mode
}

export function configureAgent(options?: { model?: string; by?: string; slug?: string; token?: string }): void {
  if (options?.model) agentModel = options.model;
  if (options?.by) agentBy = options.by;
  if (options?.slug) docSlug = options.slug;
  if (options?.token) shareToken = options.token;
}

export function isAgentReady(): boolean {
  return Boolean(docSlug || extractDocSlugFromUrl());
}

export function cleanupAgentIntegration(): void {
  if (debounceTimer) clearTimeout(debounceTimer);
  lastSeenText = '';
  processing = false;
}
