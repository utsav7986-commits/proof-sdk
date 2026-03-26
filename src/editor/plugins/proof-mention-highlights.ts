import { $prose } from '@milkdown/kit/utils';
import { Plugin, PluginKey } from '@milkdown/kit/prose/state';
import { Decoration, DecorationSet, type EditorView } from '@milkdown/kit/prose/view';

export type ProofMentionStatus = 'idle' | 'processing' | 'done' | 'error';

type ProofMentionState = {
  status: ProofMentionStatus;
  decorations: DecorationSet;
};

type ProofMentionMeta = { type: 'set-status'; status: ProofMentionStatus };

const proofMentionKey = new PluginKey<ProofMentionState>('proof-mention-highlights');

const MENTION_RE = /@proof\b[^\n]*/gi;

const STATUS_STYLES: Record<ProofMentionStatus, string> = {
  idle:       'background: linear-gradient(90deg,#7c3aed22,#7c3aed11); color: #7c3aed; border-radius: 4px; padding: 1px 4px; font-weight: 600; border-bottom: 2px solid #7c3aed88; cursor: pointer;',
  processing: 'background: linear-gradient(90deg,#f59e0b22,#f59e0b11); color: #b45309; border-radius: 4px; padding: 1px 4px; font-weight: 600; border-bottom: 2px solid #f59e0b; animation: proof-pulse 1s infinite;',
  done:       'background: linear-gradient(90deg,#10b98122,#10b98111); color: #047857; border-radius: 4px; padding: 1px 4px; font-weight: 600; border-bottom: 2px solid #10b981;',
  error:      'background: linear-gradient(90deg,#ef444422,#ef444411); color: #b91c1c; border-radius: 4px; padding: 1px 4px; font-weight: 600; border-bottom: 2px solid #ef4444;',
};

const STATUS_CLASSES: Record<ProofMentionStatus, string> = {
  idle:       'proof-mention proof-mention--idle',
  processing: 'proof-mention proof-mention--processing',
  done:       'proof-mention proof-mention--done',
  error:      'proof-mention proof-mention--error',
};

// Inject CSS once
let cssInjected = false;
function injectMentionCSS(): void {
  if (cssInjected || typeof document === 'undefined') return;
  cssInjected = true;
  const style = document.createElement('style');
  style.textContent = `
    @keyframes proof-pulse {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.55; }
    }
    .proof-mention--processing::after {
      content: ' ⟳';
      animation: proof-pulse 1s infinite;
    }
    .proof-mention--done::after { content: ' ✓'; }
    .proof-mention--error::after { content: ' ✗'; }
  `;
  document.head.appendChild(style);
}

// Map a text offset within a block's concatenated text back to a document position.
// segments = [{text, from}] collected by walking inline children of the block.
function resolveTextOffset(segments: Array<{ text: string; from: number }>, offset: number): number {
  let accumulated = 0;
  for (const seg of segments) {
    if (offset <= accumulated + seg.text.length) {
      return seg.from + (offset - accumulated);
    }
    accumulated += seg.text.length;
  }
  // offset is at or past the end — clamp to end of last segment
  const last = segments[segments.length - 1];
  return last ? last.from + last.text.length : 0;
}

function buildDecorations(doc: Parameters<typeof DecorationSet.create>[0], status: ProofMentionStatus): DecorationSet {
  const decorations: Decoration[] = [];
  const title = status === 'idle' ? 'Click to send @proof request'
    : status === 'processing' ? 'AI is thinking…'
    : status === 'done' ? 'Response added below'
    : 'Request failed';

  doc.descendants((node, pos) => {
    // Only process block nodes that hold inline content (paragraphs, headings, etc.)
    // @proof text may be split across multiple text nodes by author-tracking spans,
    // so we match against the full concatenated block text rather than per-text-node.
    if (!node.isBlock || !node.inlineContent) return;

    // Collect text segments with their document start positions
    const segments: Array<{ text: string; from: number }> = [];
    node.forEach((child, childOffset) => {
      if (child.isText && child.text) {
        segments.push({ text: child.text, from: pos + 1 + childOffset });
      }
    });
    if (segments.length === 0) return false;

    const fullText = segments.map(s => s.text).join('');
    MENTION_RE.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = MENTION_RE.exec(fullText)) !== null) {
      const from = resolveTextOffset(segments, match.index);
      const to = resolveTextOffset(segments, match.index + match[0].length);
      decorations.push(
        Decoration.inline(from, to, {
          class: STATUS_CLASSES[status],
          style: STATUS_STYLES[status],
          title,
        })
      );
    }
    return false; // children already handled above
  });
  return DecorationSet.create(doc, decorations);
}

export function setProofMentionStatus(view: EditorView, status: ProofMentionStatus): void {
  injectMentionCSS();
  const tr = view.state.tr.setMeta(proofMentionKey, { type: 'set-status', status } satisfies ProofMentionMeta);
  view.dispatch(tr);
}

export const proofMentionHighlightsPlugin = $prose(() => {
  injectMentionCSS();
  return new Plugin<ProofMentionState>({
    key: proofMentionKey,
    state: {
      init(_, state) {
        return { status: 'idle', decorations: buildDecorations(state.doc, 'idle') };
      },
      apply(tr, pluginState, _old, newState) {
        const meta = tr.getMeta(proofMentionKey) as ProofMentionMeta | undefined;
        const status = meta?.type === 'set-status' ? meta.status : pluginState.status;
        if (tr.docChanged || meta) {
          return { status, decorations: buildDecorations(newState.doc, status) };
        }
        return pluginState;
      },
    },
    props: {
      decorations(state) {
        return proofMentionKey.getState(state)?.decorations ?? DecorationSet.empty;
      },
    },
  });
});
