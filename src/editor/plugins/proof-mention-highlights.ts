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

function buildDecorations(doc: Parameters<typeof DecorationSet.create>[0], status: ProofMentionStatus): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    MENTION_RE.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = MENTION_RE.exec(node.text)) !== null) {
      const from = pos + match.index;
      const to = from + match[0].length;
      decorations.push(
        Decoration.inline(from, to, {
          class: STATUS_CLASSES[status],
          style: STATUS_STYLES[status],
          title: status === 'idle' ? 'Click to send @proof request' : status === 'processing' ? 'AI is thinking…' : status === 'done' ? 'Response added below' : 'Request failed',
        })
      );
    }
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
