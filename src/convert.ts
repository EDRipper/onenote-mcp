import TurndownService from "turndown";
import { Marked, type Token } from "marked";
import { escapeHtml } from "./pages.js";

// ---------- OneNote HTML -> Markdown (for reading) ----------

const turndown = new TurndownService({ headingStyle: "atx", codeBlockStyle: "fenced", bulletListMarker: "-" });

// OneNote to-do tags: <p data-tag="to-do">, <p data-tag="to-do:completed">
turndown.addRule("onenoteTodo", {
  filter: (node) => /(^|,)\s*to-do/.test(node.getAttribute?.("data-tag") ?? ""),
  replacement: (content, node) => {
    const done = /to-do:completed/.test((node as HTMLElement).getAttribute("data-tag") ?? "");
    return `\n- [${done ? "x" : " "}] ${content.trim()}\n`;
  },
});

// Images live behind Graph resource URLs; expose their IDs so the model can fetch them.
turndown.addRule("onenoteImage", {
  filter: "img",
  replacement: (_c, node) => {
    const el = node as HTMLElement;
    const src = el.getAttribute("data-fullres-src") || el.getAttribute("src") || "";
    const id = resourceIdFromUrl(src);
    const alt = el.getAttribute("alt") || "image";
    return id ? `![${alt}](onenote-resource:${id})` : `![${alt}](${src})`;
  },
});

// Attached files (PDFs etc.)
turndown.addRule("onenoteObject", {
  filter: "object",
  replacement: (_c, node) => {
    const el = node as HTMLElement;
    const name = el.getAttribute("data-attachment") || "attachment";
    const id = resourceIdFromUrl(el.getAttribute("data") || "");
    return `\n[📎 ${name}]${id ? `(onenote-resource:${id})` : ""}\n`;
  },
});

turndown.remove(["script", "style", "title", "meta"]);

export function resourceIdFromUrl(url: string): string | undefined {
  return url.match(/\/resources\/([^/?]+)\/(?:\$value|content)/)?.[1];
}

export function htmlToMarkdown(html: string): string {
  const body = html.match(/<body[^>]*>([\s\S]*)<\/body>/i)?.[1] ?? html;
  return turndown.turndown(body).replace(/\n{3,}/g, "\n\n").trim();
}

export function htmlTitle(html: string): string | undefined {
  return html.match(/<title>([\s\S]*?)<\/title>/i)?.[1]?.trim();
}

// ---------- Markdown -> OneNote HTML (for writing) ----------

/**
 * Convert Markdown to the HTML subset OneNote accepts.
 * GitHub task lists become native OneNote to-do checkboxes.
 */
const md = new Marked({ gfm: true, breaks: false });
md.use({
  renderer: {
    // Render GitHub task-list items as native OneNote to-do checkboxes; other items stay in a normal list.
    list(token) {
      if (!token.items.some((i) => i.task)) return false;
      const tag = token.ordered ? "ol" : "ul";
      let out = "";
      let run: string[] = [];
      const flush = () => {
        if (run.length) out += `<${tag}>${run.join("")}</${tag}>`;
        run = [];
      };
      for (const item of token.items) {
        if (!item.task) {
          run.push(`<li>${this.parser.parse(item.tokens)}</li>`);
          continue;
        }
        flush();
        const [first, ...rest] = item.tokens as Token[];
        const inline = ((first as any)?.tokens ?? []).filter((t: Token) => t.type !== "checkbox");
        out += `<p data-tag="${item.checked ? "to-do:completed" : "to-do"}">${this.parser.parseInline(inline).trim()}</p>`;
        if (rest.length) out += this.parser.parse(rest);
      }
      flush();
      return out;
    },
  },
});

export function markdownToOneNoteHtml(markdown: string): string {
  let html = md.parse(markdown, { async: false }) as string;
  // OneNote ignores <pre> styling without a font; give code blocks a monospace face.
  html = html.replace(/<pre><code[^>]*>/g, '<pre style="font-family:Consolas,monospace"><code>');
  // Make tables visible.
  html = html.replace(/<table>/g, '<table border="1">');
  return html;
}

export function contentToHtml(content: string, format: "markdown" | "html"): string {
  return format === "html" ? content : markdownToOneNoteHtml(content);
}

export function buildPageHtml(title: string, bodyHtml: string): string {
  return `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title><meta name="created" content="${new Date().toISOString()}" /></head><body>${bodyHtml}</body></html>`;
}
