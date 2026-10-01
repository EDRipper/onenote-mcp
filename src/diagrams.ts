import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { Resvg } from "@resvg/resvg-js";
import { chromium, type Browser } from "playwright-core";

export type DiagramKind = "svg" | "mermaid";

export interface RenderedImage {
  png: Buffer;
  /** Display width in CSS pixels (the PNG itself is rendered at 2x for sharpness). */
  width: number;
  alt: string;
}

export class DiagramError extends Error {}

const SCALE = 2;
const MAX_DISPLAY_WIDTH = 900;
const MAX_SOURCE_CHARS = 200_000;
const RENDER_TIMEOUT_MS = 20_000;
const MAX_CONCURRENT_BROWSER_RENDERS = 2;

// ---------------- SVG (no browser: resvg) ----------------

export function renderSvg(svg: string, alt = "diagram"): RenderedImage {
  if (svg.length > MAX_SOURCE_CHARS) throw new DiagramError("SVG is too large (max 200k characters).");
  if (!/<svg[\s>]/i.test(svg)) throw new DiagramError("That doesn't look like SVG: it must contain an <svg> element.");
  let resvg: Resvg;
  try {
    resvg = new Resvg(svg, {
      fitTo: { mode: "zoom", value: SCALE },
      background: "white",
      font: { loadSystemFonts: true, defaultFontFamily: "DejaVu Sans" },
      // resvg never runs scripts; external hrefs are not fetched.
    });
  } catch (err) {
    throw new DiagramError(`Invalid SVG: ${err instanceof Error ? err.message : err}`);
  }
  const image = resvg.render();
  const png = image.asPng();
  const width = Math.min(Math.round(image.width / SCALE), MAX_DISPLAY_WIDTH);
  return { png, width, alt };
}

// ---------------- Mermaid (sandboxed headless Chromium) ----------------

const require = createRequire(import.meta.url);
let mermaidJs: string | undefined;
const mermaidSource = () => (mermaidJs ??= readFileSync(require.resolve("mermaid/dist/mermaid.min.js"), "utf8"));

let browserPromise: Promise<Browser> | undefined;
function getBrowser(): Promise<Browser> {
  if (!browserPromise) {
    browserPromise = chromium
      .launch({
        executablePath: process.env.CHROMIUM_PATH || undefined,
        args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
      })
      .then((b) => {
        b.on("disconnected", () => (browserPromise = undefined));
        return b;
      })
      .catch((err) => {
        browserPromise = undefined;
        throw new DiagramError(`Mermaid rendering is unavailable on this server (couldn't start Chromium: ${err.message}). Use an SVG diagram instead.`);
      });
  }
  return browserPromise;
}

let active = 0;
const queue: (() => void)[] = [];
async function withSlot<T>(fn: () => Promise<T>): Promise<T> {
  if (active >= MAX_CONCURRENT_BROWSER_RENDERS) await new Promise<void>((r) => queue.push(r));
  active++;
  try {
    return await fn();
  } finally {
    active--;
    queue.shift()?.();
  }
}

export async function renderMermaid(source: string, alt = "diagram"): Promise<RenderedImage> {
  if (source.length > MAX_SOURCE_CHARS) throw new DiagramError("Mermaid diagram is too large.");
  const browser = await getBrowser();
  return withSlot(async () => {
    const context = await browser.newContext({ deviceScaleFactor: SCALE, viewport: { width: 1600, height: 1200 }, javaScriptEnabled: true });
    try {
      // No network at all: the page can only run the Mermaid bundle we inject.
      await context.route("**/*", (route) => (route.request().url() === "about:blank" ? route.continue() : route.abort()));
      const page = await context.newPage();
      page.setDefaultTimeout(RENDER_TIMEOUT_MS);
      await page.setContent(
        `<!doctype html><html><head><style>body{margin:0;background:#fff;font-family:"DejaVu Sans",sans-serif}#out{display:inline-block;padding:16px}</style></head><body><div id="out"></div></body></html>`,
      );
      await page.addScriptTag({ content: mermaidSource() });
      const result = await page.evaluate(async (src) => {
        const m = (globalThis as any).mermaid;
        m.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default", fontFamily: "DejaVu Sans, sans-serif" });
        try {
          const { svg } = await m.render("d", src);
          const out = document.getElementById("out")!;
          out.innerHTML = svg;
          const el = out.querySelector("svg")!;
          // Mermaid emits width="100%" + max-width; pin the SVG to its natural (viewBox) size.
          const vb = el.viewBox.baseVal;
          el.style.maxWidth = "none";
          if (vb && vb.width) {
            el.setAttribute("width", String(vb.width));
            el.setAttribute("height", String(vb.height));
          }
          const r = out.getBoundingClientRect();
          return { ok: true as const, width: Math.ceil(r.width), height: Math.ceil(r.height) };
        } catch (e: any) {
          return { ok: false as const, error: String(e?.message ?? e) };
        }
      }, source);
      if (!result.ok) throw new DiagramError(`Mermaid syntax error: ${result.error}`);
      const png = await page.locator("#out").screenshot({ type: "png", omitBackground: false });
      return { png, width: Math.min(result.width, MAX_DISPLAY_WIDTH), alt };
    } finally {
      await context.close().catch(() => {});
    }
  });
}

export function renderDiagram(kind: DiagramKind, source: string, alt?: string): Promise<RenderedImage> {
  return kind === "svg" ? Promise.resolve().then(() => renderSvg(source, alt)) : renderMermaid(source, alt);
}

export async function closeBrowser() {
  const b = await browserPromise?.catch(() => undefined);
  await b?.close();
}
