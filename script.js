function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttribute(text) {
  return escapeHtml(text).replace(/"/g, "&quot;");
}

const inlineMetricsCache = new Map();
let metricsScriptQueue = Promise.resolve();
let latestRenderId = 0;

function buildCodecogsLatex(latex, displayMode) {
  const trimmed = latex.trim();
  return displayMode ? trimmed : `\\inline ${trimmed}`;
}

function buildMathUrl(latex, displayMode, format = "svg") {
  const encoded = encodeURIComponent(buildCodecogsLatex(latex, displayMode));
  return `https://latex.codecogs.com/${format}.image?${encoded}`;
}

function buildInlineMetricsUrl(latex, responseFormat = "json") {
  const encoded = encodeURIComponent(buildCodecogsLatex(latex, false));
  return `https://latex.codecogs.com/gif.${responseFormat}?${encoded}`;
}

function normalizeInlineMetrics(payload) {
  const data = payload?.latex ?? payload;

  if (!data) {
    return null;
  }

  const width = Number.parseFloat(data.width);
  const height = Number.parseFloat(data.height);
  const baseline = Number.parseFloat(data.baseline);

  return {
    width: Number.isFinite(width) ? width : null,
    height: Number.isFinite(height) ? height : null,
    baseline: Number.isFinite(baseline) ? baseline : null,
  };
}

async function fetchInlineMetricsViaJson(latex) {
  const response = await fetch(buildInlineMetricsUrl(latex), {
    mode: "cors",
  });

  if (!response.ok) {
    throw new Error(`CodeCogs metadata request failed with ${response.status}.`);
  }

  return normalizeInlineMetrics(await response.json());
}

function fetchInlineMetricsViaScript(latex) {
  const request = () =>
    new Promise((resolve) => {
      const previousParseEqn = window.ParseEqn;
      const script = document.createElement("script");
      let timeoutId = 0;
      let settled = false;

      const cleanup = () => {
        window.clearTimeout(timeoutId);
        script.remove();

        if (previousParseEqn === undefined) {
          delete window.ParseEqn;
        } else {
          window.ParseEqn = previousParseEqn;
        }
      };

      const finish = (payload) => {
        if (settled) {
          return;
        }

        settled = true;
        cleanup();
        resolve(normalizeInlineMetrics(payload));
      };

      window.ParseEqn = finish;
      script.async = true;
      script.src = buildInlineMetricsUrl(latex, "javascript");
      script.onerror = () => finish(null);
      timeoutId = window.setTimeout(() => finish(null), 4000);
      document.head.appendChild(script);
    });

  metricsScriptQueue = metricsScriptQueue.then(request, request);
  return metricsScriptQueue;
}

async function getInlineMetrics(latex) {
  const trimmed = latex.trim();

  if (!trimmed) {
    return null;
  }

  if (!inlineMetricsCache.has(trimmed)) {
    inlineMetricsCache.set(
      trimmed,
      (async () => {
        try {
          return await fetchInlineMetricsViaJson(trimmed);
        } catch (error) {
          console.warn("Falling back to script-based CodeCogs metadata.", error);
          return fetchInlineMetricsViaScript(trimmed);
        }
      })()
    );
  }

  return inlineMetricsCache.get(trimmed);
}

function inlineMathStyle(metrics) {
  const style = [];

  if (metrics?.width) {
    style.push(`width:${metrics.width}px`);
  }

  if (metrics?.height) {
    style.push(`height:${metrics.height}px`);
  }

  if (Number.isFinite(metrics?.baseline)) {
    style.push(`vertical-align:-${metrics.baseline}px`);
  } else {
    style.push("vertical-align:-0.2ex");
  }

  return `${style.join(";")};`;
}

function latexToImg(latex, displayMode, format = "svg", metrics = null) {
  const trimmed = latex.trim();
  const url = buildMathUrl(trimmed, displayMode, format);
  const cls = displayMode ? "math-display" : "math-inline";
  const alt = escapeAttribute(trimmed.replace(/\s+/g, " "));
  const title = escapeAttribute(trimmed);
  const style = displayMode
    ? "display:block;margin:12px auto;max-width:100%;"
    : inlineMathStyle(metrics);

  return `<img class="${cls}" src="${url}" alt="${alt}" title="${title}" style="${style}">`;
}

function encodeMathToken(latex) {
  return btoa(unescape(encodeURIComponent(latex)));
}

function decodeMathToken(encoded) {
  return decodeURIComponent(escape(atob(encoded)));
}

async function replaceInlineMathTokens(html, format) {
  const tokenRegex = /@@INLINE_MATH_([^@]+)@@/g;
  const replacements = new Map();
  const matches = [...html.matchAll(tokenRegex)];

  await Promise.all(
    matches.map(async ([token, encoded]) => {
      if (replacements.has(token)) {
        return;
      }

      const latex = decodeMathToken(encoded);
      const metrics = await getInlineMetrics(latex);
      replacements.set(token, latexToImg(latex, false, format, metrics));
    })
  );

  for (const [token, replacement] of replacements) {
    html = html.split(token).join(replacement);
  }

  return html;
}

function replaceDisplayMathTokens(html, format) {
  return html.replace(/@@DISPLAY_MATH_([^@]+)@@/g, (_, encoded) => {
    const latex = decodeMathToken(encoded);
    return `<div>${latexToImg(latex, true, format)}</div>`;
  });
}

async function renderInputToHtml(text, format = "svg") {
  let html = escapeHtml(text);

  html = html.replace(
    /\$\$([\s\S]+?)\$\$/g,
    (_, latex) => `@@DISPLAY_MATH_${encodeMathToken(latex)}@@`
  );

  html = html.replace(
    /\$([^\$\n]+?)\$/g,
    (_, latex) => `@@INLINE_MATH_${encodeMathToken(latex)}@@`
  );

  html = html.replace(/\n/g, "<br>");
  html = replaceDisplayMathTokens(html, format);
  html = await replaceInlineMathTokens(html, format);

  return html;
}

async function copyHtmlToClipboard(html, plainText) {
  if (!navigator.clipboard || !window.ClipboardItem) {
    throw new Error("HTML clipboard API not supported in this browser.");
  }

  const item = new ClipboardItem({
    "text/html": new Blob([html], { type: "text/html" }),
    "text/plain": new Blob([plainText], { type: "text/plain" }),
  });

  await navigator.clipboard.write([item]);
}

const input = document.getElementById("input");
const preview = document.getElementById("preview");
const renderBtn = document.getElementById("renderBtn");
const copyBtn = document.getElementById("copyBtn");
const copyPngBtn = document.getElementById("copyPngBtn");

async function refreshPreview(format = "svg") {
  const renderId = ++latestRenderId;
  const html = await renderInputToHtml(input.value, format);

  if (renderId === latestRenderId) {
    preview.innerHTML = html;
  }

  return html;
}

renderBtn.addEventListener("click", async () => {
  try {
    await refreshPreview();
  } catch (err) {
    console.error(err);
    alert("Render failed: " + err.message);
  }
});

copyBtn.addEventListener("click", async () => {
  try {
    const html = await refreshPreview();
    await copyHtmlToClipboard(html, input.value);
    alert("Rendered HTML copied. Now paste into Outlook or another mail composer.");
  } catch (err) {
    console.error(err);
    alert("Copy failed: " + err.message);
  }
});

copyPngBtn.addEventListener("click", async () => {
  try {
    const html = await refreshPreview("png");
    await copyHtmlToClipboard(html, input.value);
    alert("Rendered HTML copied using PNG math images.");
  } catch (err) {
    console.error(err);
    alert("Copy failed: " + err.message);
  }
});

refreshPreview().catch((err) => {
  console.error(err);
});
