function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function latexToImg(latex, displayMode, format = "svg") {
  const trimmed = latex.trim();
  const encoded = encodeURIComponent(trimmed);
  const extension = format === "png" ? "png" : "svg";
  const url = `https://latex.codecogs.com/${extension}.latex?${encoded}`;
  const cls = displayMode ? "math-display" : "math-inline";
  const alt = trimmed.replace(/"/g, "&quot;");
  const style = displayMode
    ? "display:block;margin:12px auto;"
    : "vertical-align:middle;";

  return `<img class="${cls}" src="${url}" alt="${alt}" style="${style}">`;
}

function renderInputToHtml(text, format = "svg") {
  let html = escapeHtml(text);

  html = html.replace(
    /\$\$([\s\S]+?)\$\$/g,
    (_, latex) => `@@DISPLAY_MATH_${btoa(unescape(encodeURIComponent(latex)))}@@`
  );

  html = html.replace(
    /\$([^\$\n]+?)\$/g,
    (_, latex) => `@@INLINE_MATH_${btoa(unescape(encodeURIComponent(latex)))}@@`
  );

  html = html.replace(/\n/g, "<br>");

  html = html.replace(/@@DISPLAY_MATH_([^@]+)@@/g, (_, encoded) => {
    const latex = decodeURIComponent(escape(atob(encoded)));
    return `<div>${latexToImg(latex, true, format)}</div>`;
  });

  html = html.replace(/@@INLINE_MATH_([^@]+)@@/g, (_, encoded) => {
    const latex = decodeURIComponent(escape(atob(encoded)));
    return latexToImg(latex, false, format);
  });

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

function refreshPreview(format = "svg") {
  const html = renderInputToHtml(input.value, format);
  preview.innerHTML = html;
  return html;
}

renderBtn.addEventListener("click", refreshPreview);

copyBtn.addEventListener("click", async () => {
  try {
    const html = refreshPreview();
    await copyHtmlToClipboard(html, input.value);
    alert("Rendered HTML copied. Now paste into Outlook or another mail composer.");
  } catch (err) {
    console.error(err);
    alert("Copy failed: " + err.message);
  }
});

copyPngBtn.addEventListener("click", async () => {
  try {
    const html = refreshPreview("png");
    await copyHtmlToClipboard(html, input.value);
    alert("Rendered HTML copied using PNG math images.");
  } catch (err) {
    console.error(err);
    alert("Copy failed: " + err.message);
  }
});

refreshPreview();
