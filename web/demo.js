import init, { detect, to_plain, to_markdown } from './pkg/batdoc_core.js';

const fileInput = document.getElementById('file');
const modeSel = document.getElementById('mode');
const imagesCb = document.getElementById('images');
const ocrCb = document.getElementById('ocr');
const output = document.getElementById('output');
const status = document.getElementById('status');

function run(bytes) {
  const label = `detected ${detect(bytes)}`;
  let out;
  if (modeSel.value === 'plain') {
    out = to_plain(bytes);
  } else {
    out = to_markdown(bytes, imagesCb.checked, ocrCb.checked);
  }
  output.textContent = out;
  status.textContent = `${label} · ${bytes.length} bytes`;
}

async function load(path) {
  try {
    const res = await fetch(path);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    run(new Uint8Array(await res.arrayBuffer()));
  } catch (e) {
    output.textContent = '';
    status.textContent = `error: ${e}`;
  }
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (!file) return;
  file.arrayBuffer().then((buf) => run(new Uint8Array(buf)))
    .catch((e) => { status.textContent = `read error: ${e}`; });
});

for (const a of document.querySelectorAll('a[data-sample]')) {
  a.addEventListener('click', (ev) => {
    ev.preventDefault();
    load(a.dataset.sample);
  });
}

await init();
status.textContent = 'Ready — choose a file or click a sample.';
