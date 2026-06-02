const API_BASE_URL = "https://localhost:4100";

async function readJsonResponse(resp) {
  const raw = await resp.text();
  try {
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    const snippet = String(raw || "").trim().slice(0, 160);
    throw new Error(`Expected JSON response but received: ${snippet || resp.statusText}`);
  }
}

async function apiGet(path) {
  const resp = await fetch(`${API_BASE_URL}${path}`);
  if (!resp.ok) {
    throw new Error(`${path} failed: ${resp.statusText}`);
  }
  return readJsonResponse(resp);
}

async function apiPost(path, body) {
  const resp = await fetch(`${API_BASE_URL}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {})
  });
  const data = await readJsonResponse(resp);
  if (!resp.ok) {
    const msg = data?.message || resp.statusText;
    throw new Error(msg);
  }
  return data;
}

export { apiGet, apiPost };
