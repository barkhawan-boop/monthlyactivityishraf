const DATA_KEY = "monthlyactivityishraf:shared-data";

function jsonResponse(value, init = {}) {
  return new Response(JSON.stringify(value), {
    ...init,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      ...(init.headers || {}),
    },
  });
}

async function handleDataRequest(request, env) {
  if (!env.ACTIVITY_STORE) {
    return jsonResponse({ error: "Missing ACTIVITY_STORE KV binding" }, { status: 500 });
  }

  if (request.method === "GET") {
    const value = await env.ACTIVITY_STORE.get(DATA_KEY);
    return new Response(value || "null", {
      headers: {
        "content-type": "application/json; charset=utf-8",
        "cache-control": "no-store",
      },
    });
  }

  if (request.method === "POST") {
    const body = await request.text();
    if (body.length > 1_000_000) {
      return jsonResponse({ error: "Data is too large" }, { status: 413 });
    }
    try {
      const parsed = JSON.parse(body);
      if (!parsed || !Array.isArray(parsed.inspectors) || !parsed.global) {
        return jsonResponse({ error: "Invalid data shape" }, { status: 400 });
      }
    } catch {
      return jsonResponse({ error: "Invalid JSON" }, { status: 400 });
    }
    await env.ACTIVITY_STORE.put(DATA_KEY, body);
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: "Method not allowed" }, { status: 405 });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/api/data") {
      return handleDataRequest(request, env);
    }
    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }
    return new Response("Not found", { status: 404 });
  },
};
