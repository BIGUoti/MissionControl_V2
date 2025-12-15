import { getStore } from "@netlify/blobs";
import type { Context } from "@netlify/functions";

const store = getStore("missions");

function makeHeaders(origin?: string) {
  const o = origin || "*";
  return {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    "access-control-allow-origin": o,
    "access-control-allow-methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS",
    "access-control-allow-headers": "content-type",
  };
}

function resp(body: unknown, status = 200, origin?: string) {
  return new Response(JSON.stringify(body), { status, headers: makeHeaders(origin) });
}

export default async (req: Request, _context: Context) => {
  const origin = req.headers.get("origin") || undefined;

  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: makeHeaders(origin) });
  }

  const url = new URL(req.url);
  const code = (url.searchParams.get("code") || "").trim();
  if (!code) return resp({ ok: false, error: "Missing mission code (?code=)" }, 400, origin);

  const key = `mission:${code}`;

  if (req.method === "GET") {
    const mission = await store.get(key, { type: "json" }).catch(() => null);
    return resp({ ok: true, mission: mission || null }, 200, origin);
  }

  if (req.method === "DELETE") {
    await store.delete(key);
    return resp({ ok: true }, 200, origin);
  }

  if (req.method === "POST" || req.method === "PUT" || req.method === "PATCH") {
    let payload: any;
    try {
      payload = await req.json();
    } catch {
      return resp({ ok: false, error: "Invalid JSON" }, 400, origin);
    }

    const clientVersion = typeof payload?.version === "number" ? payload.version : null;
    const existing: any = await store.get(key, { type: "json" }).catch(() => null);
    const existingVersion = typeof existing?.version === "number" ? existing.version : 0;

    if (clientVersion !== null && clientVersion !== existingVersion) {
      return resp(
        { ok: false, error: "Version conflict", serverVersion: existingVersion, mission: existing || null },
        409,
        origin
      );
    }

    const nextVersion = existingVersion + 1;
    const mission = {
      ...(existing || {}),
      ...(payload || {}),
      version: nextVersion,
      updatedAt: Date.now(),
    };

    await store.setJSON(key, mission);
    return resp({ ok: true, mission }, 200, origin);
  }

  return resp({ ok: false, error: "Method not allowed" }, 405, origin);
};
