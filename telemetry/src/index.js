// SoddaHisobot telemetriyasi: ilova ochilganda o'z versiyasini yuboradi,
// shunda qaysi kompyuterda qaysi versiya ishlab turganini bilib turamiz.
//
//   POST /ping     {host, version}     -> ilovadan keladi (ochiq endpoint)
//   GET  /status   ?key=<SECRET>       -> bizga: kim, qaysi versiya, qachon
//
// STATUS_KEY `wrangler secret put STATUS_KEY` orqali beriladi, kodda saqlanmaydi.

const json = (data, status = 200) =>
  new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "POST" && url.pathname === "/ping") {
      let body;
      try {
        body = await request.json();
      } catch {
        return json({ error: "invalid json" }, 400);
      }
      const host = String(body.host || "").slice(0, 100) || "noma'lum";
      const version = String(body.version || "").slice(0, 20) || "?";
      const now = new Date().toISOString();

      await env.DB.prepare(
        `INSERT INTO clients (host, version, first_seen, last_seen, ping_count)
         VALUES (?1, ?2, ?3, ?3, 1)
         ON CONFLICT(host) DO UPDATE SET
           version = ?2,
           last_seen = ?3,
           ping_count = ping_count + 1`
      ).bind(host, version, now).run();

      return json({ ok: true });
    }

    if (request.method === "GET" && url.pathname === "/status") {
      if (url.searchParams.get("key") !== env.STATUS_KEY) {
        return json({ error: "forbidden" }, 403);
      }
      const { results } = await env.DB.prepare(
        `SELECT host, version, last_seen, first_seen, ping_count
         FROM clients ORDER BY last_seen DESC`
      ).all();
      return json({ clients: results });
    }

    return json({ error: "not found" }, 404);
  },
};
