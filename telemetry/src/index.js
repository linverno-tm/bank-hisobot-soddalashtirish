// SoddaHisobot telemetriyasi: ilova ochilganda o'z versiyasini yuboradi,
// shunda qaysi kompyuterda qaysi versiya ishlab turganini bilib turamiz.
//
//   POST /ping     {host, version}          -> ilovadan keladi (ochiq endpoint)
//   GET  /status   ?key=<SECRET>            -> bizga: kim, qaysi versiya, qachon
//   POST /ai       {host, model, prompt}    -> Gemini'ga vositachi
//
// STATUS_KEY va GEMINI_API_KEY `wrangler secret put ...` orqali beriladi,
// kodda saqlanmaydi.
//
// NEGA /ai KERAK: avval ilova Gemini'ga to'g'ridan-to'g'ri murojaat qilardi
// va buning uchun kalit har bir foydalanuvchining kompyuterida turishi
// kerak edi. Kalitni u yerdan ko'chirib olish oson, hisob esa bitta —
// ya'ni kalit tarqalsa, so'rovlar bizning nomimizdan ketaverardi. Endi
// kalit faqat shu yerda.
//
// MAXFIYLIK: so'rov matnida kontragent nomlari va to'lov izohlari bo'ladi.
// Ular hech qayerga yozilmaydi — na D1 ga, na jurnalga. Bu yerda faqat
// so'rovlar SONI hisoblanadi.

const json = (data, status = 200) =>
  new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });

// Kunlik cheklovlar. Bitta hisobot odatda 1 ta so'rov — ya'ni bu chegara
// oddiy ishlashda hech qachon urilmaydi. U faqat xatolik yoki suiiste'mol
// holatida hisobni himoya qiladi.
const KUNLIK_JAMI = 500;
const KUNLIK_HAR_KOMPYUTER = 100;

// Model nomi ilovadan keladi, chunki u core.py bilan birga yangilanadi va
// Worker'ni qayta joylashtirishga hojat qolmasin. Nima kelsa o'shani
// yuborib bo'lmaydi, shuning uchun shakli tekshiriladi; qimmatga tushib
// ketmasligini esa yuqoridagi kunlik chegara ushlab turadi.
const MODEL_SHAKLI = /^gemini-[a-z0-9.\-]{1,40}$/;

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

    if (request.method === "POST" && url.pathname === "/ai") {
      return await aiProxy(request, env);
    }

    return json({ error: "not found" }, 404);
  },
};

// Ilova xato sababini foydalanuvchiga tushunarli qilib ko'rsatishi uchun
// javobda aniq `error` kodi qaytadi: sozlama xatosimi, chegara urilganmi,
// yoki Gemini javob bermayaptimi — bularning har biri boshqa xabar.
async function aiProxy(request, env) {
  // Kalit `wrangler secret put` orqali quvur bilan berilganda oxiriga
  // qator ko'chirish belgisi qo'shilib qolishi mumkin — Gemini bunday
  // kalitni rad etadi va sabab "400" bo'lib ko'rinadi, xolos.
  const kalit = String(env.GEMINI_API_KEY || "").trim();
  if (!kalit) {
    return json({ ok: false, error: "config", detail: "kalit sozlanmagan" }, 500);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ ok: false, error: "bad_request" }, 400);
  }

  const host = String(body.host || "").slice(0, 100) || "noma'lum";
  const model = String(body.model || "");
  const prompt = String(body.prompt || "");

  if (!MODEL_SHAKLI.test(model)) {
    return json({ ok: false, error: "bad_request", detail: "model nomi" }, 400);
  }
  if (prompt.length < 20 || prompt.length > 200000) {
    return json({ ok: false, error: "bad_request", detail: "so'rov hajmi" }, 400);
  }

  const kun = new Date().toISOString().slice(0, 10);
  const chegara = await chegaraTekshir(env, kun, host);
  if (chegara) return json(chegara, 429);

  let javob;
  try {
    javob = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`,
      {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-goog-api-key": kalit,
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { responseMimeType: "application/json", temperature: 0 },
        }),
      }
    );
  } catch {
    return json({ ok: false, error: "upstream", detail: "ulanib bo'lmadi" }, 502);
  }

  if (!javob.ok) {
    // Gemini xatosining matnini qaytarmaymiz — unda kalit haqida ma'lumot
    // bo'lishi mumkin. Faqat holat kodi yetarli.
    return json({ ok: false, error: "upstream", detail: `holat ${javob.status}` }, 502);
  }

  let data;
  try {
    data = await javob.json();
  } catch {
    return json({ ok: false, error: "upstream", detail: "javob o'qilmadi" }, 502);
  }

  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    return json({ ok: false, error: "empty" }, 502);
  }

  await sanaOshir(env, kun, host);
  return json({ ok: true, text });
}

async function chegaraTekshir(env, kun, host) {
  const jami = await env.DB.prepare(
    `SELECT COALESCE(SUM(soni), 0) AS n FROM ai_usage WHERE kun = ?1`
  ).bind(kun).first();
  if ((jami?.n || 0) >= KUNLIK_JAMI) {
    return { ok: false, error: "limit", detail: "kunlik umumiy chegara" };
  }
  const oz = await env.DB.prepare(
    `SELECT soni FROM ai_usage WHERE kun = ?1 AND host = ?2`
  ).bind(kun, host).first();
  if ((oz?.soni || 0) >= KUNLIK_HAR_KOMPYUTER) {
    return { ok: false, error: "limit", detail: "shu kompyuter uchun kunlik chegara" };
  }
  return null;
}

async function sanaOshir(env, kun, host) {
  await env.DB.prepare(
    `INSERT INTO ai_usage (kun, host, soni) VALUES (?1, ?2, 1)
     ON CONFLICT(kun, host) DO UPDATE SET soni = soni + 1`
  ).bind(kun, host).run();
}
