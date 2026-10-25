/**
 * Vercel serverless function: email the admin credentials.
 *
 * Env vars (Vercel → Settings → Environment Variables):
 *   RESEND_API_KEY  required — https://resend.com (free plan is enough)
 *   MAIL_FROM       optional — e.g. "GWA Champions <noreply@yourdomain.com>"
 *
 * Resend "testing mode" only allows sending FROM onboarding@resend.dev TO the
 * account owner's inbox. To stay working in every situation we try the configured
 * sender first, then fall back to onboarding@resend.dev automatically.
 */
export default async function handler(req, res) {
  res.setHeader("Content-Type", "application/json");

  if (req.method !== "POST") {
    res.status(405).json({ ok: false, code: "method-not-allowed" });
    return;
  }

  let payload = req.body;
  if (typeof payload === "string") {
    try { payload = JSON.parse(payload); } catch { payload = {}; }
  } else if (payload && typeof payload === "object" && typeof payload.to === "undefined" && typeof payload.body !== "undefined") {
    payload = typeof payload.body === "string" ? safeJson(payload.body) : payload.body;
  }
  payload = payload || {};

  const to = String(payload.to || "").trim();
  const password = String(payload.password || "");
  const code = String(payload.code || "");

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    res.status(400).json({ ok: false, code: "invalid-email" });
    return;
  }

  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    res.status(200).json({ ok: false, code: "not-configured" });
    return;
  }

  const configuredFrom = (process.env.MAIL_FROM || "").trim();
  const senders = [];
  if (configuredFrom) senders.push(configuredFrom);
  if (!senders.includes("GWA Champions <onboarding@resend.dev>")) {
    senders.push("GWA Champions <onboarding@resend.dev>");
  }

  const html = buildHtml(password, code);
  const subject = "GWA Champions — Admin password & recovery code";

  let lastError = null;
  let lastStatus = 502;

  for (const from of senders) {
    try {
      const response = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
        body: JSON.stringify({ from, to: [to], subject, html }),
      });
      const data = await response.json().catch(() => ({}));

      if (response.ok) {
        res.status(200).json({ ok: true, id: data?.id || null, from });
        return;
      }

      lastError = { from, status: response.status, message: data?.message || data?.name || "provider error" };
      lastStatus = response.status === 403 ? 403 : 502;

      const msg = String(data?.message || "").toLowerCase();
      if (response.status === 401 || msg.includes("api key")) break;
    } catch (error) {
      lastError = { from, status: 0, message: String(error?.message || error) };
    }
  }

  const message = String(lastError?.message || "").toLowerCase();
  const failCode = message.includes("only send testing emails")
    ? "recipient-not-allowed"
    : lastError
      ? "provider-error"
      : "unknown";

  res.status(lastStatus).json({ ok: false, code: failCode, detail: lastError });
}

function safeJson(value) {
  try { return JSON.parse(value); } catch { return {}; }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildHtml(password, code) {
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;background:#0d1230;padding:28px;color:#ffffff">
    <div style="max-width:520px;margin:0 auto;background:#121a3d;border:1px solid rgba(255,255,255,.12);border-radius:20px;padding:26px">
      <h2 style="margin:0 0 6px;color:#FDB51F">GWA Champions Portal</h2>
      <p style="margin:0 0 18px;color:#9aa4c7;font-size:14px">Administrator access details • بيانات دخول الإدارة</p>
      <div style="background:#0b1028;border-radius:14px;padding:16px;margin-bottom:14px">
        <p style="margin:0 0 8px;font-size:12px;color:#8f9ac0">PASSWORD • كلمة المرور</p>
        <p style="margin:0;font-size:20px;font-weight:800;letter-spacing:1px;color:#ffffff">${escapeHtml(password)}</p>
      </div>
      <div style="background:#0b1028;border-radius:14px;padding:16px;margin-bottom:14px">
        <p style="margin:0 0 8px;font-size:12px;color:#8f9ac0">RECOVERY CODE • رمز الاستعادة</p>
        <p style="margin:0;font-size:20px;font-weight:800;letter-spacing:1px;color:#FDB51F">${escapeHtml(code)}</p>
      </div>
      <p style="font-size:13px;color:#9aa4c7;line-height:1.9;margin:0">
        Portal: <a href="https://gwa-champions-platform.vercel.app" style="color:#1697E8">gwa-champions-platform.vercel.app</a><br/>
        بعد الدخول يمكنك تغيير كلمة المرور من تبويب Security.
      </p>
      <p style="font-size:11px;color:#5d6788;margin:14px 0 0">إذا لم تطلب هذه الرسالة فتجاهلها.</p>
    </div>
  </div>`;
}
