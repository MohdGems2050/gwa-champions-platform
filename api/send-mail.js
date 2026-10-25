/**
 * Vercel serverless function: send the admin credentials by email.
 *
 * Requires an email provider key in the Vercel project environment:
 *   RESEND_API_KEY   (https://resend.com — free tier is enough)
 *   MAIL_FROM        optional, e.g. "GWA Champions <onboarding@resend.dev>"
 *
 * If the key is missing the function answers { ok:false, code:"not-configured" }
 * so the UI can offer the manual (mail client) fallback instead of pretending.
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
  }
  const to = String(payload?.to || "").trim();
  const password = String(payload?.password || "");
  const code = String(payload?.code || "");

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    res.status(400).json({ ok: false, code: "invalid-email" });
    return;
  }

  const apiKey = process.env.RESEND_API_KEY;
  const from = process.env.MAIL_FROM || "GWA Champions <onboarding@resend.dev>";

  if (!apiKey) {
    res.status(200).json({ ok: false, code: "not-configured" });
    return;
  }

  const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif;background:#0d1230;padding:28px;color:#fff">
    <div style="max-width:520px;margin:0 auto;background:#121a3d;border:1px solid rgba(255,255,255,.1);border-radius:20px;padding:26px">
      <h2 style="margin:0 0 6px;color:#FDB51F">GWA Champions Portal</h2>
      <p style="margin:0 0 18px;color:#9aa4c7;font-size:14px">Administrator access details • بيانات دخول الإدارة</p>
      <div style="background:#0b1028;border-radius:14px;padding:16px;margin-bottom:16px">
        <p style="margin:0 0 8px;font-size:12px;color:#8f9ac0">PASSWORD • كلمة المرور</p>
        <p style="margin:0;font-size:20px;font-weight:800;letter-spacing:1px;color:#fff">${escapeHtml(password)}</p>
      </div>
      <div style="background:#0b1028;border-radius:14px;padding:16px;margin-bottom:16px">
        <p style="margin:0 0 8px;font-size:12px;color:#8f9ac0">RECOVERY CODE • رمز الاستعادة</p>
        <p style="margin:0;font-size:20px;font-weight:800;letter-spacing:1px;color:#FDB51F">${escapeHtml(code)}</p>
      </div>
      <p style="font-size:13px;color:#9aa4c7;line-height:1.8;margin:0 0 8px">
        Portal: <a href="https://gwa-champions-platform.vercel.app" style="color:#1697E8">gwa-champions-platform.vercel.app</a><br/>
        يمكنك تغيير كلمة المرور بعد الدخول من تبويب Security.
      </p>
      <p style="font-size:11px;color:#5d6788;margin:14px 0 0">إذا لم تطلب هذه الرسالة فتجاهلها.</p>
    </div>
  </div>`;

  try {
    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        from,
        to: [to],
        subject: "GWA Champions — Admin password & recovery code",
        html,
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      res.status(502).json({ ok: false, code: "provider-error", detail: data });
      return;
    }
    res.status(200).json({ ok: true, id: data?.id || null });
  } catch (error) {
    res.status(500).json({ ok: false, code: "network-error", detail: String(error?.message || error) });
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
