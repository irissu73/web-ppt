import { createClient } from "@supabase/supabase-js";
import { sendNotificationEmail } from "../lib/sendNotificationEmail.js";
import { systemConfig } from "../config/systemConfig.js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

function buildEditUrl({ type, presentationId }) {
  const baseUrl = process.env.APP_BASE_URL || "";
  return `${baseUrl}/${type}.html?presentationId=${presentationId}`;
}

function normalizeDateString(value) {
  if (!value) return "";
  // 只取 YYYY-MM-DD，避免時間或格式差異誤判
  return String(value).slice(0, 10);
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const presentation = req.body || {};

    if (!presentation.presentationId) {
      return res.status(400).json({ error: "缺少 presentationId" });
    }

    // 先查舊資料
    const { data: existing, error: fetchError } = await supabase
      .from("presentations")
      .select("*")
      .eq("presentationId", presentation.presentationId)
      .maybeSingle();

    if (fetchError) {
      console.error("fetch existing error =", fetchError);
      return res.status(500).json({ error: "讀取舊資料失敗" });
    }

    const isNew = !existing;
    const oldExpiresAt = normalizeDateString(existing?.expiresAt);
    const newExpiresAt = normalizeDateString(presentation.expiresAt);

    // 儲存
    const { data, error } = await supabase
      .from("presentations")
      .upsert(
        {
          presentationId: presentation.presentationId,
          type: presentation.type,
          title: presentation.title,
          email: presentation.email,
          expiresAt: presentation.expiresAt,
          data: presentation.data || {}
        },
        {
          onConflict: "presentationId"
        }
      )
      .select();

    if (error) {
      console.error("supabase upsert error =", error);
      return res.status(500).json({
        error: "儲存失敗",
        detail: error.message
      });
    }

    // 只有「新建」或「保留日改變」才寄
    const hasEmail =
      !!presentation.email && presentation.email.trim() !== "";

    const shouldConsiderSend =
      systemConfig.sendEmailEnabled && hasEmail;

    if (shouldConsiderSend) {
      let status = null;

      if (isNew) {
        status = "created";
      } else if (oldExpiresAt !== newExpiresAt) {
        status = "updated";
      }

      if (status) {
        try {
          const editUrl = buildEditUrl({
            type: presentation.type,
            presentationId: presentation.presentationId
          });

          await sendNotificationEmail({
            status,
            email: presentation.email,
            title: presentation.title,
            type: presentation.type,
            expiresAt: presentation.expiresAt,
            editUrl
          });

          console.log("email sent:", status);
        } catch (mailError) {
          console.error("寄信失敗（不影響主流程）：", mailError);
        }
      }
    }

    return res.status(200).json({
      message: "儲存成功",
      saved: data
    });
  } catch (err) {
    console.error("save-presentation error =", err);
    return res.status(500).json({
      error: "儲存失敗",
      detail: err?.message || String(err)
    });
  }
}