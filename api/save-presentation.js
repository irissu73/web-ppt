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

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const presentation = req.body || {};

    if (!presentation.presentationId) {
      return res.status(400).json({ error: "缺少 presentationId" });
    }

    // ===== 先查舊資料（判斷是否存在 & expiresAt 是否改變）=====
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
    const oldExpiresAt = existing?.expiresAt || null;
    const newExpiresAt = presentation.expiresAt || null;

    // ===== upsert =====
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

    // ===== 判斷是否要寄信 =====
    const shouldSendEmail =
      systemConfig.sendEmailEnabled &&
      presentation.email &&
      presentation.email.trim() !== "";

    if (shouldSendEmail) {
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