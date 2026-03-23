import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const presentation = req.body || {};

    if (!presentation.presentationId) {
      return res.status(400).json({ error: "缺少 presentationId" });
    }

    const { data, error } = await supabase
      .from("presentations")
      .upsert({
        presentationId: presentation.presentationId,
        type: presentation.type,
        title: presentation.title,
        email: presentation.email,
        expiresAt: presentation.expiresAt,
        data: presentation.data || {}
      })
      .select();

    if (error) {
      console.error("supabase upsert error =", error);
      return res.status(500).json({
        error: "儲存失敗",
        detail: error.message
      });
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