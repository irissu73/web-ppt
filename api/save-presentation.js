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

    console.log("save-presentation input =", presentation);

    if (!presentation.presentationId) {
      return res.status(400).json({ error: "缺少 presentationId" });
    }

    const { data, error } = await supabase
      .from("presentations")
      .insert({
        presentationId: presentation.presentationId,
        type: presentation.type,
        title: presentation.title,
        email: presentation.email,
        expiresAt: presentation.expiresAt,
        data: presentation.data || {}
      })
      .select();

    if (error) {
      console.error("supabase insert error =", error);
      return res.status(500).json({
        error: "儲存失敗",
        detail: error.message
      });
    }

    console.log("supabase inserted data =", data);

    return res.status(200).json({
      message: "saved",
      saved: data
    });
  } catch (err) {
    console.error("save-presentation catch error =", err);
    return res.status(500).json({
      error: "儲存失敗",
      detail: err?.message || String(err)
    });
  }
}