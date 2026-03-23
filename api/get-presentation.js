import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const { presentationId } = req.query;

    if (!presentationId) {
      return res.status(400).json({ error: "缺少 presentationId" });
    }

    const { data, error } = await supabase
      .from("presentations")
      .select("*")
      .eq("presentationId", presentationId)
      .single();

    if (error || !data) {
      return res.status(404).json({ error: "找不到資料" });
    }

    return res.status(200).json(data);
  } catch (err) {
    console.error("get-presentation error =", err);
    return res.status(500).json({ error: "讀取失敗" });
  }
}