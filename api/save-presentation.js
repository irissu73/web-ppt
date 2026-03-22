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
    const presentation = req.body;

    const { error } = await supabase
      .from("presentations")
      .upsert({
        presentationId: presentation.presentationId,
        type: presentation.type,
        title: presentation.title,
        email: presentation.email,
        expiresAt: presentation.expiresAt,
        data: presentation.data || {}
      });

    if (error) throw error;

    return res.status(200).json({ message: "saved" });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: "save failed" });
  }
}