import { Resend } from "resend";
import { systemConfig } from "../config/systemConfig.js";

const resend = new Resend(process.env.RESEND_API_KEY);

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const { email } = req.body || {};

    if (systemConfig.requireEmail && !email) {
      return res.status(400).json({ error: "Email 必填" });
    }

    // 👇 先確認 API 有沒有吃到 env
    console.log("MAIL_FROM =", process.env.MAIL_FROM);
    console.log("RESEND_API_KEY exists =", !!process.env.RESEND_API_KEY);

    if (!systemConfig.sendEmailEnabled) {
      return res.status(200).json({
        message: "寄信功能目前關閉（測試模式）"
      });
    }

    const result = await resend.emails.send({
      from: process.env.MAIL_FROM,
      to: email,
      subject: "Web PPT 測試信",
      html: `
        <p>這是一封測試信</p>
        <p>如果你收到，代表成功 🎉</p>
      `
    });

    return res.status(200).json({
      message: "寄信成功",
      result
    });

  } catch (error) {
    console.error("send-test-email error:", error);
    return res.status(500).json({ error: "寄信失敗" });
  }
}