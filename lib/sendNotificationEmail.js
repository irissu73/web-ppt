import { Resend } from "resend";
import { systemConfig } from "../config/systemConfig.js";

const resend = new Resend(process.env.RESEND_API_KEY);

const statusTextMap = {
  created: "已建立",
  updated: "已更新保留日"
};

const typeMap = {
  travel: "行程安排",
  compare: "比較分析",
  lesson: "教學說明",
  proposal: "提案簡報"
};

export async function sendNotificationEmail({
  status,
  email,
  title,
  type,
  expiresAt,
  editUrl
}) {
  if (!systemConfig.sendEmailEnabled) return;
  if (!email) return;

  const statusText = statusTextMap[status] || "";
  const typeLabel = typeMap[type] || type;

  const subject = `【AI幫你做簡報】${title}（保留至 ${expiresAt}）`;

  const html = `
    <p>您好，</p>

    <p>你的簡報${statusText}。</p>

    <p><strong>簡報資訊</strong></p>
    <ul>
      <li>標題：${title}</li>
      <li>類型：${typeLabel}</li>
      <li>保留至：${expiresAt}</li>
    </ul>

    <p>
      👉 <a href="${editUrl}">點這裡開啟或繼續編輯</a>
    </p>

    <br/>

    <p style="color:#888;font-size:12px;">
      ※ 此為系統自動發送信件，請勿回覆<br/>
      ※ 超過保留期限後，資料將自動刪除
    </p>
  `;

  return resend.emails.send({
    from: process.env.MAIL_FROM,
    to: email,
    subject,
    html
  });
}