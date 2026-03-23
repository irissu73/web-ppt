import pptxgen from "pptxgenjs";
import { buildTravelSlides } from "../builders/buildTravelSlides.js";
import { sendNotificationEmail } from "../lib/sendNotificationEmail.js";
import { systemConfig } from "../config/systemConfig.js";

function buildCompareSlides(data) {
  return {
    title: data.title,
    slides: [
      {
        slideType: "cover",
        title: data.title,
        subtitle: "比較分析"
      },
      {
        slideType: "compare",
        title: "主要比較",
        leftTitle: data.data?.leftTitle || "左側",
        rightTitle: data.data?.rightTitle || "右側",
        leftPoints: ["重點 1", "重點 2"],
        rightPoints: ["重點 1", "重點 2"]
      },
      {
        slideType: "bullet",
        title: "結論",
        points: ["整理差異", "補充說明"]
      }
    ]
  };
}

function buildLessonSlides(data) {
  return {
    title: data.title,
    slides: [
      {
        slideType: "cover",
        title: data.title,
        subtitle: "教學說明"
      },
      {
        slideType: "bullet",
        title: data.data?.lessonTopic || "教學主題",
        points: [
          `對象：${data.data?.lessonAudience || "未填寫"}`,
          "整理重點",
          "進入說明"
        ]
      }
    ]
  };
}

function buildProposalSlides(data) {
  return {
    title: data.title,
    slides: [
      {
        slideType: "cover",
        title: data.title,
        subtitle: "提案簡報"
      },
      {
        slideType: "bullet",
        title: "提案資訊",
        points: [
          `提案對象：${data.data?.proposalTarget || "未填寫"}`,
          `提案目的：${data.data?.proposalGoal || "未填寫"}`
        ]
      }
    ]
  };
}

function buildSlidesByType(data) {
  switch (data.type) {
    case "travel":
      return buildTravelSlides(data);
    case "compare":
      return buildCompareSlides(data);
    case "lesson":
      return buildLessonSlides(data);
    case "proposal":
      return buildProposalSlides(data);
    default:
      throw new Error("不支援的類型");
  }
}

async function generatePptBuffer(slideJson) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "IRIS AI PPT";
  pptx.company = "IRIS AI PPT";
  pptx.subject = slideJson.title;
  pptx.title = slideJson.title;
  pptx.lang = "zh-TW";

  for (const slide of slideJson.slides) {
    const s = pptx.addSlide();

    if (slide.slideType === "cover") {
      s.addText(slide.title || "", {
        x: 0.8, y: 1.2, w: 11, h: 0.8,
        fontSize: 24, bold: true
      });

      s.addText(slide.subtitle || "", {
        x: 0.8, y: 2.2, w: 11, h: 0.5,
        fontSize: 14
      });

      continue;
    }

    if (slide.slideType === "bullet") {
      s.addText(slide.title || "", {
        x: 0.8, y: 0.6, w: 11, h: 0.5,
        fontSize: 20, bold: true
      });

      const text = (slide.points || []).map((p) => `• ${p}`).join("\n");

      s.addText(text, {
        x: 1.0, y: 1.5, w: 10.5, h: 4.5,
        fontSize: 18
      });

      continue;
    }

    if (slide.slideType === "timeline") {
      s.addText(slide.title || "", {
        x: 0.8, y: 0.6, w: 11, h: 0.5,
        fontSize: 20, bold: true
      });

      const text = (slide.items || [])
        .map((item, i) => `${i + 1}. ${item}`)
        .join("\n");

      s.addText(text, {
        x: 1.0, y: 1.5, w: 10.5, h: 4.5,
        fontSize: 18
      });

      continue;
    }

    if (slide.slideType === "compare") {
      s.addText(slide.title || "", {
        x: 0.8, y: 0.6, w: 11, h: 0.5,
        fontSize: 20, bold: true
      });

      s.addText(slide.leftTitle || "", {
        x: 0.8, y: 1.4, w: 5, h: 0.4,
        fontSize: 18, bold: true
      });

      s.addText(slide.rightTitle || "", {
        x: 6.7, y: 1.4, w: 5, h: 0.4,
        fontSize: 18, bold: true
      });

      s.addText((slide.leftPoints || []).map(p => `• ${p}`).join("\n"), {
        x: 0.8, y: 2.0, w: 5, h: 4,
        fontSize: 16
      });

      s.addText((slide.rightPoints || []).map(p => `• ${p}`).join("\n"), {
        x: 6.7, y: 2.0, w: 5, h: 4,
        fontSize: 16
      });

      continue;
    }

    s.addText("Unsupported slide type", {
      x: 1, y: 1, w: 10, h: 1, fontSize: 18
    });
  }

  return pptx.write({ outputType: "nodebuffer" });
}

function buildEditUrl(data) {
  const baseUrl = process.env.APP_BASE_URL || "";
  const presentationId = data.presentationId;

  return `${baseUrl}/${data.type}.html?presentationId=${presentationId}`;
}

function validateExpireDate(expiresAt) {
  if (!expiresAt) return false;

  const today = new Date();
  const expireDate = new Date(expiresAt);

  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const expireOnly = new Date(expireDate.getFullYear(), expireDate.getMonth(), expireDate.getDate());

  const diffMs = expireOnly - todayOnly;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);

  return diffDays >= 0 && diffDays <= systemConfig.maxExpireDays;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method Not Allowed");
  }

  try {
    const data = req.body || {};

    if (!data.presentationId || !data.type || !data.title || !data.expiresAt) {
      return res.status(400).json({ error: "缺少必要欄位" });
    }

    if (systemConfig.requireEmail && !data.email) {
      return res.status(400).json({ error: "Email 必填" });
    }

    if (!validateExpireDate(data.expiresAt)) {
      return res.status(400).json({ error: "最後保留日超出允許範圍" });
    }

    const slideJson = buildSlidesByType(data);
    const buffer = await generatePptBuffer(slideJson);
    const editUrl = buildEditUrl(data);

    try {
      await sendNotificationEmail({
        status: "created",
        email: data.email,
        title: data.title,
        type: data.type,
        expiresAt: data.expiresAt,
        editUrl
      });
    } catch (mailError) {
      console.error("寄信失敗，但簡報已成功建立：", mailError);
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(data.title || "presentation")}.pptx"`
    );

    return res.send(buffer);
  } catch (error) {
    console.error("generate error:", error);
    return res.status(500).json({
      error: error?.message || "PPT 產生失敗"
    });
  }
}