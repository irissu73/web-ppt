import pptxgen from "pptxgenjs";

function buildMockSlides({ type, title }) {
  return {
    title,
    slides: [
      {
        slideType: "cover",
        title,
        subtitle: type
      },
      {
        slideType: "bullet",
        title: "重點整理",
        points: ["第一點", "第二點", "第三點"]
      },
      {
        slideType: "timeline",
        title: "流程安排",
        items: ["步驟一", "步驟二", "步驟三"]
      }
    ]
  };
}

async function generatePptBuffer(slideJson) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI";
  pptx.subject = slideJson.title;
  pptx.title = slideJson.title;
  pptx.company = "IRIS AI Lab";
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

      const lines = (slide.points || []).map((p) => ({
        text: p,
        options: { bullet: { indent: 14 } }
      }));

      s.addText(lines, {
        x: 1.0, y: 1.5, w: 10.5, h: 4.5,
        fontSize: 18, breakLine: true
      });
      continue;
    }

    if (slide.slideType === "timeline") {
      s.addText(slide.title || "", {
        x: 0.8, y: 0.6, w: 11, h: 0.5,
        fontSize: 20, bold: true
      });

      const lines = (slide.items || []).map((item, i) => ({
        text: `${i + 1}. ${item}`
      }));

      s.addText(lines, {
        x: 1.0, y: 1.5, w: 10.5, h: 4.5,
        fontSize: 18, breakLine: true
      });
      continue;
    }

    s.addText("Unsupported slide type", {
      x: 1, y: 1, w: 10, h: 1, fontSize: 18
    });
  }

  return pptx.write({ outputType: "nodebuffer" });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method Not Allowed");
  }

  try {
    const { type, title, email, expiresAt } = req.body || {};

    if (!type || !title || !email || !expiresAt) {
      return res.status(400).json({ error: "缺少必要欄位" });
    }

    const slideJson = buildMockSlides({ type, title });
    const buffer = await generatePptBuffer(slideJson);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(title)}.pptx"`
    );

    return res.send(buffer);
  } catch (error) {
    console.error("generate error:", error);
    return res.status(500).json({ error: "PPT 產生失敗" });
  }
}