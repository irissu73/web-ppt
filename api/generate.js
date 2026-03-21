import pptxgen from "pptxgenjs";

function buildMockSlides({ type, title }) {
  if (type === "travel") {
    return {
      title,
      slides: [
        {
          slideType: "cover",
          title,
          subtitle: "行程安排"
        },
        {
          slideType: "timeline",
          title: "Day 1 行程",
          items: ["09:00 出發", "10:30 景點 A", "12:00 午餐"]
        },
        {
          slideType: "bullet",
          title: "行程提醒",
          points: ["提早出門", "注意天氣", "保留休息時間"]
        }
      ]
    };
  }

  if (type === "compare") {
    return {
      title,
      slides: [
        {
          slideType: "cover",
          title,
          subtitle: "比較分析"
        },
        {
          slideType: "bullet",
          title: "比較面向",
          points: ["價格", "功能", "適合對象"]
        },
        {
          slideType: "bullet",
          title: "初步結論",
          points: ["A 適合重視穩定", "B 適合重視彈性", "可依需求選擇"]
        }
      ]
    };
  }

  if (type === "lesson") {
    return {
      title,
      slides: [
        {
          slideType: "cover",
          title,
          subtitle: "教學說明"
        },
        {
          slideType: "bullet",
          title: "今天要學什麼",
          points: ["基本概念", "實際例子", "重點整理"]
        },
        {
          slideType: "bullet",
          title: "學習重點",
          points: ["先理解概念", "再看應用", "最後做整理"]
        }
      ]
    };
  }

  if (type === "proposal") {
    return {
      title,
      slides: [
        {
          slideType: "cover",
          title,
          subtitle: "提案簡報"
        },
        {
          slideType: "bullet",
          title: "目前問題",
          points: ["流程繁瑣", "溝通成本高", "產出速度慢"]
        },
        {
          slideType: "bullet",
          title: "提案方向",
          points: ["導入 AI 協助", "標準化流程", "提升產出效率"]
        }
      ]
    };
  }

  return {
    title,
    slides: [
      {
        slideType: "cover",
        title,
        subtitle: "未指定類型"
      },
      {
        slideType: "bullet",
        title: "重點整理",
        points: ["第一點", "第二點", "第三點"]
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

  const text = (slide.items || [])
    .map((item, i) => `${i + 1}. ${item}`)
    .join("\n");

  s.addText(text, {
    x: 1.0, y: 1.5, w: 10.5, h: 4.5,
    fontSize: 18,
    breakLine: false,
    valign: "top",
    margin: 0.1
  });
  continue;
}

if (slide.slideType === "compare") {
  s.addText(slide.title || "", {
    x: 0.8, y: 0.6, w: 11, h: 0.5,
    fontSize: 20, bold: true
  });

  s.addText(slide.leftTitle || "左側", {
    x: 0.8, y: 1.4, w: 5, h: 0.4,
    fontSize: 18, bold: true
  });

  s.addText(slide.rightTitle || "右側", {
    x: 6.7, y: 1.4, w: 5, h: 0.4,
    fontSize: 18, bold: true
  });

  const leftText = (slide.leftPoints || []).map(p => `• ${p}`).join("\n");
  const rightText = (slide.rightPoints || []).map(p => `• ${p}`).join("\n");

  s.addText(leftText, {
    x: 0.8, y: 2.0, w: 5, h: 4,
    fontSize: 16,
    valign: "top",
    margin: 0.1
  });

  s.addText(rightText, {
    x: 6.7, y: 2.0, w: 5, h: 4,
    fontSize: 16,
    valign: "top",
    margin: 0.1
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