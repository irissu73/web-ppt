import pptxgen from "pptxgenjs";

// 16:9 寬螢幕約 13.333 x 7.5
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

function defineIrisMaster(pptx) {
  pptx.defineSlideMaster({
    title: "IRIS_MASTER",
    background: { color: "F7F7FB" },
    objects: [
      {
        rect: {
          x: 0,
          y: 0,
          w: SLIDE_W,
          h: 0.8,
          fill: { color: "4F46E5" },
          line: { color: "4F46E5" }
        }
      },
      {
        text: {
          text: "IRIS AI PPT",
          options: {
            x: 0.5,
            y: 7.05,
            w: 2,
            h: 0.2,
            fontSize: 10,
            color: "6B7280"
          }
        }
      }
    ],
    slideNumber: {
      x: 12.3,
      y: 7.02,
      w: 0.5,
      h: 0.2,
      fontSize: 10,
      color: "6B7280",
      align: "right"
    }
  });
}

function addCoverSlide(pptx, slide) {
  const s = pptx.addSlide();

  s.background = { color: "4F46E5" };

  // 光影效果
  s.addShape("ellipse", {
    x: 7.2,
    y: 3.2,
    w: 5.2,
    h: 3.2,
    fill: { color: "6366F1", transparency: 50 },
    line: { color: "6366F1", transparency: 100 }
  });

  s.addShape("ellipse", {
    x: -1.2,
    y: -0.5,
    w: 4.2,
    h: 2.5,
    fill: { color: "818CF8", transparency: 60 },
    line: { color: "818CF8", transparency: 100 }
  });

  s.addText(slide.title || "", {
    x: 1.2,
    y: 2.2,
    w: 10.8,
    h: 0.8,
    fontSize: 28,
    bold: true,
    color: "FFFFFF",
    align: "center",
    valign: "mid"
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 1.2,
      y: 3.15,
      w: 10.8,
      h: 0.5,
      fontSize: 16,
      color: "E0E7FF",
      align: "center",
      valign: "mid"
    });
  }

  s.addText("IRIS AI PPT", {
    x: 0.5,
    y: 7.02,
    w: 2,
    h: 0.2,
    fontSize: 10,
    color: "C7D2FE"
  });
}

function addTimelineSlide(pptx, slide) {
  const s = pptx.addSlide({ masterName: "IRIS_MASTER" });

  s.addText(slide.title || "", {
    x: 0.7,
    y: 0.18,
    w: 6,
    h: 0.3,
    fontSize: 20,
    bold: true,
    color: "FFFFFF"
  });

  s.addShape("line", {
    x: 0.7,
    y: 1.25,
    w: 11.5,
    h: 0,
    line: { color: "E5E7EB", pt: 1 }
  });

  const items = slide.items || [];
  items.forEach((item, i) => {
    const y = 1.55 + i * 0.58;

    s.addShape("ellipse", {
      x: 0.78,
      y: y + 0.08,
      w: 0.12,
      h: 0.12,
      fill: { color: "4F46E5" },
      line: { color: "4F46E5" }
    });

    s.addText(item, {
      x: 1.05,
      y,
      w: 10.5,
      h: 0.3,
      fontSize: 18,
      color: "111827"
    });
  });
}

function addBulletSlide(pptx, slide) {
  const s = pptx.addSlide({ masterName: "IRIS_MASTER" });

  s.addText(slide.title || "", {
    x: 0.7,
    y: 0.18,
    w: 6,
    h: 0.3,
    fontSize: 20,
    bold: true,
    color: "FFFFFF"
  });

  const points = slide.points || [];
  points.forEach((p, i) => {
    s.addText(`• ${p}`, {
      x: 1,
      y: 1.5 + i * 0.6,
      w: 10.5,
      h: 0.3,
      fontSize: 20,
      color: "111827"
    });
  });
}

function generateSlides(pptx, data) {
  const slides = data?.data?.slides || [];

  for (const slide of slides) {
    if (slide.slideType === "cover") {
      addCoverSlide(pptx, slide);
      continue;
    }

    if (slide.slideType === "timeline") {
      addTimelineSlide(pptx, slide);
      continue;
    }

    if (slide.slideType === "bullet") {
      addBulletSlide(pptx, slide);
      continue;
    }

    // 不支援頁型時，至少不要整份炸掉
    const s = pptx.addSlide({ masterName: "IRIS_MASTER" });
    s.addText(`Unsupported slideType: ${slide.slideType || "unknown"}`, {
      x: 1,
      y: 1.5,
      w: 10,
      h: 0.5,
      fontSize: 20,
      color: "111827"
    });
  }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const data = req.body || {};

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "IRIS AI PPT";
    pptx.company = "IRIS AI PPT";
    pptx.subject = data.title || "presentation";
    pptx.title = data.title || "presentation";
    pptx.lang = "zh-TW";

    defineIrisMaster(pptx);
    generateSlides(pptx, data);

    const buffer = await pptx.write({ outputType: "nodebuffer" });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(data.title || "presentation")}.pptx"`
    );

    return res.send(buffer);
  } catch (err) {
    console.error("generate error =", err);
    return res.status(500).json({
      error: err?.message || "產生簡報失敗"
    });
  }
}