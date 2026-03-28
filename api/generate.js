import pptxgen from "pptxgenjs";

// ===== 定義 IRIS 母片 =====
function defineIrisMaster(pptx) {
  pptx.defineSlideMaster({
    title: "IRIS_MASTER",

    background: { color: "F7F7FB" },

    objects: [
      // 上方色條
      {
        rect: {
          x: 0,
          y: 0,
          w: "100%",
          h: 0.8,
          fill: { color: "4F46E5" }
        }
      },

      // footer
      {
        text: {
          text: "IRIS AI PPT",
          options: {
            x: 0.5,
            y: 6.45,
            fontSize: 10,
            color: "6B7280"
          }
        }
      }
    ],

    // 頁碼
    slideNumber: {
      x: 12.3,
      y: 6.4,
      fontSize: 10,
      color: "6B7280"
    }
  });
}

// ===== 主流程 =====
function generateSlides(pptx, data) {
  const slides = data?.data?.slides || [];

  for (const slide of slides) {
    // =========================
    // 🟣 Cover（不套母片）
    // =========================
    if (slide.slideType === "cover") {
      const s = pptx.addSlide();

      s.background = { fill: "4F46E5" };

      // 光影效果
      s.addShape(pptx.ShapeType.ellipse, {
        x: 6,
        y: 3,
        w: 6,
        h: 4,
        fill: { color: "6366F1" },
        transparency: 50
      });

      s.addShape(pptx.ShapeType.ellipse, {
        x: -2,
        y: -1,
        w: 5,
        h: 3,
        fill: { color: "818CF8" },
        transparency: 60
      });

      // 主標
      s.addText(slide.title || "", {
        x: 1,
        y: 2.2,
        w: 8,
        fontSize: 40,
        bold: true,
        color: "FFFFFF",
        align: "center"
      });

      // 副標
      if (slide.subtitle) {
        s.addText(slide.subtitle, {
          x: 1,
          y: 3.3,
          w: 8,
          fontSize: 20,
          color: "E0E7FF",
          align: "center"
        });
      }

      continue;
    }

    // =========================
    // 🔵 Timeline（套母片）
    // =========================
    if (slide.slideType === "timeline") {
      const s = pptx.addSlide({ masterName: "IRIS_MASTER" });

      // 標題（白字放在色條上）
      s.addText(slide.title || "", {
        x: 0.7,
        y: 0.2,
        fontSize: 20,
        bold: true,
        color: "FFFFFF"
      });

      // 分隔線
      s.addShape(pptx.ShapeType.rect, {
        x: 0.5,
        y: 1.2,
        w: 9,
        h: 0.03,
        fill: { color: "E5E7EB" }
      });

      const items = slide.items || [];

      items.forEach((item, i) => {
        s.addText(`• ${item}`, {
          x: 0.8,
          y: 1.5 + i * 0.6,
          fontSize: 18,
          color: "111827"
        });
      });

      continue;
    }

    // =========================
    // 🟢 Bullet（套母片）
    // =========================
    if (slide.slideType === "bullet") {
      const s = pptx.addSlide({ masterName: "IRIS_MASTER" });

      s.addText(slide.title || "", {
        x: 0.7,
        y: 0.2,
        fontSize: 20,
        bold: true,
        color: "FFFFFF"
      });

      const points = slide.points || [];

      points.forEach((p, i) => {
        s.addText(`• ${p}`, {
          x: 1,
          y: 1.5 + i * 0.6,
          fontSize: 20,
          color: "111827"
        });
      });

      continue;
    }
  }
}

// ===== API handler =====
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const data = req.body;

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";

    // ⭐ 定義母片
    defineIrisMaster(pptx);

    // ⭐ 產生內容
    generateSlides(pptx, data);

    const buffer = await pptx.write("nodebuffer");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${data.title || "presentation"}.pptx"`
    );

    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "產生簡報失敗" });
  }
}