export function buildTravelSlides(data) {
  const savedSlides = data?.data?.slides || [];

  if (savedSlides.length > 0) {
    return {
      title: data.title,
      slides: savedSlides
    };
  }

  // 如果真的沒有 slides，才退回預設內容
  return {
    title: data.title,
    slides: [
      {
        slideType: "cover",
        title: data.title,
        subtitle: "行程安排"
      },
      {
        slideType: "timeline",
        title: `Day 1｜${data.data?.location || "未填地點"}`,
        items: [
          "09:00 出發",
          "10:30 景點 A",
          "12:00 午餐"
        ]
      },
      {
        slideType: "bullet",
        title: "提醒",
        points: [
          `天數：${data.data?.days || 1} 天`,
          `地點：${data.data?.location || "未填寫"}`,
          "保留彈性"
        ]
      }
    ]
  };
}