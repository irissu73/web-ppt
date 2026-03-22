export function buildTravelSlides(data) {
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
        title: "Day 1",
        items: ["09:00 出發", "10:30 景點 A", "12:00 午餐"]
      },
      {
        slideType: "bullet",
        title: "提醒",
        points: ["提早出門", "保留彈性"]
      }
    ]
  };
}