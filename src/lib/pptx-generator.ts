import pptxgen from "pptxgenjs";
import { saveAs } from "file-saver";

export interface ReportData {
  storeName: string;
  completionDate: string;
  type: string;
  monitorLeft: string;
  monitorRight: string;
  images: {
    before?: string;
    after?: string;
    front?: string;
    sideLeft?: string;
    sideRight?: string;
    other1?: string;
    other2?: string;
  };
}

async function getBase64FromUrl(url: string): Promise<string | null> {
  if (!url) return null;
  if (url.startsWith('data:')) return url;
  try {
    const response = await fetch(url);
    if (!response.ok) return null;
    const blob = await response.blob();
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error("Fetch error:", error);
    return null;
  }
}

export async function generatePPT(data: ReportData) {
  const pptx = new pptxgen();
  const fetchedImages: Record<string, string | null> = {};
  
  // Sequential fetch to avoid overwhelming
  const keys = Object.keys(data.images);
  for (const key of keys) {
    const url = (data.images as any)[key];
    if (url) fetchedImages[key] = await getBase64FromUrl(url);
  }

  pptx.layout = "LAYOUT_WIDE";

  const addHeader = (slide: any) => {
    const tableData = [
      [{ text: "店舗名", options: { fill: { color: "D9E2F3" } } }, { text: data.storeName }],
      [{ text: "設置日", options: { fill: { color: "D9E2F3" } } }, { text: data.completionDate }],
      [{ text: "Type", options: { fill: { color: "D9E2F3" } } }, { text: data.type }],
      [{ text: "モニターL", options: { fill: { color: "D9E2F3" } } }, { text: data.monitorLeft }],
      [{ text: "モニターR", options: { fill: { color: "D9E2F3" } } }, { text: data.monitorRight }],
    ];
    slide.addTable(tableData, { x: 0.2, y: 0.2, w: 3.5, rowH: 0.3, fontSize: 10, fontFace: "Meiryo" });
  };

  const slide1 = pptx.addSlide();
  addHeader(slide1);
  slide1.addText("Before", { x: 0.2, y: 2.2, w: 4.7, h: 0.5, align: "center", fill: { color: "CCCCCC" }, fontFace: "Meiryo" });
  if (fetchedImages.before) slide1.addImage({ data: fetchedImages.before, x: 0.8, y: 2.6, w: 5.6, h: 4.2, sizing: { type: "contain", w: 5.6, h: 4.2 } });
  slide1.addText("After", { x: 5.1, y: 2.2, w: 4.7, h: 0.5, align: "center", fill: { color: "000099" }, color: "FFFFFF", fontFace: "Meiryo" });
  if (fetchedImages.after) slide1.addImage({ data: fetchedImages.after, x: 6.9, y: 2.6, w: 5.6, h: 4.2, sizing: { type: "contain", w: 5.6, h: 4.2 } });

  const slide2 = pptx.addSlide();
  addHeader(slide2);
  const views = [{ t: "正面", img: fetchedImages.front, x: 0.1 }, { t: "側面(左)", img: fetchedImages.sideLeft, x: 4.5 }, { t: "側面(右)", img: fetchedImages.sideRight, x: 8.9 }];
  views.forEach(v => {
    slide2.addText(v.t, { x: v.x, y: 2.5, w: 3.2, h: 0.6, align: "center", fill: { color: "000099" }, color: "FFFFFF", fontFace: "Meiryo" });
    if (v.img) slide2.addImage({ data: v.img, x: v.x + 0.1, y: 3.1, w: 4.2, h: 3.15, sizing: { type: "contain", w: 4.2, h: 3.15 } });
  });

  const slide3 = pptx.addSlide();
  addHeader(slide3);
  slide3.addText("他社比較", { x: 0.2, y: 2.2, w: 4.7, h: 0.5, align: "center", fill: { color: "CCCCCC" }, fontFace: "Meiryo" });
  if (fetchedImages.other1) slide3.addImage({ data: fetchedImages.other1, x: 0.8, y: 2.6, w: 5.6, h: 4.2, sizing: { type: "contain", w: 5.6, h: 4.2 } });
  slide3.addText("他社比較", { x: 5.1, y: 2.2, w: 4.7, h: 0.5, align: "center", fill: { color: "CCCCCC" }, fontFace: "Meiryo" });
  if (fetchedImages.other2) slide3.addImage({ data: fetchedImages.other2, x: 6.9, y: 2.6, w: 5.6, h: 4.2, sizing: { type: "contain", w: 5.6, h: 4.2 } });

  const fileName = `Report_${Date.now()}.pptx`;
  
  console.log("Saving with Base64 method...");
  try {
    const base64 = await pptx.write({ outputType: "base64" });
    const dataUrl = "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64," + base64;
    
    // Use saveAs with a data URI - very reliable
    saveAs(dataUrl, fileName);
    
    console.log("Success!");
    return true;
  } catch (err) {
    console.error("Save failed:", err);
    alert("保存に失敗しました: " + err);
    return false;
  }
}
