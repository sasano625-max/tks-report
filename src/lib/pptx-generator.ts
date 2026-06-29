import pptxgen from "pptxgenjs";
import { saveAs } from "file-saver";

export interface ReportData {
  storeName: string;
  completionDate: string;
  type: string;
  wifiNumber?: string;
  monitorPrefix?: string;
  monitorLeft: string;
  monitorRight: string;
  images: {
    before?: string;
    after?: string;
    front?: string;
    sideLeft?: string;
    sideRight?: string;
    storeFront?: string;
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
  const keys = Object.keys(data.images).filter(k => k !== "wifi_number");
  for (const key of keys) {
    const url = (data.images as any)[key];
    if (url) fetchedImages[key] = await getBase64FromUrl(url);
  }

  pptx.layout = "LAYOUT_WIDE";

  const addHeader = (slide: any, slideNumber: number) => {
    // Page Number
    slide.addText(slideNumber.toString(), { x: 12.8, y: 7.15, w: 0.4, h: 0.3, align: "right", fontSize: 12, fontFace: "Meiryo" });

    // Status Table
    const headerOpts = { fill: { color: "C5D9F1" }, color: "000000", align: "center", fontFace: "Meiryo", border: { type: "solid", color: "000000", pt: 1 } };
    const valueOpts = { fill: { color: "FFFFFF" }, color: "000000", align: "center", fontFace: "Meiryo", border: { type: "solid", color: "000000", pt: 1 } };

    const tableData = [
      [{ text: "店舗名", options: headerOpts as any }, { text: data.storeName || "", options: { ...valueOpts, fontSize: 8 } as any }],
      [{ text: "設置日", options: headerOpts as any }, { text: data.completionDate || "", options: valueOpts as any }],
      [{ text: "Wi-Fi番号", options: headerOpts as any }, { text: data.wifiNumber || "", options: valueOpts as any }],
      [{ text: "モニター番号", options: { ...headerOpts, colspan: 2 } as any }],
      [{ text: "左", options: headerOpts as any }, { text: "右", options: headerOpts as any }],
      [{ text: data.monitorLeft || "", options: valueOpts as any }, { text: data.monitorRight || "", options: valueOpts as any }],
    ];
    slide.addTable(tableData, { x: 0.3, y: 0.5, w: 4.0, colW: [2.0, 2.0], rowH: 0.25, fontSize: 9, fontFace: "Meiryo" });
  };

  // Slide 1: Before / After
  const slide1 = pptx.addSlide();
  addHeader(slide1, 1);
  slide1.addText("Before (設置前)", { x: 0.5, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 18 });
  if (fetchedImages.before) slide1.addImage({ data: fetchedImages.before, x: 0.5, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });
  
  slide1.addText("After (設置後)", { x: 6.8, y: 2.3, w: 6.0, h: 0.4, fill: { color: "000099" }, color: "FFFFFF", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 18 });
  if (fetchedImages.after) slide1.addImage({ data: fetchedImages.after, x: 6.8, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });

  // Slide 2: Front, SideLeft, SideRight
  const slide2 = pptx.addSlide();
  addHeader(slide2, 2);
  const views = [
    { t: "正面", img: fetchedImages.front, x: 0.3 }, 
    { t: "側面(左)", img: fetchedImages.sideLeft, x: 4.6 }, 
    { t: "側面(右)", img: fetchedImages.sideRight, x: 8.9 }
  ];
  views.forEach(v => {
    slide2.addText(v.t, { x: v.x, y: 3.0, w: 4.1, h: 0.4, align: "center", fill: { color: "000099" }, color: "FFFFFF", bold: true, shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 20 });
    if (v.img) slide2.addImage({ data: v.img, x: v.x, y: 3.4, w: 4.1, h: 3.5, sizing: { type: "contain", w: 4.1, h: 3.5 } });
  });

  // Slide 3: Store Front
  const slide3 = pptx.addSlide();
  addHeader(slide3, 3);
  slide3.addText("店舗正面写真", { x: 4.8, y: 1.2, w: 8.0, h: 0.4, fill: { color: "000099" }, color: "FFFFFF", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 20 });
  if (fetchedImages.storeFront) slide3.addImage({ data: fetchedImages.storeFront, x: 4.8, y: 1.6, w: 8.0, h: 5.3, sizing: { type: "contain", w: 8.0, h: 5.3 } });

  // Slide 4: Other Brands
  const slide4 = pptx.addSlide();
  addHeader(slide4, 4);
  slide4.addText("他社ブランドテーブル", { x: 0.5, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 18 });
  if (fetchedImages.other1) slide4.addImage({ data: fetchedImages.other1, x: 0.5, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });
  
  slide4.addText("他社ブランドテーブル", { x: 6.8, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: "Meiryo", fontSize: 18 });
  if (fetchedImages.other2) slide4.addImage({ data: fetchedImages.other2, x: 6.8, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });

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
