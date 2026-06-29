import { NextRequest, NextResponse } from "next/server";
import pptxgen from "pptxgenjs";
import fs from "fs";
import path from "path";

export async function POST(req: NextRequest) {
  const logFile = path.join(process.cwd(), "server_log.txt");
  const log = (msg: string) => {
    const entry = `[${new Date().toISOString()}] ${msg}\n`;
    console.log(entry.trim());
    try {
      fs.appendFileSync(logFile, entry);
    } catch (e) {
      // Ignore log write errors
    }
  };

  log("API called");
  
  try {
    const body = await req.json();
    
    // Handle both single report and batch reports
    let reports: any[] = [];
    let templateUrl: string | null = null;

    if (body.reports && Array.isArray(body.reports)) {
      reports = body.reports;
      templateUrl = body.templateUrl;
    } else {
      reports = [body];
      templateUrl = body.templateUrl;
    }

    log(`Generating PPT for ${reports.length} reports.`);
    
    const defaultLayout = {
      tableX: 0.2, tableY: 0.35, tableW1: 2.5, tableW2: 2.5, tableRowH: 0.25, tableFontSize: 9,
      beforeX: 0.8, beforeY: 2.6, beforeW: 5.6, beforeH: 4.2,
      afterX: 6.9, afterY: 2.6, afterW: 5.6, afterH: 4.2,
      frontX: 0.2, frontY: 3.1, frontW: 4.2, frontH: 3.15,
      sideLeftX: 4.55, sideLeftY: 3.1, sideLeftW: 4.2, sideLeftH: 3.15,
      sideRightX: 8.9, sideRightY: 3.1, sideRightW: 4.2, sideRightH: 3.15,
      other1X: 0.8, other1Y: 2.6, other1W: 5.6, other1H: 4.2,
      other2X: 6.9, other2Y: 2.6, other2W: 5.6, other2H: 4.2,
    };
    const layout = body.layoutConfig || defaultLayout;

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";

    // Helper for images (Server side fetch)
    const getBase64 = async (url: string) => {
      try {
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), 15000); // 15s timeout
        const resp = await fetch(url, { signal: controller.signal });
        clearTimeout(tid);
        
        if (!resp.ok) {
          log(`Failed to fetch image: ${url} (Status: ${resp.status})`);
          return null;
        }

        const buffer = await resp.arrayBuffer();
        const base64 = Buffer.from(buffer).toString("base64");
        const contentType = resp.headers.get("content-type") || "image/jpeg";
        return `data:${contentType};base64,${base64}`;
      } catch (e) {
        log(`Image fetch error: ${url} - ${e}`);
        return null;
      }
    };

    // Fetch Template Background once for all slides
    let templateBase64: string | null = null;
    if (templateUrl) {
      log(`Fetching template background: ${templateUrl}`);
      templateBase64 = await getBase64(templateUrl);
      log(`Template fetch result: ${templateBase64 ? "Success" : "Failed"}`);
    }

    const FONT_FACE = "Meiryo";
    const PRIMARY_BLUE = "1E40AF"; // Modern blue
    const SECONDARY_BLUE = "1D4ED8";
    const DARK_TEXT = "1E293B";
    const LIGHT_BG = "F8FAFC";

    // Process each report
    for (const rawData of reports) {
      const data = {
        storeName: rawData.store_name || rawData.storeName || "名称未設定",
        completionDate: rawData.completion_date || rawData.completionDate || "",
        type: rawData.type || "",
        wifiNumber: rawData.wifi_number || rawData.wifiNumber || rawData.images?.wifi_number || "",
        monitorLeft: rawData.monitor_left || rawData.monitorLeft || "",
        monitorRight: rawData.monitor_right || rawData.monitorRight || "",
        images: rawData.images || {}
      };

      log(`Adding slides for: ${data.storeName}`);
      
      const images = data.images || {};
      const imageKeys = Object.keys(images).filter(k => k !== "wifi_number");
      const imagePromises = imageKeys.map(async (key) => {
        const b64 = await getBase64(images[key]);
        return { key, b64 };
      });
      
      const fetchedResults = await Promise.all(imagePromises);
      const fetchedImages: any = {};
      fetchedResults.forEach(res => {
        if (res.b64) fetchedImages[res.key] = res.b64;
      });

      const addHeader = (slide: any, slideNumber: number) => {
        // Page Number
        slide.addText(slideNumber.toString(), { x: 12.8, y: 7.15, w: 0.4, h: 0.3, align: "right", fontSize: 12, fontFace: FONT_FACE });

        // Status Table
        const headerOpts = { fill: { color: "C5D9F1" }, color: "000000", align: "center", fontFace: FONT_FACE };
        const valueOpts = { fill: { color: "FFFFFF" }, color: "000000", align: "center", fontFace: FONT_FACE };

        const tableData = [
          [{ text: "店舗名", options: headerOpts as any }, { text: data.storeName || "", options: { ...valueOpts, fontSize: 8 } as any }],
          [{ text: "設置日", options: headerOpts as any }, { text: data.completionDate || "", options: valueOpts as any }],
          [{ text: "Wi-Fi番号", options: headerOpts as any }, { text: data.wifiNumber || "", options: valueOpts as any }],
          [{ text: "モニター番号", options: { ...headerOpts, colspan: 2 } as any }],
          [{ text: "左", options: headerOpts as any }, { text: "右", options: headerOpts as any }],
          [{ text: data.monitorLeft || "", options: valueOpts as any }, { text: data.monitorRight || "", options: valueOpts as any }],
        ];
        slide.addTable(tableData, { 
          x: 0.3, y: 0.5, colW: [2.0, 2.0], rowH: 0.25, 
          fontSize: 9, fontFace: FONT_FACE, 
          border: { type: "solid", color: "000000", pt: 1.0 } 
        });
      };

      // Slide 1: Before / After
      const slide1 = pptx.addSlide();
      if (templateBase64) slide1.background = { data: templateBase64 };
      addHeader(slide1, 1);
      
      slide1.addText("Before (設置前)", { x: 0.5, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 18 });
      if (fetchedImages.before) slide1.addImage({ data: fetchedImages.before, x: 0.5, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });
      
      slide1.addText("After (設置後)", { x: 6.8, y: 2.3, w: 6.0, h: 0.4, fill: { color: "000099" }, color: "FFFFFF", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 18 });
      if (fetchedImages.after) slide1.addImage({ data: fetchedImages.after, x: 6.8, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });

      // Slide 2: Front, SideLeft, SideRight
      const slide2 = pptx.addSlide();
      if (templateBase64) slide2.background = { data: templateBase64 };
      addHeader(slide2, 2);
      
      const views = [
        { t: "正面", img: fetchedImages.front, x: 0.3 }, 
        { t: "側面(左)", img: fetchedImages.sideLeft, x: 4.6 }, 
        { t: "側面(右)", img: fetchedImages.sideRight, x: 8.9 }
      ];
      views.forEach(v => {
        slide2.addText(v.t, { x: v.x, y: 3.0, w: 4.1, h: 0.4, align: "center", fill: { color: "000099" }, color: "FFFFFF", bold: true, shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 20 });
        if (v.img) slide2.addImage({ data: v.img, x: v.x, y: 3.4, w: 4.1, h: 3.5, sizing: { type: "contain", w: 4.1, h: 3.5 } });
      });

      // Slide 3: Store Front
      const slide3 = pptx.addSlide();
      if (templateBase64) slide3.background = { data: templateBase64 };
      addHeader(slide3, 3);
      
      slide3.addText("店舗正面写真", { x: 4.8, y: 1.2, w: 8.0, h: 0.4, fill: { color: "000099" }, color: "FFFFFF", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 20 });
      if (fetchedImages.storeFront) slide3.addImage({ data: fetchedImages.storeFront, x: 4.8, y: 1.6, w: 8.0, h: 5.3, sizing: { type: "contain", w: 8.0, h: 5.3 } });

      // Slide 4: Other Brands
      const slide4 = pptx.addSlide();
      if (templateBase64) slide4.background = { data: templateBase64 };
      addHeader(slide4, 4);
      
      slide4.addText("他社ブランドテーブル", { x: 0.5, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 18 });
      if (fetchedImages.other1) slide4.addImage({ data: fetchedImages.other1, x: 0.5, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });
      
      slide4.addText("他社ブランドテーブル", { x: 6.8, y: 2.3, w: 6.0, h: 0.4, fill: { color: "D9D9D9" }, color: "000000", bold: true, align: "center", shape: (pptx as any).shapes.ROUNDED_RECTANGLE, fontFace: FONT_FACE, fontSize: 18 });
      if (fetchedImages.other2) slide4.addImage({ data: fetchedImages.other2, x: 6.8, y: 2.7, w: 6.0, h: 4.2, sizing: { type: "contain", w: 6.0, h: 4.2 } });
    }

    const buffer = await pptx.write({ outputType: "nodebuffer" });
    log(`Batch PPT generated successfully. Size: ${(buffer as Buffer).length} bytes`);

    const filename = reports.length === 1 ? `${reports[0].storeName}.pptx` : `Batch_Report_${Date.now()}.pptx`;

    return new NextResponse(buffer as unknown as BodyInit, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        "Cache-Control": "no-cache"
      },
    });
  } catch (error: any) {
    log(`SERVER ERROR: ${error.message}\n${error.stack}`);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}


