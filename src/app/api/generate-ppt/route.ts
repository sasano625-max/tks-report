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
      tableX: 0.2, tableY: 0.35, tableW1: 1.0, tableW2: 3.2, tableRowH: 0.22, tableFontSize: 9,
      beforeX: 1.1, beforeY: 2.3, beforeW: 4.2, beforeH: 5.0,
      afterX: 8.0, afterY: 2.3, afterW: 4.2, afterH: 5.0,
      frontX: 0.2, frontY: 2.3, frontW: 4.2, frontH: 5.0,
      sideLeftX: 4.55, sideLeftY: 2.3, sideLeftW: 4.2, sideLeftH: 5.0,
      sideRightX: 8.9, sideRightY: 2.3, sideRightW: 4.2, sideRightH: 5.0,
      other1X: 1.1, other1Y: 2.3, other1W: 4.2, other1H: 5.0,
      other2X: 8.0, other2Y: 2.3, other2W: 4.2, other2H: 5.0,
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
        monitorLeft: rawData.monitor_left || rawData.monitorLeft || "",
        monitorRight: rawData.monitor_right || rawData.monitorRight || "",
        images: rawData.images || {}
      };

      log(`Adding slides for: ${data.storeName}`);
      
      const images = data.images || {};
      const imageKeys = Object.keys(images);
      const imagePromises = imageKeys.map(async (key) => {
        const b64 = await getBase64(images[key]);
        return { key, b64 };
      });
      
      const fetchedResults = await Promise.all(imagePromises);
      const fetchedImages: any = {};
      fetchedResults.forEach(res => {
        if (res.b64) fetchedImages[res.key] = res.b64;
      });

      const addHeader = (slide: any, title: string) => {
        // Main Title
        slide.addText(title, { 
          x: 4.0, y: 0.45, w: 9.0, h: 0.5, 
          fontSize: 24, fontFace: FONT_FACE, bold: true, color: PRIMARY_BLUE, 
          align: "right" 
        });

        const tableData = [
          [{ text: "店舗名", options: { fill: { color: "1E40AF" }, color: "FFFFFF", align: "center", bold: true, fontFace: FONT_FACE } }, { text: data.storeName, options: { align: "center", fontFace: FONT_FACE, fill: { color: "F1F5F9" } } }],
          [{ text: "設置日", options: { fill: { color: "1E40AF" }, color: "FFFFFF", align: "center", bold: true, fontFace: FONT_FACE } }, { text: data.completionDate, options: { align: "center", fontFace: FONT_FACE, fill: { color: "F1F5F9" } } }],
          [{ text: "タイプ", options: { fill: { color: "1E40AF" }, color: "FFFFFF", align: "center", bold: true, fontFace: FONT_FACE } }, { text: data.type, options: { align: "center", fontFace: FONT_FACE, fill: { color: "F1F5F9" } } }],
          [{ text: "モニターL", options: { fill: { color: "1E40AF" }, color: "FFFFFF", align: "center", bold: true, fontFace: FONT_FACE } }, { text: data.monitorLeft, options: { align: "center", fontFace: FONT_FACE, fill: { color: "F1F5F9" } } }],
          [{ text: "モニターR", options: { fill: { color: "1E40AF" }, color: "FFFFFF", align: "center", bold: true, fontFace: FONT_FACE } }, { text: data.monitorRight, options: { align: "center", fontFace: FONT_FACE, fill: { color: "F1F5F9" } } }],
        ];
        slide.addTable(tableData, { 
          x: layout.tableX, y: layout.tableY, colW: [layout.tableW1, layout.tableW2], rowH: layout.tableRowH, 
          fontSize: layout.tableFontSize, fontFace: FONT_FACE, 
          border: { type: "solid", color: "CBD5E1", pt: 0.5 } 
        });
      };

      // Slide 1: Before & After
      const slide1 = pptx.addSlide();
      if (templateBase64) slide1.background = { data: templateBase64 };
      addHeader(slide1, "設置前・設置後 比較報告");
      
      slide1.addText("BEFORE", { x: layout.beforeX, y: layout.beforeY - 0.35, w: layout.beforeW, h: 0.3, align: "center", fill: { color: "64748B" }, color: "FFFFFF", fontFace: FONT_FACE, bold: true, fontSize: 12 });
      if (fetchedImages.before) slide1.addImage({ data: fetchedImages.before, x: layout.beforeX, y: layout.beforeY, w: layout.beforeW, h: layout.beforeH, sizing: { type: "contain", w: layout.beforeW, h: layout.beforeH } });
      
      slide1.addText("AFTER", { x: layout.afterX, y: layout.afterY - 0.35, w: layout.afterW, h: 0.3, align: "center", fill: { color: PRIMARY_BLUE }, color: "FFFFFF", fontFace: FONT_FACE, bold: true, fontSize: 12 });
      if (fetchedImages.after) slide1.addImage({ data: fetchedImages.after, x: layout.afterX, y: layout.afterY, w: layout.afterW, h: layout.afterH, sizing: { type: "contain", w: layout.afterW, h: layout.afterH } });

      // Slide 2: 3-Way Views
      const slide2 = pptx.addSlide();
      if (templateBase64) slide2.background = { data: templateBase64 };
      addHeader(slide2, "設置詳細（三面写真）");
      
      const views = [
        { t: "正面 (Front)", img: fetchedImages.front, x: layout.frontX, y: layout.frontY, w: layout.frontW, h: layout.frontH }, 
        { t: "側面 左 (Left)", img: fetchedImages.sideLeft, x: layout.sideLeftX, y: layout.sideLeftY, w: layout.sideLeftW, h: layout.sideLeftH }, 
        { t: "側面 右 (Right)", img: fetchedImages.sideRight, x: layout.sideRightX, y: layout.sideRightY, w: layout.sideRightW, h: layout.sideRightH }
      ];
      views.forEach(v => {
        slide2.addText(v.t, { x: v.x, y: v.y - 0.35, w: v.w, h: 0.3, align: "center", fill: { color: PRIMARY_BLUE }, color: "FFFFFF", fontFace: FONT_FACE, bold: true, fontSize: 11 });
        if (v.img) slide2.addImage({ data: v.img, x: v.x, y: v.y, w: v.w, h: v.h, sizing: { type: "contain", w: v.w, h: v.h } });
      });

      // Slide 3: Other Brands
      const slide3 = pptx.addSlide();
      if (templateBase64) slide3.background = { data: templateBase64 };
      addHeader(slide3, "他社状況・周辺比較");
      
      slide3.addText("他社比較 1", { x: layout.other1X, y: layout.other1Y - 0.35, w: layout.other1W, h: 0.3, align: "center", fill: { color: "64748B" }, color: "FFFFFF", fontFace: FONT_FACE, bold: true, fontSize: 12 });
      if (fetchedImages.other1) slide3.addImage({ data: fetchedImages.other1, x: layout.other1X, y: layout.other1Y, w: layout.other1W, h: layout.other1H, sizing: { type: "contain", w: layout.other1W, h: layout.other1H } });
      
      slide3.addText("他社比較 2", { x: layout.other2X, y: layout.other2Y - 0.35, w: layout.other2W, h: 0.3, align: "center", fill: { color: "64748B" }, color: "FFFFFF", fontFace: FONT_FACE, bold: true, fontSize: 12 });
      if (fetchedImages.other2) slide3.addImage({ data: fetchedImages.other2, x: layout.other2X, y: layout.other2Y, w: layout.other2W, h: layout.other2H, sizing: { type: "contain", w: layout.other2W, h: layout.other2H } });
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


