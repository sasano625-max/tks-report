import { NextResponse } from "next/server";

export async function POST(request: Request) {
  try {
    const { password } = await request.json();
    const adminPassword = process.env.ADMIN_PASSWORD;

    if (!adminPassword) {
      // 開発環境などで設定されていない場合のフォールバック（本番では必ず設定する）
      if (password === "tkstks_local") {
         return NextResponse.json({ success: true });
      }
      console.error("ADMIN_PASSWORD environment variable is not set.");
      return NextResponse.json({ success: false, error: "Server configuration error" }, { status: 500 });
    }

    if (password === adminPassword) {
      return NextResponse.json({ success: true });
    }
    
    return NextResponse.json({ success: false }, { status: 401 });
  } catch (error) {
    return NextResponse.json({ success: false }, { status: 500 });
  }
}
