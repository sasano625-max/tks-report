import { NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const password = searchParams.get("password");
  const adminPassword = process.env.ADMIN_PASSWORD;
  const isPlanned = searchParams.get("planned") === "true";

  if (isPlanned) {
    const { data, error } = await supabase
      .from("reports")
      .select("*")
      .is("images", null)
      .order("completion_date", { ascending: true });
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });
    return NextResponse.json(data);
  }

  if (adminPassword && password !== adminPassword) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { data, error } = await supabase
    .from("reports")
    .select("*")
    .order("created_at", { ascending: false });

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  return NextResponse.json(data);
}

export async function POST(request: Request) {
  const body = await request.json();
  // Insert can be done by anyone (vendors)
  const { data, error } = await supabase.from("reports").insert(body).select();
  
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  return NextResponse.json(data);
}

export async function DELETE(request: Request) {
  const { searchParams } = new URL(request.url);
  const password = searchParams.get("password");
  const adminPassword = process.env.ADMIN_PASSWORD;

  if (adminPassword && password !== adminPassword) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const id = searchParams.get("id");
  if (id) {
    const { error } = await supabase.from("reports").delete().eq("id", id);
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  } else {
    // Delete all
    const { error } = await supabase.from("reports").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  }
  
  return NextResponse.json({ success: true });
}

export async function PATCH(request: Request) {
  const { searchParams } = new URL(request.url);
  const password = searchParams.get("password");
  const adminPassword = process.env.ADMIN_PASSWORD;

  if (adminPassword && password !== adminPassword) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await request.json();
  const id = searchParams.get("id");
  if (!id) return NextResponse.json({ error: "ID required" }, { status: 400 });

  const { error } = await supabase.from("reports").update(body).eq("id", id);
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  
  return NextResponse.json({ success: true });
}
