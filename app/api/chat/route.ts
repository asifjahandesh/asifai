import { NextRequest } from "next/server";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function jsonError(message: string, status = 400) {
  return Response.json({ error: message }, { status });
}

export async function POST(request: NextRequest) {
  const authHeader = request.headers.get("authorization");
  const token = authHeader?.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) return jsonError("You must sign in before chatting.", 401);

  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY;
  const gridexBase = process.env.GRIDEX_BASE_URL;
  const gridexKey = process.env.GRIDEX_API_KEY;
  const model = process.env.GRIDEX_MODEL || "smart";
  if (!supabaseUrl || !supabaseKey) return jsonError("Supabase environment variables are missing.", 500);
  if (!gridexBase || !gridexKey) return jsonError("Gridex server environment variables are missing.", 500);

  const supabase = createClient(supabaseUrl, supabaseKey, {
    global: { headers: { Authorization: `Bearer ${token}` } },
    auth: { persistSession: false, autoRefreshToken: false },
  });
  const { data: authData, error: authError } = await supabase.auth.getUser(token);
  if (authError || !authData.user) return jsonError("Your session is invalid. Please sign in again.", 401);

  let body: { message?: string; conversationId?: string | null };
  try { body = await request.json(); } catch { return jsonError("Invalid request body."); }
  const message = body.message?.trim();
  if (!message || message.length > 20000) return jsonError("Enter a message under 20,000 characters.");

  let conversationId = body.conversationId || null;
  if (conversationId) {
    const { data: existing, error } = await supabase.from("conversations").select("id").eq("id", conversationId).single();
    if (error || !existing) return jsonError("Conversation not found or access denied.", 404);
  } else {
    const title = message.slice(0, 72) || "New conversation";
    const { data, error } = await supabase.from("conversations").insert({ user_id: authData.user.id, title }).select("id").single();
    if (error || !data) return jsonError("Could not create conversation. Check the database schema and RLS policies.", 500);
    conversationId = data.id;
  }

  const { error: insertUserError } = await supabase.from("messages").insert({ conversation_id: conversationId, user_id: authData.user.id, role: "user", content: message });
  if (insertUserError) return jsonError("Could not save your message.", 500);

  const { data: history, error: historyError } = await supabase.from("messages").select("role, content").eq("conversation_id", conversationId).order("created_at", { ascending: true }).limit(80);
  if (historyError || !history) return jsonError("Could not load conversation history.", 500);

  let upstream: Response;
  try {
    upstream = await fetch(`${gridexBase.replace(/\/$/, "")}/chat/completions`, {
      method: "POST",
      headers: { Authorization: `Bearer ${gridexKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({ model, messages: history, stream: false }),
      cache: "no-store",
    });
  } catch {
    return jsonError("Could not reach Gridex. Check the API endpoint and network.", 502);
  }
  if (!upstream.ok) {
    const detail = await upstream.text().catch(() => "");
    return jsonError(`Gridex request failed (${upstream.status}): ${detail.slice(0, 500)}`, 502);
  }

  let completion: any;
  try { completion = await upstream.json(); }
  catch { return jsonError("Gridex returned an unreadable response.", 502); }
  const answer = completion?.choices?.[0]?.message?.content;
  if (typeof answer !== "string" || !answer.trim()) return jsonError("Gridex returned an empty response.", 502);

  const { error: insertAssistantError } = await supabase.from("messages").insert({ conversation_id: conversationId, user_id: authData.user.id, role: "assistant", content: answer });
  if (insertAssistantError) return jsonError("The response arrived, but could not be saved to chat history.", 500);

  await supabase.from("conversations").update({ updated_at: new Date().toISOString() }).eq("id", conversationId);
  return Response.json({ conversationId, content: answer });
}
