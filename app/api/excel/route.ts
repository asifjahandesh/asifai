import { NextRequest } from "next/server";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function jsonResponse(data: any, status = 200) {
  return Response.json(data, {
    status,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    },
  });
}

function jsonError(message: string, status = 400) {
  return jsonResponse({ error: message }, status);
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    },
  });
}

export async function POST(request: NextRequest) {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY;
  const gridexBase = process.env.GRIDEX_BASE_URL;
  const gridexKey = process.env.GRIDEX_API_KEY;
  const model = process.env.GRIDEX_MODEL || "smart";

  if (!gridexBase || !gridexKey) {
    return jsonError("Gridex server environment variables are missing on the backend.", 500);
  }

  // Parse authorization token if provided
  const authHeader = request.headers.get("authorization");
  const token = authHeader?.startsWith("Bearer ") ? authHeader.slice(7) : null;

  let authenticatedUser: { id: string; email?: string } | null = null;
  let supabaseClient: any = null;

  if (token && supabaseUrl && supabaseKey) {
    try {
      supabaseClient = createClient(supabaseUrl, supabaseKey, {
        global: { headers: { Authorization: `Bearer ${token}` } },
        auth: { persistSession: false, autoRefreshToken: false },
      });
      const { data: authData } = await supabaseClient.auth.getUser(token);
      if (authData?.user) {
        authenticatedUser = { id: authData.user.id, email: authData.user.email };
      }
    } catch {
      // Continue without user persistence if token is expired/invalid
    }
  }

  let body: {
    message?: string;
    conversationId?: string | null;
    excelContext?: {
      worksheet?: string;
      address?: string;
      rowCount?: number;
      columnCount?: number;
      headers?: string[];
      preview?: any[][];
      formulas?: any[][];
      action?: "analyze" | "explain_formula" | "generate_formula" | "general";
    };
  };

  try {
    body = await request.json();
  } catch {
    return jsonError("Invalid request body.");
  }

  const userMessage = body.message?.trim() || "";
  const excelContext = body.excelContext;

  if (!userMessage && !excelContext) {
    return jsonError("No message or Excel context provided.");
  }

  // Construct structured AI prompt incorporating Excel context
  let systemPrompt = `You are ASIF AI, an elite personal intelligence system and expert Microsoft Excel consultant.
You assist users with Microsoft Excel data analysis, advanced modeling, and formula generation.
Rules:
1. Provide concise, highly accurate, professional responses.
2. When asked for an Excel formula, always provide the exact formula clearly formatted in a markdown code block, and follow it with an EXCEL_FORMULA block:
[FORMULA]=YOUR_FORMULA_HERE[/FORMULA]
so the Excel Add-in can automatically detect and insert it upon user approval.
3. When explaining formulas, break down the component functions (e.g., INDEX, MATCH, XLOOKUP, SUMIFS) and describe their logic and purpose simply.
4. When analyzing data tables, highlight anomalies, statistical trends, outliers, or industrial KPIs.`;

  let promptContent = userMessage;

  if (excelContext) {
    const contextDetails: string[] = [];
    if (excelContext.worksheet) contextDetails.push(`Worksheet: "${excelContext.worksheet}"`);
    if (excelContext.address) contextDetails.push(`Selected Range: ${excelContext.address}`);
    if (excelContext.rowCount && excelContext.columnCount) {
      contextDetails.push(`Dimensions: ${excelContext.rowCount} rows × ${excelContext.columnCount} columns`);
    }
    if (excelContext.headers && excelContext.headers.length > 0) {
      contextDetails.push(`Headers: [${excelContext.headers.join(", ")}]`);
    }
    if (excelContext.formulas && excelContext.formulas.length > 0) {
      const flatFormulas = excelContext.formulas.flat().filter(f => typeof f === "string" && f.startsWith("="));
      if (flatFormulas.length > 0) {
        contextDetails.push(`Existing Formula(s): ${flatFormulas.slice(0, 5).join("; ")}`);
      }
    }
    if (excelContext.preview && excelContext.preview.length > 0) {
      contextDetails.push(`Selected Data Values (JSON matrix):\n${JSON.stringify(excelContext.preview.slice(0, 50))}`);
    }

    promptContent = `[EXCEL CONTEXT]\n${contextDetails.join("\n")}\n\n[USER QUERY]\n${userMessage || "Please analyze this selection."}`;
  }

  let conversationId = body.conversationId || null;

  // Persist conversation and message to Supabase if user is signed in
  if (authenticatedUser && supabaseClient) {
    try {
      if (conversationId) {
        const { data: existing } = await supabaseClient.from("conversations").select("id").eq("id", conversationId).single();
        if (!existing) conversationId = null;
      }
      if (!conversationId) {
        const title = (userMessage || "Excel Analysis").slice(0, 72);
        const { data: convData } = await supabaseClient.from("conversations").insert({ user_id: authenticatedUser.id, title }).select("id").single();
        if (convData) conversationId = convData.id;
      }
      if (conversationId) {
        await supabaseClient.from("messages").insert({
          conversation_id: conversationId,
          user_id: authenticatedUser.id,
          role: "user",
          content: promptContent,
        });
      }
    } catch {
      // Non-fatal if Supabase write fails in Excel context
    }
  }

  // Call Gridex API
  let messagesPayload = [
    { role: "system", content: systemPrompt },
    { role: "user", content: promptContent },
  ];

  let upstream: Response;
  try {
    upstream = await fetch(`${gridexBase.replace(/\/$/, "")}/chat/completions`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${gridexKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ model, messages: messagesPayload, stream: false }),
      cache: "no-store",
    });
  } catch (err: any) {
    return jsonError(`Could not reach Gridex AI backend: ${err?.message || "Network error"}`, 502);
  }

  if (!upstream.ok) {
    const detail = await upstream.text().catch(() => "");
    return jsonError(`Gridex AI request failed (${upstream.status}): ${detail.slice(0, 500)}`, 502);
  }

  let completion: any;
  try {
    completion = await upstream.json();
  } catch {
    return jsonError("Gridex returned an unreadable response.", 502);
  }

  const answer = completion?.choices?.[0]?.message?.content;
  if (typeof answer !== "string" || !answer.trim()) {
    return jsonError("Gridex returned an empty response.", 502);
  }

  // Detect generated formula if present
  let detectedFormula: string | null = null;
  const formulaMatch = answer.match(/\[FORMULA\]([\s\S]*?)\[\/FORMULA\]/);
  if (formulaMatch && formulaMatch[1]) {
    detectedFormula = formulaMatch[1].trim();
  } else {
    // Check for inline formula pattern like `=SUM(...)` or ````excel\n=...```
    const codeBlockMatch = answer.match(/```(?:excel|plaintext)?\s*\n(=[^\n]+)\n```/);
    if (codeBlockMatch && codeBlockMatch[1]) {
      detectedFormula = codeBlockMatch[1].trim();
    }
  }

  // Save assistant message to Supabase if authenticated
  if (authenticatedUser && supabaseClient && conversationId) {
    try {
      await supabaseClient.from("messages").insert({
        conversation_id: conversationId,
        user_id: authenticatedUser.id,
        role: "assistant",
        content: answer,
      });
      await supabaseClient.from("conversations").update({ updated_at: new Date().toISOString() }).eq("id", conversationId);
    } catch {}
  }

  return jsonResponse({
    conversationId,
    content: answer,
    detectedFormula,
  });
}
