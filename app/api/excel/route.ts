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
    mode?: "agent" | "chat";
    workbookContext?: {
      worksheets?: Array<{
        name: string;
        rowCount: number;
        columnCount: number;
        address: string;
        headers: string[];
        sampleRows?: any[][];
        inferredTypes?: Record<string, string>;
      }>;
      activeWorksheet?: string;
    };
    excelContext?: {
      worksheet?: string;
      address?: string;
      rowCount?: number;
      columnCount?: number;
      headers?: string[];
      preview?: any[][];
      formulas?: any[][];
      action?: "analyze" | "explain_formula" | "generate_formula" | "general" | "agent";
    };
  };

  try {
    body = await request.json();
  } catch {
    return jsonError("Invalid request body.");
  }

  const userMessage = body.message?.trim() || "";
  const excelContext = body.excelContext;
  const workbookContext = body.workbookContext;
  const isAgentMode = body.mode === "agent" || excelContext?.action === "agent";

  if (!userMessage && !excelContext && !workbookContext) {
    return jsonError("No message or Excel context provided.");
  }

  // Base system instructions with specialized Domain Intelligence (Textile, RMG, Industrial Planning & Automation)
  let systemPrompt = `You are ASIF AI, an elite Autonomous Microsoft Excel Automation Agent and industrial operations expert.
You assist users with Microsoft Excel data analysis, advanced modeling, formula generation, and automated workbook creation.

### DOMAIN EXPERTISE (Textile, RMG, Manufacturing, Operations & Planning)
Recognize common industry fields and KPIs:
- Buyer, Order No / PO, Style, Item, Fabric, Color, Season
- Order Qty, Cut Qty, Sew Qty, Pack Qty, Shipped Qty, Variance, Efficiency %, Wastage %
- Delivery Date, Ex-Factory Date, Shipment Date, Lead Time, Status, Risk Level
- Yarn Consumption, Fabric Consumption, Inventory Stock, Material Shortages

${isAgentMode ? `### AGENT MODE INSTRUCTIONS:
The user wants you to PLAN AND EXECUTE concrete actions inside their Excel workbook.
You must analyze their instruction against the workbook structure and return BOTH an explanation AND an actionable, structured execution plan in JSON.

Output format requirement:
1. Provide a concise, professional explanation of your understanding and planned actions.
2. If actions should be executed in Excel, provide a JSON block enclosed strictly between [ACTION_PLAN] and [/ACTION_PLAN]:
[ACTION_PLAN]
{
  "summary": "Brief description of the action plan",
  "requiresConfirmation": boolean (true if overwriting or creating heavy reports, false for simple non-destructive additions),
  "actions": [
    {
      "type": "create_worksheet" | "write_range" | "format_range" | "create_chart" | "create_table" | "autofit" | "highlight_variance",
      "worksheet": "Target Worksheet Name",
      "address": "A1:E10",
      "values": [[...]],
      "formulas": [[...]],
      "header": string,
      "format": {
        "bold": boolean,
        "fontSize": number,
        "fontColor": "#hex",
        "fillColor": "#hex",
        "horizontalAlignment": "Center" | "Left" | "Right",
        "numberFormat": "0.00%" | "$#,##0" | "#,##0" | "yyyy-mm-dd"
      },
      "chart": {
        "chartType": "ColumnClustered" | "Line" | "Pie" | "BarClustered",
        "dataRange": "A4:D14",
        "title": "Chart Title",
        "position": "G4",
        "width": 480,
        "height": 280
      }
    }
  ]
}
[/ACTION_PLAN]

Rules for Actions:
- For "create_worksheet", specify "worksheet": "New Sheet Name".
- For "write_range", CRITICAL DIMENSION RULE: The 2D array of "values" or "formulas" MUST match the exact row and column dimensions of the target "address".
  * For example, if "address": "A1:B1", the array MUST be exactly 1 row by 2 columns: [["Buyer", "Plan Qty"]].
  * If "address": "A2:B2", the array MUST have 1 row and 2 columns: [["H&M", 45000]].
  * If writing a single cell e.g. "B2", use "address": "B2" and an array of [[ "=SUM(...)" ]].
  * NEVER write a 1-column array like [["H&M"]] into a multi-column range like "A2:B2". If "values" has 1 column, use single-column address "A2:A2" or "A2".
- For KPIs, write KPI cards with clear metric names and calculated sums or ratios.
- For summary reports, write clean tables with titles, bold headers, dark or styled header fills, proper number formatting, and total rows.
- When generating reports from actual source worksheets (e.g. Buyer Plan Qty), use the source sheet's actual buyer names and quantities from the sample data provided, or use valid Excel formulas referencing the source sheet (e.g. "=SUMIF('Order Sheet'!A:A, A2, 'Order Sheet'!F:F)").
- For "create_chart", use valid Excel chartTypes: "ColumnClustered", "Line", "Pie", "BarClustered".
` : `### CHAT & FORMULA ASSISTANT MODE:
1. Provide concise, highly accurate, professional responses.
2. When asked for an Excel formula, always provide the exact formula clearly formatted in a markdown code block, and follow it with an EXCEL_FORMULA block:
[FORMULA]=YOUR_FORMULA_HERE[/FORMULA]
so the Excel Add-in can automatically detect and insert it upon user approval.
3. When explaining formulas, break down the component functions (e.g., INDEX, MATCH, XLOOKUP, SUMIFS) and describe their logic and purpose simply.
4. When analyzing data tables, highlight anomalies, statistical trends, outliers, or industrial KPIs.`}
`;

  let promptContent = userMessage;

  // Append full workbook context if provided
  if (workbookContext && workbookContext.worksheets && workbookContext.worksheets.length > 0) {
    const sheetSummaries = workbookContext.worksheets.map(ws => {
      let desc = `Worksheet "${ws.name}": Range ${ws.address} (${ws.rowCount} rows × ${ws.columnCount} cols)`;
      if (ws.headers && ws.headers.length > 0) {
        desc += `\n  Headers: [${ws.headers.join(", ")}]`;
      }
      if (ws.inferredTypes && Object.keys(ws.inferredTypes).length > 0) {
        desc += `\n  Detected Column Types: ${JSON.stringify(ws.inferredTypes)}`;
      }
      if (ws.sampleRows && ws.sampleRows.length > 0) {
        desc += `\n  Sample Data (first ${ws.sampleRows.length} rows):\n  ${JSON.stringify(ws.sampleRows.slice(0, 8))}`;
      }
      return desc;
    }).join("\n\n");

    promptContent = `[WORKBOOK CONTEXT]\nActive Worksheet: "${workbookContext.activeWorksheet || "Unknown"}"\n\n${sheetSummaries}\n\n[USER INSTRUCTION]\n${userMessage}`;
  } else if (excelContext) {
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

    promptContent = `[EXCEL SELECTION CONTEXT]\n${contextDetails.join("\n")}\n\n[USER QUERY]\n${userMessage || "Please analyze this selection."}`;
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
        const title = (userMessage || "Excel Automation").slice(0, 72);
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

  // Detect Action Plan if present
  let actionPlan = null;
  const actionPlanMatch = answer.match(/\[ACTION_PLAN\]([\s\S]*?)\[\/ACTION_PLAN\]/);
  if (actionPlanMatch && actionPlanMatch[1]) {
    try {
      actionPlan = JSON.parse(actionPlanMatch[1].trim());
    } catch {
      // Non-fatal if model outputted malformed JSON
    }
  }

  // Detect generated formula if present
  let detectedFormula: string | null = null;
  const formulaMatch = answer.match(/\[FORMULA\]([\s\S]*?)\[\/FORMULA\]/);
  if (formulaMatch && formulaMatch[1]) {
    detectedFormula = formulaMatch[1].trim();
  } else {
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

  // Clean the human-readable explanation by stripping raw [ACTION_PLAN] tag
  const cleanedContent = answer.replace(/\[ACTION_PLAN\][\s\S]*?\[\/ACTION_PLAN\]/, "").trim();

  return jsonResponse({
    conversationId,
    content: cleanedContent || answer,
    actionPlan,
    detectedFormula,
  });
}
