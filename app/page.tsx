"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import Image from "next/image";
import { ArrowUp, Bot, LogOut, Menu, MessageSquarePlus, Moon, Plus, ShieldCheck, Sparkles, Sun, UserRound } from "lucide-react";
import type { Session, User } from "@supabase/supabase-js";
import { getSupabase } from "@/lib/supabase";

type Conversation = { id: string; title: string; updated_at: string };
type ChatMessage = { id?: string; role: "user" | "assistant"; content: string };
type Theme = "dark" | "light";

export default function Home() {
  const [session, setSession] = useState<Session | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authMode, setAuthMode] = useState<"signin" | "signup">("signin");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authError, setAuthError] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState("");
  const [sending, setSending] = useState(false);
  const [chatError, setChatError] = useState("");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [theme, setTheme] = useState<Theme>("dark");
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Sync theme with localStorage and root class
  useEffect(() => {
    try {
      const saved = localStorage.getItem("asif_theme") as Theme | null;
      const initialTheme = saved === "light" ? "light" : "dark";
      setTheme(initialTheme);
      applyTheme(initialTheme);
    } catch {
      applyTheme("dark");
    }
  }, []);

  const applyTheme = (nextTheme: Theme) => {
    if (typeof document !== "undefined") {
      if (nextTheme === "light") {
        document.documentElement.classList.add("light");
        document.documentElement.classList.remove("dark");
      } else {
        document.documentElement.classList.add("dark");
        document.documentElement.classList.remove("light");
      }
    }
  };

  const toggleTheme = () => {
    const nextTheme: Theme = theme === "dark" ? "light" : "dark";
    setTheme(nextTheme);
    applyTheme(nextTheme);
    try {
      localStorage.setItem("asif_theme", nextTheme);
    } catch {}
  };

  useEffect(() => {
    let alive = true;
    let subscription: { unsubscribe: () => void } | undefined;
    try {
      const supabase = getSupabase();
      supabase.auth.getSession().then(({ data }) => {
        if (!alive) return;
        setSession(data.session); setUser(data.session?.user ?? null); setAuthLoading(false);
      });
      const result = supabase.auth.onAuthStateChange((_event, nextSession) => {
        if (!alive) return;
        setSession(nextSession); setUser(nextSession?.user ?? null); setAuthLoading(false);
      });
      subscription = result.data.subscription;
    } catch (e) { setAuthError(e instanceof Error ? e.message : "Configuration error"); setAuthLoading(false); }
    return () => { alive = false; subscription?.unsubscribe(); };
  }, []);

  const loadConversations = useCallback(async () => {
    if (!user) return;
    const { data, error } = await getSupabase().from("conversations").select("id,title,updated_at").order("updated_at", { ascending: false }).limit(100);
    if (!error && data) setConversations(data as Conversation[]);
  }, [user]);

  const loadMessages = useCallback(async (id: string) => {
    const { data, error } = await getSupabase().from("messages").select("id,role,content").eq("conversation_id", id).order("created_at", { ascending: true }).limit(200);
    if (!error && data) setMessages(data as ChatMessage[]);
  }, []);

  useEffect(() => { if (user) loadConversations(); else { setConversations([]); setActiveId(null); setMessages([]); } }, [user, loadConversations]);
  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages]);

  async function authenticate(e: React.FormEvent) {
    e.preventDefault(); setAuthBusy(true); setAuthError("");
    try {
      const supabase = getSupabase();
      const result = authMode === "signup"
        ? await supabase.auth.signUp({ email, password })
        : await supabase.auth.signInWithPassword({ email, password });
      if (result.error) throw result.error;
      if (authMode === "signup" && !result.data.session) setAuthError("Check your email to confirm your account, then sign in.");
    } catch (e) { setAuthError(e instanceof Error ? e.message : "Authentication failed."); }
    finally { setAuthBusy(false); }
  }

  async function signOut() { await getSupabase().auth.signOut(); }

  async function sendMessage(text = input) {
    const prompt = text.trim();
    if (!prompt || !session || sending) return;
    setInput(""); setChatError(""); setSending(true);
    const optimistic: ChatMessage[] = [...messages, { role: "user", content: prompt }, { role: "assistant", content: "" }];
    setMessages(optimistic);
    try {
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session.access_token}` },
        body: JSON.stringify({ message: prompt, conversationId: activeId }),
      });
      const result = await response.json().catch(() => ({ error: "Chat request failed." }));
      if (!response.ok) throw new Error(result.error || "Chat request failed.");
      const newId = result.conversationId as string;
      setActiveId(newId);
      setMessages([...optimistic.slice(0, -1), { role: "assistant", content: result.content as string }]);
      await loadConversations();
      await loadMessages(newId);
    } catch (e) {
      setChatError(e instanceof Error ? e.message : "Something went wrong.");
      setMessages(optimistic.slice(0, -1));
    } finally { setSending(false); textareaRef.current?.focus(); }
  }

  async function newChat() { setActiveId(null); setMessages([]); setChatError(""); setInput(""); textareaRef.current?.focus(); }
  async function selectConversation(id: string) { setActiveId(id); setChatError(""); await loadMessages(id); }

  if (authLoading) return (
    <div className="min-h-screen grid place-items-center bg-[var(--bg-primary)] text-green-500 transition-colors">
      <div className="animate-pulse tracking-[.25em] text-xs">INITIALIZING ASIF AI...</div>
    </div>
  );

  if (!user) return (
    <main className="grid-bg min-h-screen flex items-center justify-center p-5 relative overflow-hidden transition-colors">
      <div className="absolute top-5 right-5 z-20">
        <button
          onClick={toggleTheme}
          title={theme === "dark" ? "Switch to Light Mode" : "Switch to Dark Mode"}
          aria-label="Toggle theme"
          className="p-2.5 rounded-xl border border-[var(--border-strong)] bg-[var(--bg-card)] text-[var(--text-muted)] hover:text-[var(--text-primary)] hover:border-[var(--accent-primary)] shadow-sm transition"
        >
          {theme === "dark" ? <Sun size={17} className="text-amber-400" /> : <Moon size={17} className="text-slate-700" />}
        </button>
      </div>

      <div className="absolute w-[520px] h-[520px] rounded-full bg-green-500/10 blur-[130px] -top-48 -right-40 pointer-events-none" />
      <div className="absolute w-[420px] h-[420px] rounded-full bg-emerald-700/10 blur-[120px] -bottom-52 -left-40 pointer-events-none" />
      
      <section className="relative w-full max-w-md border border-[var(--border-strong)] bg-[var(--bg-card)] rounded-3xl p-8 glow transition-colors">
        <BrandMark />
        <p className="mt-8 text-[10px] tracking-[.28em] text-[var(--accent-badge-text)] font-semibold">PRIVATE ACCESS / GRIDEX SMART</p>
        <h1 className="text-3xl font-semibold mt-2 text-[var(--text-heading)]">{authMode === "signin" ? "Welcome back." : "Create your access."}</h1>
        <p className="text-sm text-[var(--text-muted)] mt-2">Your private intelligence workspace.</p>
        <form onSubmit={authenticate} className="mt-7 space-y-4">
          <label className="block text-xs text-[var(--text-muted)]">
            EMAIL
            <input
              type="email"
              required
              value={email}
              onChange={e=>setEmail(e.target.value)}
              className="mt-2 w-full rounded-xl bg-[var(--bg-input)] border border-[var(--border-primary)] focus:border-[var(--accent-primary)] outline-none px-4 py-3 text-sm text-[var(--text-heading)] transition"
              placeholder="you@example.com"
            />
          </label>
          <label className="block text-xs text-[var(--text-muted)]">
            PASSWORD
            <input
              type="password"
              required
              minLength={8}
              value={password}
              onChange={e=>setPassword(e.target.value)}
              className="mt-2 w-full rounded-xl bg-[var(--bg-input)] border border-[var(--border-primary)] focus:border-[var(--accent-primary)] outline-none px-4 py-3 text-sm text-[var(--text-heading)] transition"
              placeholder="At least 8 characters"
            />
          </label>
          {authError && <p className="text-xs text-rose-500 dark:text-rose-300 bg-rose-500/10 dark:bg-rose-950/40 border border-rose-500/30 dark:border-rose-900/60 rounded-lg p-3">{authError}</p>}
          <button
            disabled={authBusy}
            className="w-full rounded-xl bg-[var(--accent-primary)] hover:bg-[var(--accent-primary-hover)] disabled:opacity-50 text-white dark:text-black font-semibold py-3 text-sm shadow-[0_0_24px_rgba(34,197,94,.18)] transition cursor-pointer"
          >
            {authBusy ? "Please wait..." : authMode === "signin" ? "Sign in" : "Create account"}
          </button>
        </form>
        <button
          onClick={()=>{setAuthMode(authMode === "signin" ? "signup" : "signin");setAuthError("");}}
          className="w-full text-xs text-[var(--text-muted)] hover:text-[var(--text-accent)] mt-5 transition"
        >
          {authMode === "signin" ? "New here? Create an account" : "Already registered? Sign in"}
        </button>
        <div className="mt-7 pt-5 border-t border-[var(--border-primary)] flex items-center justify-center gap-2 text-[10px] text-[var(--text-subtle)]">
          <ShieldCheck size={13}/> PRIVATE SESSION · ENCRYPTED TRANSPORT
        </div>
      </section>
    </main>
  );

  return (
    <main className="h-screen bg-[var(--bg-primary)] text-[var(--text-primary)] flex overflow-hidden transition-colors">
      <aside className={`${sidebarOpen ? "w-[280px]" : "w-0"} transition-all duration-200 shrink-0 border-r border-[var(--border-primary)] bg-[var(--bg-surface)] flex flex-col overflow-hidden`}>
        <div className="p-4 border-b border-[var(--border-primary)]"><BrandMark /></div>
        <div className="p-3">
          <button
            onClick={newChat}
            className="w-full flex items-center gap-2 rounded-xl border border-[var(--accent-badge-border)] bg-[var(--accent-badge-bg)] hover:opacity-90 px-3 py-3 text-sm text-[var(--accent-badge-text)] font-medium transition"
          >
            <MessageSquarePlus size={16}/> New conversation <Plus size={15} className="ml-auto"/>
          </button>
        </div>
        <div className="px-4 py-2 text-[10px] tracking-[.2em] text-[var(--text-subtle)]">YOUR HISTORY</div>
        <div className="flex-1 overflow-y-auto scrollbar-thin px-2 space-y-1">
          {conversations.map(c=>(
            <button
              key={c.id}
              onClick={()=>selectConversation(c.id)}
              className={`w-full text-left px-3 py-2.5 rounded-lg text-xs truncate flex items-center gap-2 transition ${
                activeId===c.id
                  ? "bg-[var(--accent-badge-bg)] text-[var(--accent-badge-text)] border border-[var(--accent-badge-border)]"
                  : "text-[var(--text-muted)] hover:bg-[var(--bg-hover)] hover:text-[var(--text-primary)]"
              }`}
            >
              <MessageSquarePlus size={13} className="shrink-0"/><span className="truncate">{c.title}</span>
            </button>
          ))}
          {!conversations.length && <p className="px-3 py-4 text-xs text-[var(--text-subtle)]">Your conversations will appear here.</p>}
        </div>
        <div className="p-3 border-t border-[var(--border-primary)]">
          <div className="flex items-center gap-2 px-2 py-2">
            <div className="w-8 h-8 rounded-full bg-[var(--user-avatar-bg)] border border-[var(--user-avatar-border)] grid place-items-center text-[var(--text-muted)]">
              <UserRound size={15}/>
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-xs text-[var(--text-primary)] font-medium truncate">{user.email}</p>
              <p className="text-[10px] text-[var(--accent-badge-text)]">PRIVATE ACCOUNT</p>
            </div>
            <button onClick={signOut} title="Sign out" className="p-2 text-[var(--text-muted)] hover:text-rose-500 transition">
              <LogOut size={15}/>
            </button>
          </div>
        </div>
      </aside>

      <section className="flex-1 min-w-0 flex flex-col grid-bg relative">
        <header className="h-16 shrink-0 border-b border-[var(--border-primary)] bg-[var(--bg-header)] backdrop-blur flex items-center px-4 md:px-6 gap-3 transition-colors">
          <button onClick={()=>setSidebarOpen(!sidebarOpen)} className="p-2 rounded-lg hover:bg-[var(--bg-hover)] text-[var(--text-muted)] hover:text-[var(--text-primary)] transition">
            <Menu size={18}/>
          </button>
          <div className="flex-1">
            <div className="text-sm font-medium text-[var(--text-primary)]">
              ASIF AI <span className="text-[var(--accent-primary)]">/</span> <span className="text-[var(--text-muted)] text-xs">Workspace</span>
            </div>
          </div>
          
          {/* Status Indicator */}
          <div className="flex items-center gap-2 text-[10px] text-[var(--accent-badge-text)] border border-[var(--accent-badge-border)] bg-[var(--accent-badge-bg)] rounded-full px-3 py-1.5 font-medium">
            <span className="w-1.5 h-1.5 rounded-full bg-[var(--accent-dot)] shadow-[0_0_8px_var(--accent-dot-glow)]"/> SMART ONLINE
          </div>

          {/* Theme Toggle Button in top-right workspace header */}
          <button
            onClick={toggleTheme}
            title={theme === "dark" ? "Switch to Light Mode" : "Switch to Dark Mode"}
            aria-label="Toggle theme"
            className="p-2 rounded-xl border border-[var(--border-primary)] bg-[var(--bg-card)] text-[var(--text-muted)] hover:text-[var(--text-primary)] hover:border-[var(--accent-primary)] hover:bg-[var(--bg-hover)] transition flex items-center justify-center cursor-pointer shadow-sm"
          >
            {theme === "dark" ? (
              <Sun size={17} className="text-amber-400" />
            ) : (
              <Moon size={17} className="text-slate-700" />
            )}
          </button>
        </header>

        <div className="flex-1 overflow-y-auto scrollbar-thin px-4 py-8 md:px-8">
          <div className="max-w-4xl mx-auto">
            {!messages.length ? (
              <div className="min-h-[55vh] flex flex-col justify-center items-center text-center">
                <div className="w-16 h-16 rounded-2xl bg-black border border-green-500/80 overflow-hidden grid place-items-center glow">
                  <Image src="/icon.svg" alt="ASIF AI" width={64} height={64} className="w-full h-full object-cover" priority />
                </div>
                <p className="mt-7 text-[10px] tracking-[.3em] text-[var(--accent-badge-text)] font-semibold">PERSONAL INTELLIGENCE SYSTEM</p>
                <h1 className="mt-3 text-4xl md:text-5xl font-semibold tracking-tight text-[var(--text-heading)] neon-text">
                  Think beyond limits<span className="text-[var(--accent-primary)]">.</span>
                </h1>
                <p className="mt-4 text-sm text-[var(--text-muted)] max-w-lg">
                  Your private AI workspace for engineering, data, programming and ideas. Powered by Gridex Smart.
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-10 w-full max-w-2xl">
                  {[
                    "Explain an Excel formula with an example",
                    "Help me analyze a production planning problem",
                    "Write a clean SQL query",
                    "Explain an industrial engineering KPI"
                  ].map((s,i)=>(
                    <button
                      key={s}
                      onClick={()=>sendMessage(s)}
                      className="text-left border border-[var(--border-primary)] bg-[var(--bg-card)] hover:border-[var(--accent-primary)] hover:bg-[var(--bg-card-subtle)] rounded-xl p-4 text-xs text-[var(--text-muted)] hover:text-[var(--text-primary)] transition shadow-sm"
                    >
                      <span className="text-[var(--accent-badge-text)] font-semibold block mb-2">0{i+1} / PROMPT</span>
                      {s}
                    </button>
                  ))}
                </div>
              </div>
            ) : (
              <div className="space-y-7 pb-8">
                {messages.map((m,i)=>(
                  <div key={m.id || i} className={`flex gap-3 ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                    {m.role === "assistant" && (
                      <div className="w-8 h-8 shrink-0 rounded-lg bg-[var(--brand-icon-bg)] border border-[var(--brand-icon-border)] grid place-items-center">
                        <Bot size={16} className="text-[var(--accent-primary)]"/>
                      </div>
                    )}
                    <div
                      className={`max-w-[85%] rounded-2xl px-4 py-3 text-sm prose-ai shadow-sm ${
                        m.role === "user"
                          ? "bg-[var(--bg-user-bubble)] border border-[var(--border-user-bubble)] text-[var(--text-primary)]"
                          : "bg-[var(--bg-assistant-bubble)] border border-[var(--border-assistant-bubble)] text-[var(--text-primary)]"
                      }`}
                    >
                      {m.content || (sending ? <span className="text-[var(--accent-primary)] animate-pulse">Thinking...</span> : "")}
                    </div>
                    {m.role === "user" && (
                      <div className="w-8 h-8 shrink-0 rounded-lg bg-[var(--user-avatar-bg)] border border-[var(--user-avatar-border)] grid place-items-center text-[var(--text-muted)]">
                        <UserRound size={15}/>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
            {chatError && <div className="max-w-3xl mx-auto mb-4 rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 text-xs text-rose-500 dark:text-rose-200">{chatError}</div>}
            <div ref={bottomRef}/>
          </div>
        </div>

        <div className="shrink-0 px-4 pb-5 md:px-8 md:pb-7">
          <div className="max-w-4xl mx-auto">
            <div className="rounded-2xl border border-[var(--border-strong)] bg-[var(--bg-card)] glow focus-within:border-[var(--accent-primary)] transition shadow-sm">
              <textarea
                ref={textareaRef}
                value={input}
                onChange={e=>setInput(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMessage();}}}
                rows={2}
                placeholder="Message ASIF AI..."
                className="w-full resize-none bg-transparent outline-none px-5 pt-4 pb-2 text-sm text-[var(--text-primary)] placeholder:text-[var(--text-subtle)]"
              />
              <div className="flex items-center px-3 pb-3 gap-2">
                <div className="text-[10px] text-[var(--text-subtle)] flex items-center gap-1 font-medium">
                  <Sparkles size={12} className="text-[var(--accent-primary)]"/> Gridex Smart
                </div>
                <div className="flex-1"/>
                <span className="hidden sm:block text-[10px] text-[var(--text-faint)]">ENTER TO SEND · SHIFT + ENTER FOR NEW LINE</span>
                <button
                  onClick={()=>sendMessage()}
                  disabled={!input.trim()||sending}
                  className="w-9 h-9 rounded-xl bg-[var(--accent-primary)] hover:bg-[var(--accent-primary-hover)] disabled:opacity-40 disabled:hover:bg-[var(--accent-primary)] text-white dark:text-black grid place-items-center transition cursor-pointer"
                >
                  <ArrowUp size={18}/>
                </button>
              </div>
            </div>
            <p className="text-center text-[10px] text-[var(--text-subtle)] mt-3">ASIF AI can make mistakes. Verify important information.</p>
          </div>
        </div>
      </section>
    </main>
  );
}

function BrandMark() {
  return (
    <div className="flex items-center gap-3">
      <div className="w-10 h-10 rounded-xl bg-black border border-green-500/80 overflow-hidden grid place-items-center shadow-[0_0_22px_rgba(34,197,94,.2)] shrink-0">
        <Image src="/icon.svg" alt="ASIF AI Logo" width={40} height={40} className="w-full h-full object-cover" priority />
      </div>
      <div>
        <div className="font-bold tracking-[.16em] text-[var(--accent-badge-text)] neon-text leading-tight">ASIF AI</div>
        <div className="text-[9px] tracking-[.2em] text-[var(--text-subtle)] mt-0.5 font-medium">INTELLIGENCE SYSTEM</div>
      </div>
    </div>
  );
}
