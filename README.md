# ASIF AI

A private, cyberpunk-green AI chat application using Gridex's OpenAI-compatible Chat Completions API, Supabase Auth/Postgres, and Vercel.

## Features
- Supabase email/password authentication
- Per-user chat history protected by PostgreSQL Row Level Security
- Gridex `smart` model called only from a server route
- Reliable OpenAI-compatible chat completions (non-streaming mode, verified against the supplied Gridex request format)
- Responsive dark neon-green interface

## 1. Requirements
- Node.js 20.9+ (Node.js 22 recommended)
- A Supabase project
- A Gridex API key
- A Vercel account for hosting

## 2. Configure Supabase
1. Create a dedicated Supabase project for ASIF AI.
2. Open SQL Editor and run `supabase/schema.sql`.
3. In Project Settings -> API, copy the Project URL and Publishable key.
4. In Authentication -> URL Configuration, set the Site URL to your eventual Vercel URL. Add the Vercel preview URL to Redirect URLs if you use previews.
5. For a private single-owner app, create your own account first, then disable public signups in Authentication -> Providers -> Email. Existing users can still sign in.

## 3. Configure environment variables
Copy `.env.example` to `.env.local` and fill in:

```env
NEXT_PUBLIC_SUPABASE_URL=https://YOUR_PROJECT_REF.supabase.co
NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=sb_publishable_...
GRIDEX_BASE_URL=https://gridex-jtlpk2wknu.agent37.app/v1
GRIDEX_API_KEY=your_private_gridex_key
GRIDEX_MODEL=smart
```

Never prefix `GRIDEX_API_KEY` with `NEXT_PUBLIC_`. Do not commit `.env.local`.

## 4. Run locally

```bash
npm install
npm run dev
```

Open http://localhost:3000.

## 5. Deploy to Vercel
1. Push this folder to a private GitHub repository, or import it directly into Vercel.
2. Add all five environment variables in Vercel -> Project -> Settings -> Environment Variables (Production, Preview, Development as required).
3. Deploy.
4. Set Supabase Auth Site URL and Redirect URLs to the deployed Vercel URL.
5. Sign up with your own email, verify it if email confirmation is enabled, and sign in.
6. Once your account is created, disable public email signups if this should remain single-user/private.

## Notes
- This project does not use Open WebUI. It is a separate ASIF AI website.
- The Gridex API key is only read by `app/api/chat/route.ts` on the server.
- The app currently requests non-streaming completions because the supplied Gridex test confirmed the standard Chat Completions response format. Streaming can be added after confirming that the Gridex gateway supports it.
- For public multi-user use, add per-user quotas/rate limits and abuse monitoring before sharing the URL.
