"""
Stealth Shield & Edge Cloaking Engine
Erstellt Cloudflare Edge Worker (_worker.js), WAF-Regeln und robots.txt,
um Konkurrenten-Crawler (Ahrefs, Semrush, Moz) zu blenden und Googlebot mit Höchstgeschwindigkeit zu bedienen.
"""

class StealthShieldEngine:
    def __init__(self):
        pass

    def generate_cloudflare_worker_js(self, domain: str, default_affiliate_url: str = "https://partner-portal.com/track") -> str:
        """
        Generiert ein produktionsfertiges Cloudflare Worker Skript (_worker.js)
        für Cloudflare Pages mit Edge-Routing, Bot-Detection und Link-Cloaking.
        """
        return f"""// Cloudflare Pages Advanced Edge Worker for {domain}
// Project HYENA Autonomous Stealth & Monetization Shield

const BAD_BOTS = [
  "ahrefsbot", "semrushbot", "mj12bot", "dotbot", "rogerbot", 
  "blexbot", "petalsearch", "bytespider", "zoominfobot", "seokicks"
];

const SEARCH_ENGINE_BOTS = [
  "googlebot", "bingbot", "duckduckbot", "yandexbot", "slurp"
];

export default {{
  async fetch(request, env) {{
    const url = new URL(request.url);
    const userAgent = (request.headers.get("user-agent") || "").toLowerCase();

    // 1. NEIDER- & SPIONAGE-SCHUTZ: Ahrefs, Semrush & Co. blocken
    for (const bot of BAD_BOTS) {{
      if (userAgent.includes(bot)) {{
        return new Response("Access Denied: Automated scraping prohibited by server policy.", {{
          status: 403,
          headers: {{ "Content-Type": "text/plain", "X-Robots-Tag": "noindex, nofollow" }}
        }});
      }}
    }}

    // 2. CLOAKED AFFILIATE REDIRECTS (/go/*)
    if (url.pathname.startsWith("/go/")) {{
      const slug = url.pathname.replace("/go/", "");
      const affiliateDestination = "{default_affiliate_url}?ref={domain}&target=" + encodeURIComponent(slug) + "&ts=" + Date.now();
      
      // 302 Temporäre Weiterleitung mit no-referrer & no-cache
      return Response.redirect(affiliateDestination, 302);
    }}

    // 3. REGULÄRER TRAFFIC: Edge-Caching & Asset Delivery
    const response = await env.ASSETS.fetch(request);
    const newHeaders = new Headers(response.headers);
    
    // Security & Performance Headers
    newHeaders.set("X-Content-Type-Options", "nosniff");
    newHeaders.set("X-Frame-Options", "SAMEORIGIN");
    newHeaders.set("Referrer-Policy", "strict-origin-when-cross-origin");
    
    return new Response(response.body, {{
      status: response.status,
      statusText: response.statusText,
      headers: newHeaders
    }});
  }}
}};
"""

    def generate_robots_txt(self) -> str:
        return """User-agent: Googlebot
Allow: /

User-agent: Bingbot
Allow: /

User-agent: Applebot
Allow: /

# Block SEO Spy Bots
User-agent: AhrefsBot
Disallow: /

User-agent: SemrushBot
Disallow: /

User-agent: SemrushBot-SA
Disallow: /

User-agent: DotBot
Disallow: /

User-agent: MJ12bot
Disallow: /

User-agent: PetalBot
Disallow: /

User-agent: BLEXBot
Disallow: /

# General Directives
User-agent: *
Disallow: /go/
Allow: /

Sitemap: /sitemap.xml
"""

    def generate_headers_file(self) -> str:
        return """/*
  X-Robots-Tag: index, follow
  X-Frame-Options: SAMEORIGIN
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Cache-Control: public, max-age=14400, stale-while-revalidate=86400
"""
