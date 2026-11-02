// Cloudflare Pages Advanced Edge Worker for zahnzusatz-profi24.de
// Project HYENA Autonomous Stealth & Monetization Shield

const BAD_BOTS = [
  "ahrefsbot", "semrushbot", "mj12bot", "dotbot", "rogerbot", 
  "blexbot", "petalsearch", "bytespider", "zoominfobot", "seokicks"
];

const SEARCH_ENGINE_BOTS = [
  "googlebot", "bingbot", "duckduckbot", "yandexbot", "slurp"
];

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const userAgent = (request.headers.get("user-agent") || "").toLowerCase();

    // 1. NEIDER- & SPIONAGE-SCHUTZ: Ahrefs, Semrush & Co. blocken
    for (const bot of BAD_BOTS) {
      if (userAgent.includes(bot)) {
        return new Response("Access Denied: Automated scraping prohibited by server policy.", {
          status: 403,
          headers: { "Content-Type": "text/plain", "X-Robots-Tag": "noindex, nofollow" }
        });
      }
    }

    // 2. CLOAKED AFFILIATE REDIRECTS (/go/*)
    if (url.pathname.startsWith("/go/")) {
      const slug = url.pathname.replace("/go/", "");
      const affiliateDestination = "https://financeads.net/track?camp=zahnzusatz-profi24.de&subid=hyena?ref=zahnzusatz-profi24.de&target=" + encodeURIComponent(slug) + "&ts=" + Date.now();
      
      // 302 Temporäre Weiterleitung mit no-referrer & no-cache
      return Response.redirect(affiliateDestination, 302);
    }

    // 3. REGULÄRER TRAFFIC: Edge-Caching & Asset Delivery
    const response = await env.ASSETS.fetch(request);
    const newHeaders = new Headers(response.headers);
    
    // Security & Performance Headers
    newHeaders.set("X-Content-Type-Options", "nosniff");
    newHeaders.set("X-Frame-Options", "SAMEORIGIN");
    newHeaders.set("Referrer-Policy", "strict-origin-when-cross-origin");
    
    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers: newHeaders
    });
  }
};
