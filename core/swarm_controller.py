"""
Swarm Controller - Project HYENA Autonomous Multi-Agent Core
Orchestriert den gesamten Workflow:
1. Scout / Qualifikation
2. Wayback Necromancy & Path-Recovery
3. Affiliate & Lead-Funnel Injection
4. Stealth Shield & Cloudflare Deployment-Packaging
"""
import os
import re
import json
import sqlite3
import datetime
from typing import Dict, Any, List, Optional
from .wayback_resurrector import WaybackResurrector
from .affiliate_engine import AffiliateFunnelEngine
from .stealth_shield import StealthShieldEngine

class HyenaSwarmController:
    def __init__(self, db_path: str = "data/domains.sqlite", deployments_dir: str = "deployments"):
        self.db_path = db_path
        self.deployments_dir = os.path.abspath(deployments_dir)
        self.resurrector = WaybackResurrector()
        self.affiliate = AffiliateFunnelEngine()
        self.shield = StealthShieldEngine()
        os.makedirs(self.deployments_dir, exist_ok=True)

    def resurrect_domain_bundle(self, domain: str, niche: str = "", custom_affiliate_url: str = "") -> Dict[str, Any]:
        """
        Erstellt ein vollständiges, sofort deploybares statisches Webseiten-Bundle
        für eine gegebene Domain inklusive aller historischen Backlink-Pfade.
        """
        bundle_dir = os.path.join(self.deployments_dir, domain)
        os.makedirs(bundle_dir, exist_ok=True)

        # 1. Nischenprofil & Monetarisierungshebel bestimmen
        profile = self.affiliate.get_profile(domain, niche)
        entities = self.resurrector.extract_core_entities(domain, profile["detected_niche"])
        
        # 2. Historische URL-Pfade ermitteln
        historical_paths = self.resurrector.fetch_historical_urls(domain, max_urls=8)

        # 3. Interaktiven Rechner bauen
        calculator_html = self.affiliate.build_interactive_calculator_html(domain, profile)

        # 4. Schema.org JSON-LD erzeugen
        schema_json = {
            "@context": "https://schema.org",
            "@graph": [
                {
                    "@type": "WebSite",
                    "name": f"{domain.upper()} - {profile['headline']}",
                    "url": f"https://{domain}",
                    "description": profile["subline"]
                },
                {
                    "@type": "FAQPage",
                    "mainEntity": [
                        {
                            "@type": "Question",
                            "name": f"Wie funktioniert der Vergleich auf {domain}?",
                            "acceptedAnswer": {
                                "@type": "Answer",
                                "text": f"Unser Algorithmus vergleicht tagesaktuell die führenden Anbieter für {profile['detected_niche']} und ermittelt das beste Preis-Leistungs-Verhältnis."
                            }
                        },
                        {
                            "@type": "Question",
                            "name": "Ist der Service kostenlos?",
                            "acceptedAnswer": {
                                "@type": "Answer",
                                "text": "Ja, der gesamte Berechnungsprozess und die Angebotserstellung sind zu 100% kostenfrei und unverbindlich."
                            }
                        }
                    ]
                }
            ]
        }

        # 5. Navigation & Menü
        nav_links_html = "".join([
            f'<a href="{p["path"] if not p["path"].endswith(".html") else p["path"]}" class="text-sm font-medium text-slate-600 hover:text-blue-600 transition">{p["path"].strip("/").replace("-", " ").title()}</a>'
            for p in historical_paths[:4]
        ])

        # 6. Content Section (Trust & Authority)
        content_section = f"""
        <section class="max-w-4xl mx-auto py-12 px-4 sm:px-0 border-t border-slate-200">
            <h3 class="text-2xl font-bold text-slate-900 mb-6">Warum {domain}? Qualitätsstandards & Methodik 2026</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-8 text-slate-600 text-sm leading-relaxed mb-8">
                <div>
                    <h4 class="font-bold text-slate-800 text-base mb-2">Transparente Marktübersicht</h4>
                    <p>Wir analysieren den Markt für {profile['detected_niche']} nach strengen Kriterien: Transparenz der Kosten, Kundenzufriedenheit und Ausführungsgeschwindigkeit. Unsere Daten werden wöchentlich mit den Konditionen der führenden Anbieter synchronisiert.</p>
                </div>
                <div>
                    <h4 class="font-bold text-slate-800 text-base mb-2">Unabhängige Vorauswahl</h4>
                    <p>Sie erhalten keine pauschalen Standardangebote, sondern maßgeschneiderte Empfehlungen, die exakt auf Ihre Parameter abgestimmt sind. Dadurch sparen Sie wertvolle Recherchezeit und vermeiden versteckte Zusatzkosten.</p>
                </div>
            </div>
            
            <div class="bg-blue-50 border border-blue-200 rounded-xl p-6 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div>
                    <div class="font-bold text-slate-900">Möchten Sie eine persönliche Erstberatung?</div>
                    <div class="text-xs text-slate-600">Unsere Partner-Experten stehen Ihnen unverbindlich zur Seite.</div>
                </div>
                <a href="/go/{profile['go_slug']}" class="px-5 py-2.5 bg-blue-600 text-white font-bold text-xs rounded-lg hover:bg-blue-700 transition">
                    Jetzt unverbindlich anfragen &rarr;
                </a>
            </div>
        </section>
        """

        # 7. Vollständiges index.html generieren
        homepage_html = f"""<!DOCTYPE html>
<html lang="de" class="scroll-smooth">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{profile['headline']} | {domain}</title>
    <meta name="description" content="{profile['subline']}">
    <link rel="canonical" href="https://{domain}/">
    <script src="https://cdn.tailwindcss.com"></script>
    <script type="application/ld+json">
{json.dumps(schema_json, indent=2)}
    </script>
</head>
<body class="bg-slate-50 text-slate-800 font-sans antialiased min-h-screen flex flex-col justify-between">
    <header class="bg-white/95 backdrop-blur-sm sticky top-0 z-50 border-b border-slate-200 shadow-sm">
        <div class="max-w-6xl mx-auto px-4 py-3 flex justify-between items-center">
            <a href="/" class="flex items-center space-x-2">
                <div class="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center text-white font-black text-lg">H</div>
                <span class="font-black text-xl tracking-tight text-slate-900">{domain.upper()}</span>
            </a>
            <nav class="hidden md:flex space-x-6 items-center">
                {nav_links_html}
                <a href="/go/{profile['go_slug']}" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition shadow-sm">
                    Vergleich starten
                </a>
            </nav>
        </div>
    </header>

    <main class="max-w-5xl mx-auto px-4 py-10 flex-grow w-full">
        <!-- Hero Tag -->
        <div class="text-center max-w-3xl mx-auto mb-8">
            <span class="inline-block py-1 px-3 rounded-full text-xs font-bold bg-blue-100 text-blue-800 tracking-wide uppercase mb-3">
                Verbraucherportal • Stand 2026
            </span>
            <h1 class="text-3xl sm:text-5xl font-black text-slate-950 tracking-tight leading-tight mb-4">
                {profile['headline']}
            </h1>
            <p class="text-base sm:text-lg text-slate-600 max-w-2xl mx-auto">
                {profile['subline']}
            </p>
        </div>

        <!-- Interactive Calculator -->
        {calculator_html}

        <!-- Restored Authority & Content -->
        {content_section}
    </main>

    <footer class="bg-slate-900 text-slate-400 py-10 border-t border-slate-800 text-xs">
        <div class="max-w-6xl mx-auto px-4 grid grid-cols-1 md:grid-cols-3 gap-8 mb-8">
            <div>
                <span class="font-bold text-white text-sm block mb-2">{domain.upper()}</span>
                <p class="text-slate-400 leading-relaxed">Unabhängiges Verbraucher- und Vergleichsportal. Alle Berechnungen erfolgen auf Basis tagesaktueller Marktdaten führender Anbieter.</p>
            </div>
            <div>
                <span class="font-bold text-white text-sm block mb-2">Relevante Bereiche</span>
                <ul class="space-y-1">
                    {''.join([f'<li><a href="{p["path"]}" class="hover:text-white transition">{p["path"].strip("/").replace("-", " ").title()}</a></li>' for p in historical_paths[:5]])}
                </ul>
            </div>
            <div>
                <span class="font-bold text-white text-sm block mb-2">Rechtliches & Transparenz</span>
                <p class="text-slate-400 mb-2">Hinweis: Diese Seite finanziert sich über qualifizierte Vermittlungsprovisionen (Affiliate). Für Sie entstehen dadurch keinerlei Mehrkosten.</p>
                <div class="space-x-4">
                    <a href="/impressum" class="hover:text-white underline">Impressum</a>
                    <a href="/datenschutz" class="hover:text-white underline">Datenschutz</a>
                </div>
            </div>
        </div>
        <div class="max-w-6xl mx-auto px-4 text-center text-slate-500 border-t border-slate-800 pt-6">
            &copy; 2026 {domain}. Alle Rechte vorbehalten. Re-Engineered with Project HYENA Autonomous Swarm.
        </div>
    </footer>
</body>
</html>"""

        with open(os.path.join(bundle_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(homepage_html)

        # 8. Unterseiten für die historischen Pfade generieren (keine 404 Fehler!)
        for p in historical_paths:
            path_clean = p["path"].strip("/")
            file_name = f"{path_clean}.html" if not path_clean.endswith(".html") else path_clean
            file_name = file_name.replace("/", "_")
            if not file_name:
                continue
            
            subpage_title = path_clean.replace("-", " ").replace("_", " ").title()
            subpage_html = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{subpage_title} - Aktueller Leitfaden 2026 | {domain}</title>
    <meta name="description" content="Detaillierte Informationen und aktuelle Vergleichsdaten zu {subpage_title} auf {domain}.">
    <link rel="canonical" href="https://{domain}/{path_clean}">
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-50 text-slate-800 font-sans antialiased min-h-screen flex flex-col justify-between">
    <header class="bg-white border-b border-slate-200 py-4 px-6 flex justify-between items-center">
        <a href="/" class="font-black text-xl text-blue-600">{domain.upper()}</a>
        <a href="/" class="text-sm font-semibold text-slate-600 hover:text-blue-600">&larr; Zurück zur Hauptseite</a>
    </header>
    <main class="max-w-4xl mx-auto px-4 py-12 flex-grow">
        <span class="text-xs uppercase tracking-wider text-blue-600 font-bold bg-blue-50 px-3 py-1 rounded-full">Fach-Dossier</span>
        <h1 class="text-3xl sm:text-4xl font-black text-slate-900 mt-3 mb-6">{subpage_title}</h1>
        <div class="prose max-w-none text-slate-700 leading-relaxed space-y-4 mb-8">
            <p>Dieser Leitfaden fasst die wichtigsten gesetzlichen und wirtschaftlichen Rahmenbedingungen zu <strong>{subpage_title}</strong> zusammen.</p>
            <p>Aufgrund aktueller Marktveränderungen empfehlen wir, Konditionen vor Abschluss sorgfältig zu vergleichen.</p>
        </div>
        
        <div class="p-6 bg-blue-50 rounded-2xl border border-blue-200">
            <h3 class="text-lg font-bold text-slate-900 mb-2">Interaktiven Vergleichsrechner nutzen</h3>
            <p class="text-sm text-slate-600 mb-4">Ermitteln Sie Ihre individuellen Einsparpotenziale in 60 Sekunden.</p>
            <a href="/#vergleich" class="inline-block px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm rounded-xl shadow-md transition">
                Zum Hauptrechner wechseln &rarr;
            </a>
        </div>
    </main>
    <footer class="bg-white border-t border-slate-200 py-6 text-center text-xs text-slate-500">
        &copy; 2026 {domain} • Historischer Pfad erfolgreich reaktiviert.
    </footer>
</body>
</html>"""
            with open(os.path.join(bundle_dir, file_name), "w", encoding="utf-8") as f:
                f.write(subpage_html)

        # 9. Edge-Worker & Stealth-Dateien schreiben
        default_aff = custom_affiliate_url or f"https://financeads.net/track?camp={domain}&subid=hyena"
        with open(os.path.join(bundle_dir, "_worker.js"), "w", encoding="utf-8") as f:
            f.write(self.shield.generate_cloudflare_worker_js(domain, default_aff))

        with open(os.path.join(bundle_dir, "_headers"), "w", encoding="utf-8") as f:
            f.write(self.shield.generate_headers_file())

        with open(os.path.join(bundle_dir, "robots.txt"), "w", encoding="utf-8") as f:
            f.write(self.shield.generate_robots_txt())

        # 10. Sitemap XML
        sitemap_entries = [f"<url><loc>https://{domain}/</loc><priority>1.0</priority></url>"]
        for p in historical_paths:
            sitemap_entries.append(f"<url><loc>https://{domain}{p['path']}</loc><priority>0.8</priority></url>")
        sitemap_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
{''.join(sitemap_entries)}
</urlset>"""
        with open(os.path.join(bundle_dir, "sitemap.xml"), "w", encoding="utf-8") as f:
            f.write(sitemap_xml)

        return {
            "domain": domain,
            "status": "DEPLOYMENT_BUNDLE_READY",
            "detected_niche": profile["detected_niche"],
            "bundle_path": bundle_dir,
            "generated_files": os.listdir(bundle_dir),
            "historical_paths_recovered": len(historical_paths),
            "lead_payout": profile["lead_payout"],
            "affiliate_partner": profile["affiliate_partner"]
        }
