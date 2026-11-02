#!/usr/bin/env python3
"""
PROJECT HYENA - Autonomous Greyhat Domain Snapper & Affiliate Swarm CLI
Kommandozentrale zur Steuerung der 4 Schwarm-Agenten:
1. @drop_scout: Findet und qualifiziert freie High-Yield Domains
2. @archive_resurrector: Stellt historische Pfade & Entitäten wieder her
3. @affiliate_hijacker: Injiziert interaktive Rechner & Lead-Funnels
4. @stealth_shield: Generiert Cloudflare Edge Worker & Bot-Defense
"""
import sys
import os
import argparse
import sqlite3
import json
from core.swarm_controller import HyenaSwarmController

def print_banner():
    banner = """
  ██╗  ██╗██╗   ██╗███████╗███╗   ██╗ █████╗ 
  ██║  ██║╚██╗ ██╔╝██╔════╝████╗  ██║██╔══██╗
  ███████║ ╚████╔╝ █████╗  ██╔██╗ ██║███████║
  ██╔══██║  ╚██╔╝  ██╔══╝  ██║╚██╗██║██╔══██║
  ██║  ██║   ██║   ███████╗██║ ╚████║██║  ██║
  ╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝  ╚═══╝╚═╝  ╚═╝
  >>> AUTONOMOUS GREYHAT DOMAIN & AFFILIATE SWARM <<<
    """
    print(banner)

def get_db_connection(db_path: str = "data/domains.sqlite"):
    if not os.path.exists(db_path):
        print(f"[!] Fehler: Datenbank {db_path} nicht gefunden!")
        sys.exit(1)
    return sqlite3.connect(db_path)

def cmd_scout(args):
    conn = get_db_connection()
    cur = conn.cursor()
    limit = args.limit or 10

    print(f"\n🔍 [SCOUT-AGENT] Suche die Top {limit} lukrativsten SOFORT-KAUF Domains...\n")
    cur.execute("""
        SELECT domain, tld, reg_cost_eur, cpc, search_volume, lead_value_est, roi_value_10x, niche, affiliate_strategy
        FROM instant_buy_domains
        WHERE whois_status = '100% FREI'
        ORDER BY lead_value_est DESC
        LIMIT ?
    """, (limit,))
    rows = cur.fetchall()

    print(f"{'#':<3} | {'Domain':<32} | {'Reg (€)':<8} | {'Ø CPC':<7} | {'Volumen':<8} | {'Leadwert ($)':<12} | {'Nische':<18}")
    print("-" * 100)
    for idx, r in enumerate(rows, 1):
        dom, tld, cost, cpc, vol, val, roi, niche, strat = r
        print(f"{idx:<3} | {dom:<32} | {cost:<8.2f} | ${cpc:<6.2f} | {vol:<8} | ${val:<11.0f} | {niche:<18}")

    print("\n💡 Tipp: Führe 'python hyena_swarm.py resurrect <domain>' aus, um eine Domain sofort komplett zu bauen.\n")

def cmd_resurrect(args):
    domain = args.domain
    niche = args.niche or ""
    print(f"\n⚡ [PROJECT HYENA] Starte autonome Wiederauferstehung für: {domain}...")
    
    controller = HyenaSwarmController()
    result = controller.resurrect_domain_bundle(domain, niche=niche)

    print(f"\n[+] Status: {result['status']}")
    print(f"    - Nische erkannt: {result['detected_niche']}")
    print(f"    - Affiliate-Hebel: {result['affiliate_partner']} (Ø {result['lead_payout']} / Lead)")
    print(f"    - Historische Pfade gerettet: {result['historical_paths_recovered']} URLs (Keine 404s!)")
    print(f"    - Bundle generiert unter: {result['bundle_path']}")
    print(f"    - Dateien: {', '.join(result['generated_files'])}")
    print(f"\n🚀 Das Deployment-Paket ist bereit für Cloudflare Pages (inkl. _worker.js Bot-Shield).\n")

def cmd_bundle_top(args):
    count = args.count or 5
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT domain, niche FROM instant_buy_domains
        WHERE whois_status = '100% FREI'
        ORDER BY lead_value_est DESC
        LIMIT ?
    """, (count,))
    targets = cur.fetchall()

    print(f"\n⚡ [SWARM BATCH RUN] Baue autonome Deployment-Bundles für die Top {len(targets)} Gems...\n")
    controller = HyenaSwarmController()
    
    for idx, (dom, niche) in enumerate(targets, 1):
        print(f"[{idx}/{len(targets)}] Verarbeite {dom} ({niche})...")
        res = controller.resurrect_domain_bundle(dom, niche=niche)
        print(f"     -> {len(res['generated_files'])} Dateien erzeugt in deployments/{dom}")

    print(f"\n✅ Alle {len(targets)} Bundles erfolgreich erzeugt!")
    print(f"📂 Verzeichnis: {os.path.abspath('deployments')}\n")

def cmd_dashboard(args):
    conn = get_db_connection()
    cur = conn.cursor()

    total_free = cur.execute("SELECT count(*) FROM instant_buy_domains WHERE whois_status = '100% FREI'").fetchone()[0]
    total_val = cur.execute("SELECT sum(lead_value_est) FROM instant_buy_domains").fetchone()[0] or 0
    avg_cpc = cur.execute("SELECT avg(cpc) FROM instant_buy_domains").fetchone()[0] or 0
    total_vol = cur.execute("SELECT sum(search_volume) FROM instant_buy_domains").fetchone()[0] or 0

    exp_count = cur.execute("SELECT count(*) FROM expired_domains").fetchone()[0]
    auc_count = cur.execute("SELECT count(*) FROM auction_domains").fetchone()[0]

    deployments_dir = os.path.abspath("deployments")
    active_deployments = len([d for d in os.listdir(deployments_dir) if os.path.isdir(os.path.join(deployments_dir, d))]) if os.path.exists(deployments_dir) else 0

    print("\n" + "="*70)
    print("📊 PROJECT HYENA - LIVE ASSET DASHBOARD")
    print("="*70)
    print(f"  • Verifizierte Sofort-Kauf Domains:      {total_free} Domains")
    print(f"  • Expired Domains mit SEO-Kapital:       {exp_count} Domains")
    print(f"  • Top-Auktionen im Radar:                {auc_count} Domains")
    print(f"  • Aktive Bereitgestellte Deployments:   {active_deployments} fertige Bundles")
    print("-" * 70)
    print(f"  • Geschätzter Jahres-Leadwert (Gems):    ${total_val:,.2f}")
    print(f"  • Durchschnittlicher CPC:                ${avg_cpc:.2f}")
    print(f"  • Monatliches Suchvolumen:               {total_vol:,} Anfragen/Monat")
    print("="*70 + "\n")

def main():
    print_banner()
    parser = argparse.ArgumentParser(description="PROJECT HYENA - Greyhat Domain Snapper & Affiliate Swarm")
    subparsers = parser.add_subparsers(dest="command", help="Verfügbare Befehle")

    # scout
    p_scout = subparsers.add_parser("scout", help="Durchsucht die Datenbank nach lukrativen Targets")
    p_scout.add_argument("--limit", type=int, default=10, help="Anzahl der Ergebnisse")

    # resurrect
    p_res = subparsers.add_parser("resurrect", help="Erstellt das komplette Deployment-Bundle für eine Domain")
    p_res.add_argument("domain", type=str, help="Domainname (z.B. finanzberater-direkt24.de)")
    p_res.add_argument("--niche", type=str, default="", help="Optionale Nischen-Spezifikation")

    # bundle-top
    p_bt = subparsers.add_parser("bundle-top", help="Generiert Bundles für die Top N Domains")
    p_bt.add_argument("--count", type=int, default=5, help="Anzahl der Top-Domains")

    # dashboard
    subparsers.add_parser("dashboard", help="Zeigt das Asset & Portfolio Dashboard an")

    args = parser.parse_args()

    if args.command == "scout":
        cmd_scout(args)
    elif args.command == "resurrect":
        cmd_resurrect(args)
    elif args.command == "bundle-top":
        cmd_bundle_top(args)
    elif args.command == "dashboard":
        cmd_dashboard(args)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
