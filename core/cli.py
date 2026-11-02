"""
Kommandozeilen-Interface (CLI) für den Domain Agenten.
"""
import sys
import click
from .models import ActionRecommendation
from .orchestrator import DomainAcquisitionOrchestrator
from .sources import LiveScoutingEngine

@click.group()
def cli():
    """Autonomous Expired Domain Acquisition & SEO Automation CLI"""
    pass

@cli.command("crawl-site")
@click.argument("url")
@click.option("--auto-catch", is_flag=True, help="Automatischen Drop-Catch / Registrierung scharf schalten")
@click.option("--pseo", is_flag=True, help="Sofortige pSEO-Landingpage für Treffer generieren")
def crawl_site(url, auto_catch, pseo):
    """Crawlt eine Ziel-Website nach toten, registrierbaren Outbound-Links."""
    scout = LiveScoutingEngine()
    orchestrator = DomainAcquisitionOrchestrator()

    click.echo(f"[*] Starte Broken Link Crawl auf: {url}")
    candidates = scout.crawl_target_site_for_broken_domains(url)

    if not candidates:
        click.echo("[i] Keine abgelaufenen externen Domains auf der Zielseite gefunden.")
        return

    click.echo(f"\n[+] {len(candidates)} potenzielle Expired Domains gefunden. Starte Filter & Valuation...\n")
    _display_and_process(orchestrator, candidates, auto_catch, pseo)

@cli.command("import-csv")
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--auto-catch", is_flag=True, help="Automatischen Drop-Catch / Registrierung scharf schalten")
@click.option("--pseo", is_flag=True, help="Sofortige pSEO-Landingpage für Treffer generieren")
def import_csv(file_path, auto_catch, pseo):
    """Importiert und bewertet eine ExpiredDomains-Löschliste im CSV-Format."""
    scout = LiveScoutingEngine()
    orchestrator = DomainAcquisitionOrchestrator()

    click.echo(f"[*] Importiere Datei: {file_path}")
    candidates = scout.ingest_expired_domains_csv(file_path)

    if not candidates:
        click.echo("[!] Keine gültigen Domains in der Datei gefunden.")
        return

    click.echo(f"\n[+] {len(candidates)} Domains importiert. Starte Analyse...\n")
    _display_and_process(orchestrator, candidates, auto_catch, pseo)

def _display_and_process(orchestrator, candidates, auto_catch, pseo):
    click.echo(f"{'Domain':<32} | {'DP/BL':<12} | {'EstiBot':<9} | {'Spam':<6} | {'Empfohlene Aktion':<24}")
    click.echo("-" * 90)

    for lead in candidates:
        res = orchestrator.process_domain_candidate(lead, auto_snipe=auto_catch, generate_pseo=pseo)
        m = res["metrics"]
        ev = res["evaluation"]
        dp_bl = f"{m['dp']}/{m['bl']}"
        val = f"${m['estibot_value']:.0f}" if m.get("estibot_value") else "$0"
        spam = f"{m['spam_score']:.2f}"
        rec = ev["recommendation"].replace("ActionRecommendation.", "")
        click.echo(f"{lead.domain:<32} | {dp_bl:<12} | {val:<9} | {spam:<6} | {rec:<24}")

if __name__ == "__main__":
    cli()
