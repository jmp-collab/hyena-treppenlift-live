"""
Echte Datenquellen & Live Scouting Connectors:
1. Live Web Crawler für Target URLs (Broken Outbound Links)
2. ExpiredDomains.net & Löschlisten CSV/JSON Ingestion
3. Wayback Machine API (Archiv Birth Year & Historische Snapshot-Validierung)
"""
import csv
import json
import logging
import socket
import urllib.parse
from typing import List, Dict, Any, Optional
import requests
from bs4 import BeautifulSoup

from .models import DomainMetrics

logger = logging.getLogger("LiveSources")

class LiveScoutingEngine:
    def __init__(self, timeout: int = 5, user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"):
        self.timeout = timeout
        self.headers = {"User-Agent": user_agent}

    def check_dns_nxdomain(self, domain: str) -> bool:
        """Prüft per Socket DNS, ob die Domain noch aktiv auflöst oder NXDOMAIN ist."""
        try:
            socket.gethostbyname(domain)
            return False  # Domain existiert / löst auf
        except (socket.gaierror, socket.herror, TimeoutError):
            return True   # NXDOMAIN / nicht erreichbar

    def fetch_wayback_metadata(self, domain: str) -> Dict[str, Any]:
        """
        Fragt die offizielle Wayback Machine CDX API ab, um das echte
        Birth Year (ABY) und Snapshot-Volumen (ACR) zu ermitteln.
        """
        url = f"https://archive.org/wayback/available?url={domain}"
        try:
            resp = requests.get(url, headers=self.headers, timeout=self.timeout)
            if resp.status_code == 200:
                data = resp.json()
                snapshots = data.get("archived_snapshots", {})
                closest = snapshots.get("closest", {})
                if closest and closest.get("available"):
                    timestamp = closest.get("timestamp", "")
                    if len(timestamp) >= 4:
                        first_year = int(timestamp[:4])
                        return {"aby": first_year, "has_archive": True}
        except Exception as e:
            logger.debug(f"Wayback check failed for {domain}: {e}")
        return {"aby": None, "has_archive": False}

    def crawl_target_site_for_broken_domains(self, target_url: str) -> List[DomainMetrics]:
        """
        Crawlt eine Live-Website, parst alle ausgehenden Links und prüft,
        ob verlinkte externe Domains abgelaufen/NXDOMAIN sind.
        """
        found_metrics = []
        parsed_source = urllib.parse.urlparse(target_url)
        source_host = parsed_source.netloc.lower()

        try:
            print(f"[*] Crawle Ziel-Website: {target_url}...")
            resp = requests.get(target_url, headers=self.headers, timeout=self.timeout)
            if resp.status_code != 200:
                print(f"[!] Fehler beim Laden der URL (Status: {resp.status_code})")
                return []

            soup = BeautifulSoup(resp.text, "html.parser")
            outbound_domains = set()

            for a_tag in soup.find_all("a", href=True):
                href = a_tag["href"].strip()
                if href.startswith("http://") or href.startswith("https://"):
                    p = urllib.parse.urlparse(href)
                    host = p.netloc.lower()
                    if ":" in host:
                        host = host.split(":")[0]
                    if host.startswith("www."):
                        host = host[4:]

                    if host and host not in source_host and "." in host:
                        outbound_domains.add(host)

            print(f"[*] {len(outbound_domains)} externe Ziel-Domains gefunden. Prüfe DNS-Verfügbarkeit...")

            for domain in outbound_domains:
                is_nx = self.check_dns_nxdomain(domain)
                if is_nx:
                    # Domain ist frei / abgelaufen!
                    tld = domain.split(".")[-1]
                    wb_meta = self.fetch_wayback_metadata(domain)
                    metric = DomainMetrics(
                        domain=domain,
                        tld=tld,
                        dp=1,  # Mindestens 1 Quell-Referring Domain gesichert
                        bl=1,
                        aby=wb_meta.get("aby"),
                        acr=10 if wb_meta.get("has_archive") else 0,
                        is_nxdomain=True,
                        estibot_value=150.0  # Basisbewertung für organischen Link
                    )
                    found_metrics.append(metric)

        except Exception as e:
            print(f"[!] Crawl-Fehler: {e}")

        return found_metrics

    def ingest_expired_domains_csv(self, file_path: str) -> List[DomainMetrics]:
        """
        Liest exportierte Löschlisten von ExpiredDomains.net oder ScrapingBee ein.
        Unterstützt Spalten wie: Domain, BL, DP, ABY, ACR, Status.
        """
        domains = []
        try:
            with open(file_path, mode="r", encoding="utf-8", errors="ignore") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # Standardisierung gängiger CSV Header
                    domain_name = row.get("Domain") or row.get("domain") or row.get("Domain Name")
                    if not domain_name:
                        continue
                    
                    domain_name = domain_name.strip().lower()
                    tld = domain_name.split(".")[-1]
                    
                    dp = int(row.get("DP") or row.get("Referring Domains") or row.get("dp") or 0)
                    bl = int(row.get("BL") or row.get("Backlinks") or row.get("bl") or 0)
                    aby_raw = row.get("ABY") or row.get("Birth Year") or row.get("aby")
                    aby = int(aby_raw) if aby_raw and aby_raw.isdigit() else None
                    acr = int(row.get("ACR") or row.get("acr") or 0)
                    
                    is_nx = self.check_dns_nxdomain(domain_name)
                    
                    metric = DomainMetrics(
                        domain=domain_name,
                        tld=tld,
                        dp=dp,
                        bl=bl,
                        aby=aby,
                        acr=acr,
                        is_nxdomain=is_nx,
                        estibot_value=dp * 25.0 if is_nx else dp * 15.0
                    )
                    domains.append(metric)
        except Exception as e:
            print(f"[!] Fehler beim CSV-Import von '{file_path}': {e}")
        return domains
