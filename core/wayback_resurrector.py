"""
Wayback Machine Necromancer & Entity Re-Synthesizer
Autonome Rekonstruktion historischer Backlink-Pfade aus dem Internet Archive (CDX API)
mit intelligenter Entity-Erkennung und Content-Neugenerierung.
"""
import urllib.request
import urllib.parse
import json
import re
import logging
from typing import List, Dict, Any, Optional

logger = logging.getLogger("WaybackResurrector")

class WaybackResurrector:
    def __init__(self, user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"):
        self.headers = {"User-Agent": user_agent}

    def fetch_historical_urls(self, domain: str, max_urls: int = 15) -> List[Dict[str, Any]]:
        """
        Fragt die offizielle Wayback Machine CDX API ab, um die historischen
        URLs mit Status 200 abzurufen. Falls offline oder Timeout, greift ein
        ausgeklügeltes Nischen-Heuristik-Fallback.
        """
        cdx_url = f"https://web.archive.org/cdx/search/cdx?url={domain}/*&output=json&fl=original,timestamp,statuscode,mimetype&filter=statuscode:200&filter=mimetype:text/html&collapse=urlkey&limit={max_urls*2}"
        
        results = []
        try:
            req = urllib.request.Request(cdx_url, headers=self.headers)
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                if len(data) > 1:
                    headers_row = data[0]
                    for row in data[1:]:
                        orig_url = row[0]
                        timestamp = row[1]
                        parsed = urllib.parse.urlparse(orig_url)
                        path = parsed.path
                        if path and path not in ["/", ""]:
                            # Filter statischen Müll aus
                            if not re.search(r'\.(jpg|png|gif|css|js|pdf|ico|svg)$', path, re.I):
                                results.append({
                                    "original_url": orig_url,
                                    "path": path,
                                    "timestamp": timestamp,
                                    "source": "WAYBACK_CDX_LIVE"
                                })
        except Exception as e:
            logger.warning(f"Wayback CDX API nicht erreichbar ({e}). Aktiviere Heuristik-Engine.")

        if not results:
            # Autonomes Heuristik-Fallback basierend auf Nische & Keyword-Struktur
            results = self._generate_heuristic_historical_paths(domain)

        # Dubletten nach Pfad bereinigen & limitieren
        unique_paths = {}
        for r in results:
            cleaned_path = r["path"].strip().lower()
            if cleaned_path not in unique_paths and cleaned_path not in ["/", ""]:
                unique_paths[cleaned_path] = r

        return list(unique_paths.values())[:max_urls]

    def _generate_heuristic_historical_paths(self, domain: str) -> List[Dict[str, Any]]:
        """
        Erzeugt semantisch hochrelevante historische URL-Pfade für Domains,
        um sicherzustellen, dass eingehende typische Deep-Links niemals 404 werfen.
        """
        base_name = domain.split(".")[0].lower()
        parts = [p for p in re.split(r'[-_]', base_name) if len(p) > 2]
        primary_kw = parts[0] if parts else "service"
        secondary_kw = parts[1] if len(parts) > 1 else "vergleich"

        patterns = [
            f"/{primary_kw}-ratgeber",
            f"/{primary_kw}-kosten-rechner",
            f"/{secondary_kw}-uebersicht",
            f"/testsieger-vergleich",
            f"/haeufige-fragen-faq",
            f"/kontakt-beratung",
            f"/erfahrungen-bewertungen"
        ]

        return [
            {
                "original_url": f"https://{domain}{p}",
                "path": p,
                "timestamp": "20230101120000",
                "source": "HEURISTIC_SEO_RECOVERY"
            }
            for p in patterns
        ]

    def extract_core_entities(self, domain: str, niche: str) -> Dict[str, Any]:
        """
        Extrahiert die semantischen Entitäten für maximale Relevanz bei Google 2026.
        """
        domain_clean = domain.replace("-", " ").replace(".", " ")
        return {
            "primary_topic": niche,
            "core_keywords": [domain.split(".")[0].replace("-", " "), f"{niche} Rechner", f"{niche} Test 2026", f"beste {niche} Anbieter"],
            "target_intent": "TRANSACTIONAL_COMMERCIAL",
            "trust_signals": [
                "Zertifizierte Vergleichsdaten 2026",
                "Unabhängige Marktanalyse nach Verbraucherstandards",
                "TÜV & DIN-Geprüfte Methodik"
            ]
        }
