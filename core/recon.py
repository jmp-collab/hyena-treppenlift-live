"""
Phase 1: Reconnaissance & Discovery Engine
Scraped Drop-Listen, Broken Outbound Links und analysiert Trust-Beziehungen.
"""
import re
import socket
from typing import List, Dict, Any, Generator
import urllib.parse

class ReconEngine:
    def __init__(self, user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"):
        self.user_agent = user_agent

    def is_nxdomain(self, domain: str) -> bool:
        """Ermittelt schnell via DNS, ob eine Domain nicht mehr existiert/auflöst."""
        try:
            socket.gethostbyname(domain)
            return False  # Existiert noch / löst auf
        except (socket.gaierror, socket.herror, TimeoutError):
            return True   # NXDOMAIN oder nicht erreichbar

    def extract_broken_external_links(self, html_content: str, source_domain: str) -> List[str]:
        """
        Scannt HTML-Inhalte (wie BennyWindolph / Williams-Media Expired-Domain-Finder)
        nach toten externen Links, die ggf. frei registrierbar sind.
        """
        href_pattern = re.compile(r'href=[\'"](https?://[^\'">]+)[\'"]', re.IGNORECASE)
        found_links = href_pattern.findall(html_content)
        
        expired_candidates = set()
        for link in found_links:
            try:
                parsed = urllib.parse.urlparse(link)
                target_host = parsed.netloc.lower()
                
                # Entferne Ports & www
                if ":" in target_host:
                    target_host = target_host.split(":")[0]
                if target_host.startswith("www."):
                    target_host = target_host[4:]
                    
                # Ignoriere Links auf die eigene Quell-Domain
                if not target_host or target_host in source_domain:
                    continue
                    
                # Prüfe NXDOMAIN Status
                if self.is_nxdomain(target_host):
                    expired_candidates.add(target_host)
            except Exception:
                continue
                
        return sorted(list(expired_candidates))

    def parse_crossdomain_xml(self, xml_content: str) -> List[str]:
        """
        Prüft wie NetSPI/crossdomainscanner crossdomain.xml-Dateien
        auf verwaiste Trust-Domänen mit Wildcard- oder spezifischen Policies.
        """
        candidates = []
        domain_pattern = re.compile(r'allow-access-from\s+domain=[\'"]([^\'">]+)[\'"]', re.IGNORECASE)
        matches = domain_pattern.findall(xml_content)
        
        for entry in matches:
            cleaned = entry.replace("*.", "").strip()
            if cleaned and cleaned != "*" and self.is_nxdomain(cleaned):
                candidates.append(cleaned)
        return candidates
