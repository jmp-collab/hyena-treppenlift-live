# Expired Domain Acquisition & SEO Automation Hub

Ein modulares, autonomes System zur Identifikation, Bewertung, zum Sniping und zur Monetarisierung von abgelaufenen Domains mit historischem SEO-Kapital.

## Architektur & 5-Phasen-Workflow

```
[Phase 1: Reconnaissance]
  ├── ScrapingBee / ExpiredDomains.net Scraper
  ├── Broken Link Crawler (Target Website Scan)
  └── Crossdomain XML & Trust Relationship Inspector
          │
          ▼
[Phase 2: Filtering & Spam Quality Gates]
  ├── DP (Referring Domains) vs. BL (Total Backlinks) Ratio
  ├── ABY (Wayback Machine Birth Year) & ACR (Archive Crawl Records)
  └── DNS NXDOMAIN Resolution & Active Redirection Filter
          │
          ▼
[Phase 3: Valuation & Appraisal]
  ├── EstiBot API Valuation Engine (Search Volume, CPC, Trademark Risk)
  ├── DomainFlip Machine-Learning / Regression Scoring
  └── Affiliate/pSEO Topical Authority Matching
          │
          ▼
[Phase 4: Drop Catching & Sniping Engine]
  ├── Fast WHOIS / DNS Poller
  ├── High-Frequency EPP Drop-Catch Runner (Port 700 / Registrar Direct)
  └── DropCatch.com API & Backorder Integration (UNMI / Develanet)
          │
          ▼
[Phase 5: Re-Deployment & pSEO Engine]
  ├── Historical Content & URL Reconstruction (Wayback Mapping)
  ├── 301 / Wildcard Authority Preservation
  └── Programmatic SEO Landing Page Generator (Schema.org, JSON-LD, Affiliate CRO)
```

## Projektstruktur

- `core/`: Kernmodule für Scouting, Valuation, Sniping und Content-Generierung.
- `api/`: REST & CLI Schnittstellen für Agenten- und MCP-Aufrufe.
- `mcp/`: Lokale MCP-Server Definitionen für nahtlose Tool-Bereitstellung.
- `skills/`: Automatisierte Workflows und Prompts für Affiliate-pSEO.
- `data/`: SQLite/JSON Speicher für gecrawlte Leads, Scores und Catch-Logs.
- `tests/`: Unit- und Integrationstests (inkl. Mock-EPP- und DNS-Server).
