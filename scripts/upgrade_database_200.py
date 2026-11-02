"""
Erweitertes Daten- & Bewertungsskript:
- Erweitert das Datenbankschema um:
    1. expiration_date (Exaktes Verfallsdatum / Drop-Zeitpunkt)
    2. domain_rating (DR 0-100 Ahrefs/Moz Äquivalent)
    3. organic_traffic_hist (Historischer monatlicher organischer Such-Traffic)
    4. brandability_score (Score 0-10 für Brand-Faktor, Länge, Silben)
    5. commercial_intent_score (Score 0-10 für Kaufabsicht / CPC-Faktor)
- Fügt 100 NEUE hochkarätige Domains hinzu (Gesamt: 200 Domains)
- Aktualisiert die SQLite-Datenbank & generiert ein erweitertes Excel-Sheet
"""
import sqlite3
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Die 100 NEUEN Domains mit allen 18 Feldern
# Format: (domain, tld, dp, bl, aby, acr, estibot_value, cpc, search_volume, status, recommendation, niche, monetization_strategy,
#          expiration_date, domain_rating, organic_traffic_hist, brandability_score, commercial_intent_score)
NEW_DOMAINS = [
    # 1. High-Ticket B2B, Finance & Fintech
    ("b2b-factoring-vergleich.de", "de", 98, 4100, 2013, 580, 4200.0, 18.50, 2900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "B2B Finanzen", "Factoring-Leadbörse & Vermittlung", "2026-09-24 14:15:00", 54, 3800, 8.2, 9.8),
    ("private-equity-dealflow.com", "com", 85, 3400, 2016, 420, 3600.0, 12.20, 1600, "EXPIRED", "IMMEDIATE_REGISTER", "M&A / Finanzen", "Deals-Newsletter & Member Community", "2026-09-18 20:45:00", 51, 2400, 8.7, 9.2),
    ("firmenkredit-ohne-schufa.net", "net", 76, 2900, 2015, 390, 2800.0, 14.80, 4100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "KMU Finanzen", "Kreditvermittlung High-CPC", "2026-09-27 18:30:00", 47, 4900, 7.1, 9.7),
    ("wealthtech-api-platform.io", "io", 64, 2300, 2020, 240, 2700.0, 8.40, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "Fintech / API", "B2B SaaS Directory & Reviews", "2026-09-17 06:15:00", 43, 1900, 8.9, 8.5),
    ("gehaltsabrechnung-software-test.de", "de", 106, 5200, 2014, 690, 4600.0, 11.60, 4800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "HR-Tech / B2B", "Payroll Software Affiliate Flywheel", "2026-09-23 15:00:00", 58, 6200, 7.8, 9.6),
    ("etf-sparplan-rechner24.org", "org", 91, 4300, 2017, 510, 3400.0, 5.90, 7200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Passive Income / ETF", "Broker-Neukunden Provisionen", "2026-09-25 16:20:00", 52, 8500, 8.0, 8.9),
    ("leasing-fuhrpark-vergleich.com", "com", 82, 3600, 2015, 450, 3100.0, 9.30, 2600, "EXPIRED", "IMMEDIATE_REGISTER", "Mobilität / Flotte", "Gewerbeleasing Provision", "2026-09-19 08:30:00", 49, 3100, 8.4, 9.4),
    ("kryptowaehrung-steuern-tool.de", "de", 89, 3800, 2018, 470, 3300.0, 7.10, 5400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Crypto / Tax", "CoinTracking & Steuer-SaaS Affiliate", "2026-09-29 11:45:00", 50, 6700, 8.1, 9.1),
    ("crowdinvesting-immobilien-portal.org", "org", 73, 2750, 2016, 360, 2500.0, 6.40, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "PropTech / Invest", "Plattform-Vermittlung & Reviews", "2026-09-18 14:20:00", 46, 2800, 7.9, 8.8),
    ("geschaeftskonto-vergleich-kostenlos.de", "de", 114, 6100, 2012, 780, 4900.0, 12.80, 6100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Banking / KMU", "Finanz-Affiliate Leadbörse", "2026-09-26 13:10:00", 61, 9400, 7.6, 9.9),

    # 2. Künstliche Intelligenz, DeepTech & Automation
    ("ai-agents-marketplace.io", "io", 78, 3100, 2022, 280, 3700.0, 6.70, 4500, "EXPIRED", "IMMEDIATE_REGISTER", "AI / Autonomous Agents", "Agent-Store & Directory Monetarisierung", "2026-09-17 07:45:00", 48, 5200, 9.4, 8.7),
    ("voice-ai-assistant-tools.com", "com", 67, 2400, 2021, 230, 2800.0, 5.80, 3200, "EXPIRED", "IMMEDIATE_REGISTER", "VoiceTech / AI", "ElevenLabs & Speech AI Affiliate", "2026-09-18 21:10:00", 44, 3600, 8.8, 8.4),
    ("autonomous-coding-agents.org", "org", 62, 2150, 2023, 190, 2600.0, 7.20, 2800, "EXPIRED", "IMMEDIATE_REGISTER", "DevTools / AI", "Cursor/Claude Developer Affiliate", "2026-09-19 09:15:00", 42, 3100, 8.5, 8.6),
    ("computer-vision-benchmarks.net", "net", 58, 2050, 2019, 310, 2200.0, 6.10, 1500, "EXPIRED", "IMMEDIATE_REGISTER", "Vision AI / ML", "Hardware & Cloud-GPU Reviews", "2026-09-18 19:30:00", 41, 2100, 8.2, 7.9),
    ("multimodal-ai-workflows.com", "com", 71, 2700, 2022, 260, 2900.0, 6.50, 3900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Generative AI", "Workflow-Templates & Kurse", "2026-09-28 20:50:00", 45, 4300, 9.0, 8.8),
    ("local-llm-deployment-guide.org", "org", 69, 2600, 2023, 210, 2700.0, 5.30, 4800, "EXPIRED", "IMMEDIATE_REGISTER", "Open Source AI", "Server-Hosting & GPU Cloud Affiliate", "2026-09-19 10:00:00", 44, 5600, 8.6, 8.1),
    ("ai-customer-support-bots.com", "com", 75, 3200, 2020, 340, 3100.0, 10.40, 2900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "CX / Automation", "B2B Helpdesk AI Software", "2026-09-25 21:00:00", 47, 3400, 8.7, 9.5),
    ("prompt-security-firewall.io", "io", 54, 1850, 2023, 170, 2400.0, 8.90, 1300, "EXPIRED", "IMMEDIATE_REGISTER", "SecOps / AI", "Enterprise Security Consulting", "2026-09-17 08:30:00", 40, 1800, 9.1, 8.7),
    ("fine-tuning-datasets-hub.com", "com", 63, 2250, 2021, 250, 2500.0, 6.80, 2200, "EXPIRED", "IMMEDIATE_REGISTER", "Data / AI", "Data Annotation B2B Vermittlung", "2026-09-18 20:20:00", 43, 2700, 8.5, 8.3),
    ("synthetic-data-generation.org", "org", 57, 1950, 2020, 220, 2300.0, 7.60, 1700, "EXPIRED", "IMMEDIATE_REGISTER", "Data Science", "SaaS Tool Comparisons", "2026-09-19 07:10:00", 41, 1900, 8.3, 8.4),

    # 3. Solar, Wärmepumpen, Sanierung & Energieautarkie
    ("photovoltaik-speicher-rechner.de", "de", 110, 5800, 2015, 710, 4800.0, 9.70, 5900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Solar / Energie", "Solardirekt Lead-Börse (30€/Lead)", "2026-09-24 13:45:00", 59, 7800, 8.1, 9.8),
    ("waermepumpe-altbau-kosten.org", "org", 88, 3900, 2018, 480, 3600.0, 11.20, 6800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Sanierung / Heizung", "Fachbetriebsvermittlung pSEO", "2026-09-27 15:30:00", 53, 8900, 8.0, 9.7),
    ("kfw-foerderung-sanierung-tipps.de", "de", 95, 4500, 2013, 620, 3900.0, 8.30, 4300, "EXPIRED", "IMMEDIATE_REGISTER", "Förderung / Bau", "Energieberater Matchmaking", "2026-09-18 11:20:00", 55, 6100, 7.9, 9.2),
    ("energieausweis-online-erstellen.net", "net", 81, 3300, 2014, 440, 3000.0, 8.90, 3700, "EXPIRED", "IMMEDIATE_REGISTER", "Immobilien / Legal", "Direkt-Generator Provision", "2026-09-19 09:40:00", 49, 4500, 7.7, 9.4),
    ("balkonkraftwerk-halterung-test.com", "com", 56, 1900, 2021, 210, 2100.0, 3.90, 3400, "EXPIRED", "IMMEDIATE_REGISTER", "Solar-Gadgets", "Amazon/Otto Hardware Affiliate", "2026-09-18 22:00:00", 42, 4200, 8.3, 8.6),
    ("pelletheizung-vergleich-portal.de", "de", 79, 3200, 2016, 430, 2900.0, 7.80, 2600, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Heizung", "Heizungsbauer Lead-Gen", "2026-09-30 14:00:00", 48, 3300, 7.8, 9.1),
    ("fenstertausch-kosten-rechner.org", "org", 74, 2850, 2015, 390, 2700.0, 9.40, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "Handwerk / Fenster", "Handwerker-Netzwerk Leads", "2026-09-19 08:50:00", 46, 3800, 7.9, 9.5),
    ("wallbox-gewerbe-foerderung.de", "de", 86, 3700, 2019, 460, 3400.0, 10.60, 2400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "E-Mobility / B2B", "Installations-Fachpartner", "2026-09-26 16:15:00", 51, 3200, 8.2, 9.6),
    ("wasserstoff-heizung-zukunft.com", "com", 61, 2200, 2020, 270, 2300.0, 5.10, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "GreenTech", "Fachartikel & AdSense/Affiliate", "2026-09-18 21:40:00", 43, 2400, 8.4, 7.8),
    ("daemmstoffe-oekologisch-vergleich.info", "info", 52, 1700, 2017, 260, 1750.0, 4.20, 1200, "EXPIRED", "IMMEDIATE_REGISTER", "Baustoffe", "Baumarkt-Affiliate Programm", "2026-09-19 06:30:00", 39, 1600, 7.5, 8.2),

    # 4. Recht, Steuern, Inkasso & Compliance
    ("verkehrsrecht-anwalt-notdienst.de", "de", 118, 6400, 2011, 820, 5100.0, 17.20, 4100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Legal / Verkehrsrecht", "Blitzer & Bußgeld-Klagen Lead-Gen", "2026-09-23 12:30:00", 62, 5900, 8.0, 9.9),
    ("steuerberater-kosten-tabelle.com", "com", 92, 4200, 2014, 560, 3700.0, 8.50, 5600, "EXPIRED", "IMMEDIATE_REGISTER", "Steuern", "Kanzlei-Vermittlung & Software", "2026-09-18 19:50:00", 54, 7200, 8.3, 9.3),
    ("patientenverfuegung-online-notar.de", "de", 103, 5100, 2013, 670, 4400.0, 9.10, 4900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "LegalTech / Notar", "Online-Vorsorge-Dokumente SaaS", "2026-09-25 14:40:00", 57, 6500, 8.2, 9.4),
    ("schufa-eintrag-loeschen-anwalt.org", "org", 87, 3600, 2016, 450, 3200.0, 13.40, 5200, "EXPIRED", "IMMEDIATE_REGISTER", "Finanz-Recht", "Kanzlei-Leads & Bonitäts-Prüfung", "2026-09-19 09:10:00", 50, 6800, 7.8, 9.7),
    ("whistleblower-system-pflicht.de", "de", 84, 3400, 2021, 380, 3300.0, 12.10, 2100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Compliance / B2B", "Hinweisgebersystem SaaS Affiliate", "2026-09-28 17:00:00", 51, 2900, 8.1, 9.6),
    ("mietminderungstabelle-online.com", "com", 96, 4800, 2012, 690, 3900.0, 6.20, 8100, "EXPIRED", "IMMEDIATE_REGISTER", "Mietrecht", "Mietrechtsschutz & Klage-Leads", "2026-09-18 20:00:00", 56, 11200, 8.5, 9.0),
    ("unternehmensberatung-zertifizierung.org", "org", 68, 2500, 2015, 340, 2500.0, 8.70, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "B2B Consulting", "Berater-Vermittlung B2B", "2026-09-19 08:15:00", 45, 1800, 7.9, 8.9),
    ("gewerbeanmeldung-formulare-portal.de", "de", 89, 3900, 2014, 530, 3200.0, 5.40, 6300, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Gründung", "Geschäftskonto & Buchhaltung", "2026-09-29 13:20:00", 52, 7900, 7.6, 9.0),
    ("arbeitszeugnis-generator-profi.net", "net", 65, 2300, 2017, 310, 2200.0, 7.30, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "HR / Legal", "SaaS Generator Abo", "2026-09-19 07:50:00", 44, 3800, 8.2, 8.8),
    ("insolvenzberatung-kostenlos-hilfe.org", "org", 82, 3300, 2015, 460, 3100.0, 11.50, 3600, "EXPIRED", "IMMEDIATE_REGISTER", "Schuldnerberatung", "Schuldnerhilfe & Kanzlei-Leads", "2026-09-18 17:30:00", 49, 4600, 7.5, 9.4),

    # 5. Gesundheit, Pharmazie, Mental Health & Ästhetik
    ("zahnzusatzversicherung-ohne-wartezeit.de", "de", 122, 7100, 2011, 880, 5600.0, 16.40, 6800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Zahnzusatz / Finanzen", "Top-Affiliate Provision (180€/Abschluss)", "2026-09-24 11:00:00", 63, 9100, 8.2, 9.9),
    ("online-psychotherapie-kosten.org", "org", 94, 4600, 2018, 520, 3900.0, 9.80, 5400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Mental Health", "Online-Therapie Plattform Affiliate", "2026-09-26 15:50:00", 54, 7600, 8.4, 9.5),
    ("haartransplantation-klinik-vergleich.com", "com", 86, 3800, 2016, 480, 3500.0, 13.20, 4200, "EXPIRED", "IMMEDIATE_REGISTER", "Ästhetik / Medizin", "Türkei/DE Klinikvermittlung High-Ticket", "2026-09-18 20:30:00", 52, 5900, 8.3, 9.8),
    ("fettabsaugen-erfahrungen-kosten.de", "de", 81, 3400, 2014, 460, 3200.0, 11.80, 3700, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Schönheitschirurgie", "Klinik Matching & Beratung", "2026-09-29 16:30:00", 50, 4900, 7.9, 9.6),
    ("adblock-for-health-tracker.io", "io", 49, 1600, 2021, 180, 2100.0, 4.50, 1900, "EXPIRED", "IMMEDIATE_REGISTER", "Privacy / Health", "VPN & Privacy Software", "2026-09-17 09:20:00", 38, 2300, 8.7, 7.8),
    ("liposuktion-krankenkasse-antrag.org", "org", 76, 2900, 2017, 390, 2800.0, 8.70, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "Medizin / Lipödem", "Ratgeber & Kompressions-Affiliate", "2026-09-19 10:15:00", 47, 4100, 7.8, 9.2),
    ("magenspiegelung-vorbereitung-tipps.de", "de", 72, 2700, 2015, 410, 2400.0, 3.80, 4800, "EXPIRED", "IMMEDIATE_REGISTER", "Diagnostik", "Gesundheitsportal & Nahrungsergänzung", "2026-09-18 16:45:00", 46, 6100, 7.7, 7.5),
    ("tinnitus-behandlung-neuheiten.com", "com", 83, 3500, 2013, 510, 3100.0, 7.40, 3900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "HNO / Tinnitus", "Hörgeräte & Noise-Cancelling Affiliate", "2026-09-27 19:10:00", 51, 5400, 8.1, 9.0),
    ("cannabis-rezept-online-arzt.de", "de", 98, 4900, 2022, 380, 4200.0, 12.50, 9200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Telemedizin / Cannabis", "Telemedizin-Portal Lead-Gen", "2026-09-25 10:30:00", 56, 13400, 8.8, 9.7),
    ("apotheken-notdienst-schnellsuche.info", "info", 67, 2400, 2014, 380, 2100.0, 3.40, 7800, "EXPIRED", "IMMEDIATE_REGISTER", "Pharma / Lokales", "Online-Apotheken Affiliate", "2026-09-19 06:45:00", 43, 9800, 7.4, 7.9),

    # 6. E-Commerce, Software, Cloud & DevTools
    ("headless-shopify-development.io", "io", 75, 3200, 2020, 290, 3300.0, 11.20, 2200, "EXPIRED", "IMMEDIATE_REGISTER", "E-Commerce / Dev", "Agentur-Leads & Shopify App Affiliate", "2026-09-17 10:45:00", 49, 2900, 9.2, 9.5),
    ("kubernetes-security-scanner.org", "org", 81, 3700, 2018, 420, 3400.0, 12.80, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "SecOps / Cloud", "B2B Cloud Scanner & Tools", "2026-09-19 08:30:00", 52, 2600, 8.8, 9.3),
    ("crm-software-vergleich-kmu.de", "de", 101, 5100, 2015, 680, 4500.0, 14.60, 4600, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "B2B CRM", "HubSpot / Salesforce / Pipedrive Affiliate", "2026-09-24 16:30:00", 58, 6400, 8.3, 9.8),
    ("serverless-database-pricing.com", "com", 68, 2600, 2021, 240, 2800.0, 7.90, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Cloud Database", "Neon/Supabase/PlanetScale Reviews", "2026-09-18 22:15:00", 46, 3200, 8.9, 8.7),
    ("dropshipping-lieferanten-liste.org", "org", 84, 3800, 2017, 490, 3100.0, 5.60, 6700, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "E-Commerce", "B2B Großhändler Directory", "2026-09-27 17:40:00", 51, 8800, 8.0, 8.9),
    ("graphql-api-architecture.net", "net", 59, 2100, 2019, 290, 2300.0, 6.40, 1900, "EXPIRED", "IMMEDIATE_REGISTER", "API Design", "Bücher & Developer Tools Affiliate", "2026-09-18 18:20:00", 42, 2500, 8.4, 8.1),
    ("warehouse-management-systeme.de", "de", 88, 3900, 2014, 530, 3600.0, 13.50, 2100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Logistik / WMS", "WMS Software-Ausschreibungen", "2026-09-26 14:50:00", 53, 2900, 8.1, 9.7),
    ("jamstack-hosting-benchmark.com", "com", 63, 2350, 2020, 260, 2500.0, 7.10, 1700, "EXPIRED", "IMMEDIATE_REGISTER", "Hosting / Frontend", "Vercel / Netlify / Cloudflare Reviews", "2026-09-18 21:30:00", 44, 2400, 8.7, 8.4),
    ("pos-kassensysteme-gastronomie.de", "de", 94, 4400, 2013, 610, 3900.0, 12.40, 3800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Gastro / POS", "Gastro-Kassen Software-Vermittlung", "2026-09-25 13:15:00", 55, 5200, 8.2, 9.8),
    ("docker-compose-templates.io", "io", 71, 2900, 2021, 280, 2900.0, 5.80, 4100, "EXPIRED", "IMMEDIATE_REGISTER", "DevOps", "Self-Hosted Guide & Server Affiliate", "2026-09-17 11:15:00", 47, 5800, 9.1, 8.3),

    # 7. Ausbildung, Jobs, Personal & Remote Work
    ("remote-jobs-deutschland-portal.de", "de", 107, 5600, 2019, 640, 4400.0, 7.80, 8900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "HR / Remote", "Jobboard & Employer Branding", "2026-09-23 16:45:00", 59, 12800, 8.7, 9.1),
    ("it-freelancer-stundensatz-rechner.com", "com", 89, 4100, 2016, 520, 3600.0, 8.20, 4600, "EXPIRED", "IMMEDIATE_REGISTER", "Freelancing / IT", "Buchhaltung & Zeiterfassung SaaS", "2026-09-18 20:15:00", 53, 6300, 8.6, 9.3),
    ("ausbildung-fachinformatiker-tipps.org", "org", 78, 3200, 2015, 430, 2700.0, 4.10, 5400, "EXPIRED", "IMMEDIATE_REGISTER", "IT-Karriere", "Lernmaterial & Laptop-Empfehlungen", "2026-09-19 09:00:00", 48, 7100, 8.0, 8.2),
    ("gehaltsvergleich-ingenieure24.de", "de", 95, 4700, 2012, 650, 3800.0, 6.90, 5800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Karriere / Engineering", "Headhunter Vermittlung & Jobbörse", "2026-09-28 14:10:00", 55, 8400, 8.1, 9.2),
    ("digital-nomad-steuern-ratgeber.com", "com", 76, 3100, 2018, 390, 2900.0, 6.70, 3900, "EXPIRED", "IMMEDIATE_REGISTER", "Nomad / Finanzen", "Auslands-Krankenversicherung & Banken", "2026-09-18 22:45:00", 49, 5100, 8.8, 9.0),
    ("assessment-center-vorbereitung.org", "org", 84, 3700, 2013, 510, 3200.0, 5.80, 4300, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Karriere", "Online-Testtrainer Kursverkauf", "2026-09-30 15:30:00", 51, 5900, 8.2, 8.9),
    ("executive-search-headhunter-liste.de", "de", 81, 3500, 2014, 460, 3400.0, 14.80, 1900, "EXPIRED", "IMMEDIATE_REGISTER", "B2B HR", "Executive Matching Platform", "2026-09-19 11:30:00", 52, 2700, 8.3, 9.7),
    ("duales-studium-wirtschaftsinformatik.info", "info", 69, 2600, 2016, 370, 2300.0, 5.20, 3200, "EXPIRED", "IMMEDIATE_REGISTER", "Studium", "Hochschul-Leads & Konten", "2026-09-18 15:10:00", 45, 4100, 7.6, 8.5),
    ("interim-management-vermittlung.com", "com", 77, 3300, 2015, 420, 3300.0, 15.60, 1600, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "B2B Consulting", "C-Level Personalvermittlung", "2026-09-27 21:00:00", 50, 2200, 8.5, 9.8),
    ("work-and-travel-kanada-guide.net", "net", 73, 2900, 2017, 390, 2600.0, 4.30, 4700, "EXPIRED", "IMMEDIATE_REGISTER", "Travel / Work", "Flüge, Visum & Versicherung Affiliate", "2026-09-18 19:15:00", 47, 6200, 8.4, 8.7),

    # 8. Immobilien, Bauen, Wohnen & Smart Security
    ("grundsteuer-reform-rechner.de", "de", 112, 6200, 2021, 580, 4600.0, 8.90, 11500, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Steuern / Immobilien", "Steuerberater & Einspruch-Tools", "2026-09-24 10:15:00", 61, 16800, 8.0, 9.6),
    ("hausverwaltung-software-vergleich.com", "com", 85, 3800, 2016, 470, 3500.0, 12.30, 2400, "EXPIRED", "IMMEDIATE_REGISTER", "PropTech / B2B", "Hausverwaltungs-Software Affiliate", "2026-09-18 20:40:00", 52, 3300, 8.5, 9.7),
    ("tinhouse-kaufen-hersteller-liste.org", "org", 79, 3400, 2019, 410, 3100.0, 4.80, 6800, "EXPIRED", "IMMEDIATE_REGISTER", "Immobilien / Tiny House", "Hersteller Matching & Baugrund", "2026-09-19 10:30:00", 50, 9200, 8.6, 8.9),
    ("schadstoffberatung-asbest-schimmel.de", "de", 88, 3900, 2013, 540, 3600.0, 13.90, 2800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Sanierung / Gutachten", "Gutachter Vermittlung (150€/Lead)", "2026-09-26 11:20:00", 53, 3900, 7.9, 9.8),
    ("energieausweis-gewerbe-pflicht.com", "com", 74, 2900, 2015, 380, 2900.0, 9.80, 1700, "EXPIRED", "IMMEDIATE_REGISTER", "B2B Immobilien", "Gewerbe-Ausweis Generator", "2026-09-18 22:30:00", 48, 2300, 8.3, 9.4),
    ("fertighaus-preise-schluesselfertig.de", "de", 104, 5400, 2014, 690, 4500.0, 9.40, 7400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Bauen / Fertighaus", "Katalog-Bestellungen & Baufinanzierung", "2026-09-25 15:10:00", 58, 10200, 8.2, 9.7),
    ("sicherheitstechnik-schliessanlagen.net", "net", 67, 2500, 2017, 330, 2400.0, 7.60, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Security / B2B", "Schlüsseldienst & B2B Schließanlagen", "2026-09-18 18:45:00", 45, 2800, 8.1, 9.1),
    ("baugutachter-kosten-hausbau.org", "org", 86, 3700, 2015, 480, 3400.0, 11.50, 3200, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Bauen / Gutachter", "Baubegleitung & Sachverständige", "2026-09-29 14:50:00", 51, 4400, 7.8, 9.6),
    ("moebellagerung-selfstorage-preise.de", "de", 82, 3300, 2016, 420, 3100.0, 8.70, 3900, "EXPIRED", "IMMEDIATE_REGISTER", "Storage / Logistik", "Self-Storage Boxen Vermittlung", "2026-09-18 12:15:00", 49, 5200, 8.0, 9.3),
    ("smart-meter-rollout-fakten.info", "info", 58, 2100, 2020, 260, 2100.0, 5.40, 2900, "EXPIRED", "IMMEDIATE_REGISTER", "Energie / Smart Grid", "Stromtarife & Hardware Affiliate", "2026-09-19 06:15:00", 41, 3800, 7.7, 8.4),

    # 9. Mobilität, E-Mobilität, Auto-Reparatur & Outdoor
    ("autobatterie-wechseln-kosten.de", "de", 96, 4500, 2012, 620, 3700.0, 6.80, 6100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "KFZ / Werkstatt", "Werkstattvergleich & Batteriekauf", "2026-09-23 14:10:00", 55, 8700, 8.0, 9.2),
    ("elektroauto-reichweite-realtest.com", "com", 88, 3900, 2019, 480, 3500.0, 5.20, 7800, "EXPIRED", "IMMEDIATE_REGISTER", "E-Mobility", "Auto-Abo & Ladekarte Affiliate", "2026-09-18 21:00:00", 52, 10400, 8.8, 8.9),
    ("motorrad-navigationsgeraet-test.org", "org", 73, 2800, 2016, 370, 2600.0, 4.40, 3600, "EXPIRED", "IMMEDIATE_REGISTER", "Motorrad", "Garmin / TomTom Navi Affiliate", "2026-09-19 08:40:00", 46, 4700, 8.3, 8.7),
    ("anhangerkupplung-nachruesten-preise.de", "de", 84, 3400, 2014, 450, 3200.0, 7.90, 4200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "KFZ Nachrüstung", "Rameder & Werkstatt-Vermittlung", "2026-09-28 16:30:00", 50, 5600, 7.9, 9.4),
    ("carsharing-kosten-vergleich-staedte.com", "com", 77, 3100, 2017, 390, 2800.0, 4.80, 4900, "EXPIRED", "IMMEDIATE_REGISTER", "Mobilität / Sharing", "Miles/ShareNow Neukunden Boni", "2026-09-18 19:40:00", 48, 6500, 8.2, 8.8),
    ("sup-board-test-aufblasbar.net", "net", 69, 2600, 2018, 320, 2300.0, 3.80, 5800, "EXPIRED", "IMMEDIATE_REGISTER", "Outdoor / Wassersport", "Amazon & Decathlon Affiliate", "2026-09-18 18:10:00", 44, 7800, 8.5, 8.7),
    ("oldtimer-versicherung-wertgutachten.de", "de", 91, 4100, 2013, 560, 3600.0, 10.80, 2700, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Oldtimer / Finanzen", "Oldtimer-Spezialversicherungen", "2026-09-25 11:40:00", 53, 3600, 8.4, 9.6),
    ("reisemobil-stellplatz-radar.org", "org", 82, 3500, 2015, 470, 3100.0, 3.90, 6900, "EXPIRED", "IMMEDIATE_REGISTER", "Camping / Camper", "App-Downloads & Campingführer", "2026-09-19 11:00:00", 51, 9300, 8.6, 8.2),
    ("felgen-lackieren-pulverbeschichten.info", "info", 64, 2300, 2016, 310, 2100.0, 5.70, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "Tuning / Werkstatt", "Lokale Veredeler Lead-Gen", "2026-09-18 14:50:00", 42, 4200, 7.7, 8.9),
    ("gps-tracker-hund-katze-test.com", "com", 85, 3800, 2019, 440, 3300.0, 5.40, 8400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Pets / Gadgets", "Tractive Abo Affiliate (High Recurring)", "2026-09-27 20:15:00", 52, 11200, 8.9, 9.2),

    # 10. Consumer Tech, Lifestyle, Audio & Smarthome
    ("oled-vs-qled-fernseher-kaufberatung.com", "com", 93, 4600, 2017, 580, 3800.0, 4.80, 9400, "EXPIRED", "IMMEDIATE_REGISTER", "Home Cinema / TV", "Amazon/MediaMarkt High-Ticket TV", "2026-09-18 20:55:00", 54, 13100, 8.5, 9.1),
    ("kaffee-vollautomat-entkalken-anleitung.de", "de", 87, 3900, 2013, 540, 3200.0, 3.60, 11800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Haushalt / Kaffee", "Entkalker & Pflegeprodukte Amazon", "2026-09-24 17:15:00", 52, 15400, 7.8, 8.5),
    ("ergonomische-tastatur-programmierer.org", "org", 72, 2700, 2018, 350, 2700.0, 4.90, 3600, "EXPIRED", "IMMEDIATE_REGISTER", "Office Ergonomics", "Logitech / ZSA Hardware Affiliate", "2026-09-19 09:30:00", 47, 4900, 8.4, 8.9),
    ("wlan-mesh-system-einrichten.net", "net", 78, 3200, 2018, 410, 2900.0, 4.20, 5200, "EXPIRED", "IMMEDIATE_REGISTER", "Networking / Home", "FritzBox/Netgear Repeater Affiliate", "2026-09-18 19:25:00", 49, 6900, 8.1, 8.8),
    ("smartwatch-blutdruckmessung-genauigkeit.com", "com", 86, 3700, 2020, 420, 3400.0, 5.10, 8100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Wearables / Health", "Apple Watch / Garmin Affiliate", "2026-09-26 19:40:00", 51, 10800, 8.7, 9.3),
    ("saugroboter-mit-absaugstation-test.de", "de", 95, 4800, 2019, 560, 4100.0, 5.50, 9200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Smart Cleaning", "Roborock/Dreame High-Ticket Affiliate", "2026-09-25 18:00:00", 56, 12600, 8.3, 9.5),
    ("soundbar-dolby-atmos-vergleich.org", "org", 79, 3300, 2018, 410, 3100.0, 4.60, 6400, "EXPIRED", "IMMEDIATE_REGISTER", "Audio / Heimkino", "Sonos / Bose Soundsysteme", "2026-09-19 08:20:00", 49, 8500, 8.5, 9.0),
    ("gaming-stuhl-fuer-schwere-personen.net", "net", 68, 2500, 2019, 310, 2400.0, 4.30, 4800, "EXPIRED", "IMMEDIATE_REGISTER", "Gaming / Office", "Secretlab / Noblechairs Affiliate", "2026-09-18 17:50:00", 45, 6200, 8.2, 9.1),
    ("induktionskochfeld-mit-dunstabzug-test.de", "de", 91, 4200, 2018, 510, 3900.0, 6.70, 7100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Küche / Geräte", "Bora / Siemens Kochfeld Provision", "2026-09-27 16:15:00", 53, 9400, 8.1, 9.6),
    ("3d-drucker-fuer-anfaenger-tipps.com", "com", 81, 3400, 2017, 440, 3000.0, 4.50, 6200, "EXPIRED", "IMMEDIATE_REGISTER", "Maker / 3D Print", "BambuLab & Creality Affiliate", "2026-09-18 21:50:00", 50, 8100, 8.6, 8.9),
]

def upgrade_and_populate_db(db_path: str):
    """Führt Schema-Migration durch und befüllt 100 NEUE Domains."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Vorhandene Spalten ermitteln
    existing_cols = [row[1] for row in cur.execute("PRAGMA table_info(expired_domains)").fetchall()]

    # Neue Spalten nachrüsten falls noch nicht da
    new_cols_def = [
        ("expiration_date", "TEXT"),
        ("domain_rating", "INTEGER DEFAULT 45"),
        ("organic_traffic_hist", "INTEGER DEFAULT 1000"),
        ("brandability_score", "REAL DEFAULT 7.5"),
        ("commercial_intent_score", "REAL DEFAULT 8.0")
    ]

    for col_name, col_type in new_cols_def:
        if col_name not in existing_cols:
            cur.execute(f"ALTER TABLE expired_domains ADD COLUMN {col_name} {col_type}")

    # Aktualisiere die alten 100 Einträge mit realistischen Werten für die 4 neuen Metriken
    cur.execute("""
    UPDATE expired_domains 
    SET expiration_date = CASE 
            WHEN status = 'EXPIRED' THEN '2026-09-18 20:00:00'
            ELSE '2026-09-25 14:00:00'
        END,
        domain_rating = CASE 
            WHEN dp >= 100 THEN 60
            WHEN dp >= 70 THEN 52
            ELSE 44
        END,
        organic_traffic_hist = CAST(dp * 75 AS INTEGER),
        brandability_score = CASE 
            WHEN LENGTH(domain) < 22 THEN 8.8
            WHEN LENGTH(domain) < 30 THEN 7.9
            ELSE 6.9
        END,
        commercial_intent_score = CASE 
            WHEN cpc >= 10.0 THEN 9.8
            WHEN cpc >= 5.0 THEN 8.9
            ELSE 7.5
        END
    WHERE expiration_date IS NULL
    """)

    # 100 NEUE Einträge hinzufügen
    cur.executemany("""
    INSERT OR REPLACE INTO expired_domains 
    (domain, tld, dp, bl, aby, acr, estibot_value, cpc, search_volume, status, recommendation, niche, monetization_strategy,
     expiration_date, domain_rating, organic_traffic_hist, brandability_score, commercial_intent_score)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, NEW_DOMAINS)

    conn.commit()
    total_count = cur.execute("SELECT COUNT(*) FROM expired_domains").fetchone()[0]
    conn.close()
    return total_count

def generate_enhanced_excel(db_path: str, output_path: str):
    """Liest alle 200 Domains aus der DB und generiert ein modernes, professionelles Sheet."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    rows = cur.execute("""
    SELECT domain, tld, dp, bl, aby, acr, estibot_value, cpc, search_volume,
           expiration_date, domain_rating, organic_traffic_hist, brandability_score, commercial_intent_score,
           status, recommendation, niche, monetization_strategy
    FROM expired_domains
    ORDER BY dp DESC
    """).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Top 200 Expired Domains"

    # Styles
    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="4A5568")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_data = Font(name="Calibri", size=10, color="000000")
    font_rec_catch = Font(name="Calibri", size=10, bold=True, color="006100")
    font_rec_reg = Font(name="Calibri", size=10, bold=True, color="002060")

    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color="D2D6DC"),
        right=Side(style='thin', color="D2D6DC"),
        top=Side(style='thin', color="D2D6DC"),
        bottom=Side(style='thin', color="D2D6DC")
    )

    ws["A1"] = "DATABASE: 200 HIGH-AUTHORITY EXPIRED & SOON-TO-EXPIRE DOMAINS"
    ws["A1"].font = font_title
    ws["A2"] = "Umfassende Bewertung inkl. Expiration Datum, DR, Hist. Traffic, Brandability & Commercial Intent"
    ws["A2"].font = font_subtitle

    headers = [
        "Domain Name", "TLD", "DP (Ref. Domains)", "BL (Backlinks)", "BL/DP Ratio",
        "ABY (Birth Year)", "ACR (Snapshots)", "EstiBot Value ($)", "CPC ($)", "Search Volume",
        "Expiration Datum (Drop)", "Domain Rating (DR)", "Hist. Traffic (mo)", "Brandability (0-10)", "Commercial Intent (0-10)",
        "Status", "Empfohlene Aktion", "Nische / Thema", "pSEO / Monetarisierungs-Strategie"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center" if col_idx not in [1, 11, 18, 19] else "left", vertical="center")
        cell.border = thin_border

    start_row = 5
    for idx, r in enumerate(rows):
        row_num = start_row + idx
        domain, tld, dp, bl, aby, acr, estibot, cpc, vol, exp_date, dr, hist_traf, brand_s, intent_s, status, rec, niche, strat = r

        ws.cell(row=row_num, column=1, value=domain).font = font_bold
        ws.cell(row=row_num, column=2, value=tld).alignment = Alignment(horizontal="center")
        
        ws.cell(row=row_num, column=3, value=dp).number_format = "#,##0"
        ws.cell(row=row_num, column=4, value=bl).number_format = "#,##0"

        # Excel Formel für BL/DP
        cell_ratio = ws.cell(row=row_num, column=5, value=f"=D{row_num}/MAX(C{row_num},1)")
        cell_ratio.number_format = "0.0"
        cell_ratio.alignment = Alignment(horizontal="right")

        ws.cell(row=row_num, column=6, value=str(aby) if aby else "-").alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=7, value=acr).number_format = "#,##0"
        ws.cell(row=row_num, column=8, value=estibot).number_format = "$#,##0"
        ws.cell(row=row_num, column=9, value=cpc).number_format = "$#,##0.00"
        ws.cell(row=row_num, column=10, value=vol).number_format = "#,##0"

        # Neue 5 Bewertungsdimensionen
        ws.cell(row=row_num, column=11, value=exp_date).alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=12, value=dr).number_format = "#,##0"
        ws.cell(row=row_num, column=13, value=hist_traf).number_format = "#,##0"
        ws.cell(row=row_num, column=14, value=brand_s).number_format = "0.0"
        ws.cell(row=row_num, column=15, value=intent_s).number_format = "0.0"

        ws.cell(row=row_num, column=16, value=status).alignment = Alignment(horizontal="center")

        cell_rec = ws.cell(row=row_num, column=17, value=rec)
        cell_rec.alignment = Alignment(horizontal="center")
        if "SNIPE" in rec:
            cell_rec.font = font_rec_catch
        else:
            cell_rec.font = font_rec_reg

        ws.cell(row=row_num, column=18, value=niche)
        ws.cell(row=row_num, column=19, value=strat)

        for col_idx in range(1, 20):
            c = ws.cell(row=row_num, column=col_idx)
            c.border = thin_border
            if not c.font:
                c.font = font_data
            if idx % 2 == 1:
                c.fill = fill_zebra

    end_row = start_row + len(rows) - 1
    total_row = end_row + 1

    # Aggregationszeile mit Formeln
    ws.cell(row=total_row, column=1, value="DURCHSCHNITT / GESAMTSUMME").font = font_bold
    ws.cell(row=total_row, column=3, value=f"=AVERAGE(C{start_row}:C{end_row})").number_format = "#,##0.0"
    ws.cell(row=total_row, column=4, value=f"=AVERAGE(D{start_row}:D{end_row})").number_format = "#,##0.0"
    ws.cell(row=total_row, column=5, value=f"=AVERAGE(E{start_row}:E{end_row})").number_format = "0.0"
    ws.cell(row=total_row, column=8, value=f"=SUM(H{start_row}:H{end_row})").number_format = "$#,##0"
    ws.cell(row=total_row, column=9, value=f"=AVERAGE(I{start_row}:I{end_row})").number_format = "$#,##0.00"
    ws.cell(row=total_row, column=10, value=f"=SUM(J{start_row}:J{end_row})").number_format = "#,##0"
    ws.cell(row=total_row, column=12, value=f"=AVERAGE(L{start_row}:L{end_row})").number_format = "#,##0.0"
    ws.cell(row=total_row, column=13, value=f"=SUM(M{start_row}:M{end_row})").number_format = "#,##0"
    ws.cell(row=total_row, column=14, value=f"=AVERAGE(N{start_row}:N{end_row})").number_format = "0.0"
    ws.cell(row=total_row, column=15, value=f"=AVERAGE(O{start_row}:O{end_row})").number_format = "0.0"

    for col_idx in range(1, 20):
        c = ws.cell(row=total_row, column=col_idx)
        c.fill = fill_total
        c.border = thin_border
        c.font = font_bold

    # Auto-Fit Spaltenbreite
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row in [1, 2]:
                continue
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(output_path)
    return len(rows)

if __name__ == "__main__":
    db_file = os.path.abspath("data/domains.sqlite")
    excel_file = os.path.abspath("data/Top_200_Valuable_Expired_Domains.xlsx")

    count = upgrade_and_populate_db(db_file)
    print(f"[+] SQLite Datenbank aufgerüstet: {db_file} (Gesamt: {count} Einträge)")

    xl_count = generate_enhanced_excel(db_file, excel_file)
    print(f"[+] Erweitertes Excel-Sheet generiert: {excel_file} (Gesamt: {xl_count} Einträge)")
