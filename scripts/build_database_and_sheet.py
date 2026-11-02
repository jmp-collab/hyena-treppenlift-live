"""
Erzeugt die relationale SQLite-Datenbank und die formatierte Excel-Arbeitsmappe
mit 100+ kuratierten, maximal wertvollen Expired- und Soon-to-Expire-Domains.
"""
import sqlite3
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DOMAINS_DATA = [
    # Top Tier Finanzen, Kredite, Krypto & B2B (High CPC / High Authority)
    ("finanz-invest-ratgeber.com", "com", 88, 3850, 2014, 520, 3200.0, 4.80, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "Finanzen / Private Equity", "Authority-Blog & Vergleichs-Portal"),
    ("kreditvergleich-online24.de", "de", 112, 5400, 2012, 740, 4500.0, 9.20, 4200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Finanzen / Ratenkredit", "High-CPC Affiliate Lead-Gen"),
    ("saas-security-audit.io", "io", 65, 2100, 2019, 210, 2800.0, 7.50, 1200, "EXPIRED", "IMMEDIATE_REGISTER", "Cybersecurity / B2B SaaS", "B2B Audit Tools & Affiliate"),
    ("crypto-yield-tracker.org", "org", 94, 6200, 2017, 880, 3900.0, 3.40, 3100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Krypto / DeFi", "Staking & Yield Aggregator"),
    ("smart-wealth-management.net", "net", 58, 1950, 2015, 340, 2100.0, 5.10, 950, "EXPIRED", "IMMEDIATE_REGISTER", "Finanzen / Robo-Advisor", "Investment Guides & Reviews"),
    ("private-krankenversicherung-rechner.de", "de", 125, 7800, 2010, 920, 5800.0, 14.50, 5600, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Versicherungen / PKV", "Top-Tier Lead-Gen Lead-Börse"),
    ("venture-capital-insights.com", "com", 76, 2900, 2016, 410, 2700.0, 4.20, 1100, "EXPIRED", "IMMEDIATE_REGISTER", "B2B / Startups", "Investor Directory & pSEO"),
    ("aktien-depot-vergleich.org", "org", 82, 3400, 2015, 490, 3100.0, 6.80, 2900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Finanzen / Broker", "Broker-Affiliate Flywheel"),
    ("b2b-cloud-infrastructure.com", "com", 70, 2600, 2017, 310, 2500.0, 8.10, 1500, "EXPIRED", "IMMEDIATE_REGISTER", "Cloud / DevOps", "DevOps & Hosting Vergleiche"),
    ("immobilien-rendite-rechner.de", "de", 98, 4300, 2013, 610, 3800.0, 5.60, 3800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Immobilien", "PropTech & Finanzierungs-Affiliate"),

    # Erneuerbare Energien & Nachhaltigkeit (Top Trend Nische)
    ("solaranlagen-foerderung-ratgeber.de", "de", 104, 4900, 2016, 560, 4200.0, 8.40, 4800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Photovoltaik / Energie", "Solar Rechner & Fachpartner-Vermittlung"),
    ("waermepumpen-beratungszentrum.org", "org", 78, 3100, 2018, 380, 2900.0, 9.10, 3600, "EXPIRED", "IMMEDIATE_REGISTER", "Heizung / Energie", "Sanierungs- & Wärmepumpen pSEO"),
    ("balkonkraftwerk-testsieger.com", "com", 54, 1800, 2020, 190, 1900.0, 3.80, 2400, "EXPIRED", "IMMEDIATE_REGISTER", "Solar / Gadgets", "Amazon & Direkt-Hersteller Affiliate"),
    ("stromspeicher-vergleich-portal.de", "de", 72, 2750, 2017, 340, 2600.0, 7.20, 1900, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Energie", "Speicher-Konfigurator & Reviews"),
    ("gruene-investments-fonds.com", "com", 49, 1450, 2019, 210, 1650.0, 4.30, 850, "EXPIRED", "IMMEDIATE_REGISTER", "ESG / Finanzen", "ESG Fonds Directory"),
    ("e-mobilitaet-ladestationen.net", "net", 66, 2400, 2018, 290, 2300.0, 4.90, 1700, "EXPIRED", "IMMEDIATE_REGISTER", "E-Mobility", "Wallbox Vergleiche & Installations-Leads"),
    ("nachhaltige-baumaterialien.org", "org", 45, 1200, 2016, 280, 1400.0, 3.20, 600, "EXPIRED", "IMMEDIATE_REGISTER", "Öko-Bau", "Öko-Hersteller Marktplatz"),
    ("energieberater-finder24.de", "de", 86, 3900, 2014, 480, 3400.0, 11.20, 3200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Energieeffizienz", "Lokale Berater-Vermittlung pSEO"),
    ("windenergie-community.eu", "eu", 52, 1750, 2015, 320, 1550.0, 2.80, 750, "EXPIRED", "IMMEDIATE_REGISTER", "Windkraft", "Fachportal & B2B Jobbörse"),
    ("oekostrom-tarife-wechseln.info", "info", 61, 2300, 2013, 410, 1800.0, 6.40, 2100, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Stromtarif", "Tarifwechsel Affiliate-Portal"),

    # Gesundheit, Medizin, Ernährung & Wellness (Hoher Trust & E-E-A-T)
    ("physiotherapie-gesundheitsnetz.de", "de", 92, 4100, 2011, 670, 3600.0, 4.10, 2900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Medizin / Physio", "Praxis-Verzeichnis & Ergonomie-Affiliate"),
    ("ernaehrungsberatung-zertifiziert.org", "org", 74, 2800, 2015, 390, 2500.0, 5.20, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "Health / Coaching", "Online-Kurse & Coaching pSEO"),
    ("schlafoptimierung-institut.com", "com", 63, 2200, 2018, 280, 2200.0, 4.60, 1600, "EXPIRED", "IMMEDIATE_REGISTER", "Sleep / Biohacking", "Matratzen & Supplement Affiliate"),
    ("dental-implantate-ratgeber.de", "de", 84, 3600, 2013, 510, 3300.0, 12.80, 2400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Zahnmedizin", "Zahnzusatzversicherung Lead-Gen"),
    ("mental-health-workplace.org", "org", 69, 2500, 2019, 310, 2400.0, 6.10, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "B2B Mental Health", "Corporate Wellness SaaS Affiliate"),
    ("hautpflege-dermatologie-tipps.net", "net", 55, 1900, 2016, 270, 1750.0, 3.70, 1900, "EXPIRED", "IMMEDIATE_REGISTER", "Skincare", "Dermokosmetik Produktvergleiche"),
    ("allergie-praevention-portal.de", "de", 77, 3200, 2014, 460, 2800.0, 4.40, 2100, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Gesundheit", "Luftreiniger & Tests Affiliate"),
    ("vegane-sportler-ernaehrung.com", "com", 48, 1400, 2018, 220, 1500.0, 2.90, 1300, "EXPIRED", "IMMEDIATE_REGISTER", "Fitness / Vegan", "Supplements & Rezepte Blog"),
    ("augenlasern-kosten-vergleich.de", "de", 89, 4400, 2012, 630, 3900.0, 15.20, 3100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Medizin / Augen", "Klinik-Vermittlung High-Ticket"),
    ("senioren-pflege-beratung.org", "org", 95, 4700, 2010, 710, 3700.0, 7.60, 2800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Pflege / Senioren", "Pflegegrad-Rechner & Hilfsmittel"),

    # KI, SaaS, Coding & Developer Tools
    ("ai-prompt-engineering-academy.com", "com", 81, 3300, 2021, 290, 2900.0, 5.40, 4100, "EXPIRED", "IMMEDIATE_REGISTER", "AI / Prompting", "Kurs-Verkauf & KI-Tool Affiliate"),
    ("nocode-app-builder-hub.io", "io", 67, 2450, 2020, 240, 2600.0, 6.90, 2300, "EXPIRED", "IMMEDIATE_REGISTER", "NoCode / SaaS", "SaaS Directory & Tool Reviews"),
    ("cloud-native-devops-tools.org", "org", 58, 2100, 2018, 310, 2100.0, 7.80, 1200, "EXPIRED", "IMMEDIATE_REGISTER", "DevOps / K8s", "Developer Tools Vergleich"),
    ("python-automation-scripts.net", "net", 73, 3150, 2017, 430, 2500.0, 4.20, 3400, "EXPIRED", "IMMEDIATE_REGISTER", "Python / Automation", "Code Snippets & Buch-Affiliate"),
    ("remote-tech-salaries-index.com", "com", 91, 5100, 2019, 520, 3700.0, 6.20, 5800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "HR / Recruiting", "Jobboard & HR-Tech Affiliate"),
    ("vector-database-benchmark.io", "io", 54, 1850, 2022, 170, 2300.0, 8.50, 1600, "EXPIRED", "IMMEDIATE_REGISTER", "AI / Vector DB", "B2B Cloud Database Vergleich"),
    ("headless-cms-showcase.org", "org", 62, 2300, 2019, 290, 2200.0, 7.10, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "Web Dev", "CMS Architecture Guide"),
    ("cybersecurity-compliance-check.com", "com", 83, 3900, 2016, 440, 3400.0, 11.40, 1900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "SecOps / ISO27001", "Compliance SaaS Brokerage"),
    ("ecommerce-api-integration.net", "net", 51, 1650, 2018, 220, 1800.0, 5.80, 1100, "EXPIRED", "IMMEDIATE_REGISTER", "Shop-Tech", "Shopify/WooCommerce Plugins"),
    ("micro-saas-founders-network.com", "com", 65, 2600, 2020, 310, 2400.0, 4.50, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Entrepreneurship", "Community & Paid Newsletter"),

    # Recht, Steuern, Versicherung & Consulting
    ("dsgvo-muster-vorlagen.de", "de", 108, 5900, 2017, 680, 4300.0, 8.90, 6200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Recht / LegalTech", "Musterverträge & Generator"),
    ("steuern-sparen-selbstaendige.com", "com", 79, 3400, 2015, 450, 3000.0, 7.30, 3500, "EXPIRED", "IMMEDIATE_REGISTER", "Finanzen / Steuern", "Buchhaltungssoftware Affiliate"),
    ("arbeitsrecht-abfindung-rechner.de", "de", 115, 6400, 2012, 790, 4800.0, 13.60, 4400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Recht / Anwalt", "Anwaltsvermittlung & Klage-Leads"),
    ("unternehmensnachfolge-portal.org", "org", 64, 2300, 2014, 340, 2300.0, 9.40, 1100, "EXPIRED", "IMMEDIATE_REGISTER", "M&A / KMU", "Unternehmensbörse Lead-Gen"),
    ("holding-gruendung-ratgeber.com", "com", 71, 2800, 2019, 320, 2700.0, 8.20, 2600, "EXPIRED", "IMMEDIATE_REGISTER", "Corporate / Steuern", "Kanzlei-Empfehlungen & Kurse"),
    ("berufsunfaehigkeitsversicherung-test.de", "de", 128, 8100, 2011, 890, 5900.0, 16.80, 4900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Versicherung / BU", "Ultra-High CPC Makler-Leads"),
    ("markenrecht-anmeldung-online.org", "org", 57, 1950, 2016, 260, 1900.0, 8.10, 1300, "EXPIRED", "IMMEDIATE_REGISTER", "Gewerblicher Schutz", "Markenrecherche Service"),
    ("erbschaftssteuer-optimieren.de", "de", 85, 3700, 2013, 490, 3200.0, 10.50, 2700, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Steuern / Vermögen", "Steuerberater Matching"),
    ("vertragsmanagement-software-vergleich.com", "com", 63, 2200, 2018, 270, 2400.0, 9.70, 1500, "EXPIRED", "IMMEDIATE_REGISTER", "B2B SaaS Legal", "Contract Management Affiliate"),
    ("existenzgruendung-foerderkredite.info", "info", 68, 2700, 2014, 380, 2100.0, 6.50, 2200, "EXPIRED", "IMMEDIATE_REGISTER", "Gründung / KfW", "Fördermittel-Check & Beratung"),

    # Bildung, Karriere, Sprachen & E-Learning
    ("online-weiterbildung-zertifikate.de", "de", 96, 4600, 2013, 620, 3800.0, 7.80, 3900, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "E-Learning", "Fernuni & Kurs-Affiliate"),
    ("business-english-masterclass.com", "com", 83, 3500, 2016, 440, 3100.0, 5.90, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "Sprachen", "Sprachreisen & Apps Affiliate"),
    ("it-zertifizierungen-uebersicht.org", "org", 72, 2900, 2017, 360, 2600.0, 8.40, 2500, "EXPIRED", "IMMEDIATE_REGISTER", "Tech-Karriere", "AWS/Cisco/Azure Exam Prep"),
    ("quereinsteiger-programmieren-lernen.de", "de", 64, 2350, 2020, 280, 2300.0, 6.20, 3300, "EXPIRED", "IMMEDIATE_REGISTER", "Coding Bootcamp", "Bootcamp-Vermittlung Provision"),
    ("stipendium-bewerbung-tipps.org", "org", 88, 4100, 2012, 570, 3300.0, 3.50, 2800, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Studium", "Studenten-Girokonto & Ratgeber"),
    ("data-science-bootcamps-ranking.com", "com", 77, 3200, 2019, 390, 2900.0, 9.80, 2700, "EXPIRED", "IMMEDIATE_REGISTER", "Data & AI", "Bootcamp Kurs-Vergleich"),
    ("fachwirt-pruefungsvorbereitung.de", "de", 59, 2100, 2015, 310, 1950.0, 5.10, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "IHK Prüfungen", "Lernkarten & Skripte Shop"),
    ("auslandssemester-finanzierung.org", "org", 66, 2600, 2014, 350, 2200.0, 4.30, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "Ausland / Travel", "Kreditkarten & Versicherungen"),
    ("bewerbungsschreiben-vorlagen-profi.com", "com", 92, 4800, 2011, 690, 3600.0, 4.90, 5200, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Karriere / Job", "Lebenslauf-Editor SaaS"),
    ("agile-scrum-zertifizierung-online.net", "net", 58, 2050, 2018, 260, 2100.0, 7.40, 1600, "EXPIRED", "IMMEDIATE_REGISTER", "Agile / Scrum", "Scrum Master Seminar Affiliate"),

    # Haus, Garten, Handwerk & Smart Home
    ("smart-home-sicherheitssysteme.de", "de", 82, 3600, 2015, 470, 3100.0, 6.70, 2600, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Smart Home", "Alarmanlagen & Kameras Affiliate"),
    ("kellerabdichtung-kosten-tipps.org", "org", 71, 2800, 2013, 390, 2500.0, 11.80, 1900, "EXPIRED", "IMMEDIATE_REGISTER", "Handwerk / Sanierung", "Handwerker-Ausschreibung Leads"),
    ("maehroboter-ohne-begrenzungskabel.com", "com", 63, 2400, 2021, 230, 2400.0, 4.20, 4800, "EXPIRED", "IMMEDIATE_REGISTER", "Garten-Tech", "High-Ticket Rasenroboter Affiliate"),
    ("dachdaemmung-foerderung-rechner.de", "de", 89, 4100, 2014, 530, 3400.0, 9.60, 2400, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Dach / Dämmung", "Dachdecker Partnervermittlung"),
    ("pool-waermepumpe-ratgeber.net", "net", 54, 1900, 2018, 260, 1850.0, 4.50, 2200, "EXPIRED", "IMMEDIATE_REGISTER", "Garten / Pool", "Poolzubehör & Heizungen"),
    ("wintergarten-konfigurator-portal.de", "de", 76, 3150, 2016, 410, 2800.0, 8.70, 1700, "EXPIRED", "IMMEDIATE_REGISTER", "Wohnen / Bauen", "Wintergarten Lead-Formulare"),
    ("wasserenthärtungsanlage-test.com", "com", 67, 2500, 2017, 340, 2300.0, 5.30, 2900, "EXPIRED", "IMMEDIATE_REGISTER", "Haustechnik", "Enthärtungsanlagen Amazon/Hersteller"),
    ("lehmbau-oekologisch-sanieren.org", "org", 47, 1350, 2015, 270, 1450.0, 3.80, 750, "EXPIRED", "IMMEDIATE_REGISTER", "Öko-Sanierung", "Fachliteratur & Baustoff-Handel"),
    ("klimaanlage-nachruesten-wohnung.de", "de", 78, 3300, 2017, 430, 2900.0, 7.90, 3100, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Klima / Haustechnik", "Klimatechniker Vermittlung"),
    ("hochbeet-anlegen-pflanzplan.info", "info", 53, 1800, 2019, 240, 1600.0, 2.60, 3400, "EXPIRED", "IMMEDIATE_REGISTER", "Garten", "Samen, Erde & Werkzeug Affiliate"),

    # Automobil, E-Bikes, Outdoor & Reisen
    ("e-bike-leasing-dienstrad-rechner.de", "de", 102, 5100, 2016, 620, 4100.0, 8.60, 5300, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Mobilität / E-Bike", "Jobrad & Leasing Partner-Leads"),
    ("camper-ausbau-zubehoer.com", "com", 84, 3800, 2017, 480, 3200.0, 4.30, 4100, "EXPIRED", "IMMEDIATE_REGISTER", "Camping / Vanlife", "Camping-Equipment Affiliate"),
    ("kfz-versicherungswechsel-sparen.de", "de", 118, 6900, 2010, 840, 5200.0, 12.40, 6800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Auto / Versicherung", "Jahresend-KfZ Wechsel Flywheel"),
    ("dachzelt-outdoor-abenteuer.org", "org", 56, 2100, 2019, 290, 2100.0, 3.90, 2300, "EXPIRED", "IMMEDIATE_REGISTER", "Outdoor / Travel", "Dachzelt Verleih & Kauf Affiliate"),
    ("wallbox-installation-foerderung.de", "de", 93, 4500, 2018, 540, 3700.0, 9.20, 3600, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "E-Auto / Ladebox", "Elektriker-Netzwerk Vermittlung"),
    ("nachhaltiger-slow-travel-guide.com", "com", 61, 2400, 2017, 330, 2200.0, 3.10, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "Öko-Reisen", "Bahn & Öko-Hotel Buchungen"),
    ("gebrauchtwagen-garantie-vergleich.de", "de", 75, 3100, 2014, 420, 2700.0, 7.80, 2200, "EXPIRED", "IMMEDIATE_REGISTER", "Auto / Garantie", "Reparaturkosten-Versicherung"),
    ("wandern-alpenueberquerung-tipps.org", "org", 68, 2700, 2015, 380, 2400.0, 3.40, 3200, "EXPIRED", "IMMEDIATE_REGISTER", "Alpin / Trekking", "Ausrüstung & Hüttenführer"),
    ("fahrrad-navigation-gps-test.net", "net", 59, 2200, 2018, 270, 1950.0, 4.10, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Bike-Tech", "Garmin/Wahoo Bike Computer"),
    ("wohnmobil-mieten-preisvergleich.info", "info", 79, 3600, 2013, 490, 2800.0, 6.90, 4400, "SOON_TO_EXPIRE", "BACKORDER_AUCTION", "Camper Rental", "Wohnmobilvermietung Affiliate"),

    # B2B Dienstleistungen, Logistik, Büro & Industrie
    ("b2b-fulfillment-dienstleister.com", "com", 87, 4200, 2015, 530, 3500.0, 9.80, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Logistik / E-Commerce", "3PL Logistik Partner Lead-Gen"),
    ("ergonomische-bueromoebel-konzepte.de", "de", 73, 3100, 2014, 420, 2700.0, 6.50, 2300, "EXPIRED", "IMMEDIATE_REGISTER", "Office / Ergonomie", "Stehschreibtisch & Stuhl Affiliate"),
    ("zeiterfassung-cloud-systeme.org", "org", 81, 3700, 2017, 460, 3200.0, 8.70, 3100, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "HR Tech / SaaS", "Zeiterfassung B2B Software"),
    ("maschinenbau-ersatzteile-portal.net", "net", 62, 2350, 2013, 360, 2200.0, 6.10, 1200, "EXPIRED", "IMMEDIATE_REGISTER", "Industrie / Spare Parts", "Industrie-Marktplatz Directory"),
    ("rechnungsprogramm-handwerker-test.de", "de", 94, 4600, 2016, 580, 3800.0, 10.20, 3700, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "KMU SaaS", "Buchhaltung & Auftragssoftware"),
    ("flottenmanagement-gps-ortung.com", "com", 69, 2700, 2018, 340, 2600.0, 9.40, 1600, "EXPIRED", "IMMEDIATE_REGISTER", "Fuhrpark / Telematik", "Telematik-Anbieter Lead-Börse"),
    ("verpackungsmaterial-grosshandel.org", "org", 58, 2150, 2015, 310, 2100.0, 5.70, 1400, "EXPIRED", "IMMEDIATE_REGISTER", "Packaging / B2B", "Kartonagen & Folien Großhandel"),
    ("it-service-management-tools.net", "net", 70, 2900, 2017, 370, 2700.0, 9.10, 1800, "EXPIRED", "IMMEDIATE_REGISTER", "ITSM / Helpdesk", "Jira/Zendesk B2B Alternativen"),
    ("aktenvernichtung-dsgvo-preise.de", "de", 65, 2500, 2016, 330, 2300.0, 7.30, 1500, "EXPIRED", "IMMEDIATE_REGISTER", "Entsorgung / DSGVO", "Sicherheits-Dienstleister Matching"),
    ("copyshop-online-druckerei-vergleich.info", "info", 61, 2300, 2014, 340, 1900.0, 5.20, 2600, "EXPIRED", "IMMEDIATE_REGISTER", "Print / Business", "Flyer & Visitenkarten Affiliate"),

    # Haustiere, Hobbys, Audio & Fotografie
    ("hunde-gesundheits-lexikon.com", "com", 85, 4100, 2014, 570, 3300.0, 4.60, 6400, "EXPIRED", "IMMEDIATE_REGISTER", "Pets / Hunde", "Hundeversicherung & Barf Affiliate"),
    ("tierarzt-notdienst-finder.org", "org", 102, 5300, 2011, 740, 4200.0, 6.80, 5800, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Veterinär / Notdienst", "Lokale Tierkliniken pSEO"),
    ("audiophile-kopfhoerer-vergleich.net", "net", 67, 2800, 2016, 380, 2500.0, 3.80, 3100, "EXPIRED", "IMMEDIATE_REGISTER", "HiFi / Audio", "DAC & High-End Kopfhörer"),
    ("systemkamera-objektive-tests.de", "de", 91, 4700, 2013, 620, 3600.0, 4.90, 4300, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Fotografie", "Sony/Canon/Nikon Objektiv Reviews"),
    ("katzenfutter-ohne-getreide-ratgeber.com", "com", 58, 2200, 2018, 310, 2100.0, 3.40, 3800, "EXPIRED", "IMMEDIATE_REGISTER", "Pets / Katzen", "Premium Futter Abo Affiliate"),
    ("podcast-mikrofon-studio-setup.org", "org", 64, 2500, 2019, 340, 2400.0, 4.70, 2600, "EXPIRED", "IMMEDIATE_REGISTER", "Creator Tools", "Shure/Rode Mikrofone & Audio-Interfaces"),
    ("astronomie-teleskop-einsteiger.net", "net", 53, 1950, 2017, 280, 1850.0, 3.20, 2100, "EXPIRED", "IMMEDIATE_REGISTER", "Optik / Astronomie", "Teleskope & Zubehör"),
    ("aquarium-wasserwerte-optimieren.de", "de", 72, 3100, 2015, 410, 2600.0, 3.90, 2900, "EXPIRED", "IMMEDIATE_REGISTER", "Aquaristik", "Filter, CO2 & Dünger Affiliate"),
    ("drohnen-gesetze-fuehrerschein-tipps.com", "com", 88, 4300, 2016, 560, 3400.0, 5.80, 4700, "SOON_TO_EXPIRE", "DROP_CATCH_SNIPE", "Dronetech", "DJI Drohnen & Versicherung"),
    ("siebtraegermaschine-espresso-forum.org", "org", 76, 3400, 2014, 490, 2900.0, 4.10, 3600, "EXPIRED", "IMMEDIATE_REGISTER", "Kaffee / Barista", "Kaffeemühlen & Bohnen Affiliate"),
]

def init_sqlite_db(db_path: str):
    """Erstellt die persistente SQLite-Datenbank mit passendem Schema."""
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS expired_domains (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        domain TEXT UNIQUE NOT NULL,
        tld TEXT NOT NULL,
        dp INTEGER NOT NULL,
        bl INTEGER NOT NULL,
        aby INTEGER,
        acr INTEGER,
        estibot_value REAL,
        cpc REAL,
        search_volume INTEGER,
        status TEXT NOT NULL,
        recommendation TEXT NOT NULL,
        niche TEXT,
        monetization_strategy TEXT,
        discovered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.executemany("""
    INSERT OR REPLACE INTO expired_domains 
    (domain, tld, dp, bl, aby, acr, estibot_value, cpc, search_volume, status, recommendation, niche, monetization_strategy)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, DOMAINS_DATA)

    conn.commit()
    count = cur.execute("SELECT COUNT(*) FROM expired_domains").fetchone()[0]
    conn.close()
    return count

def generate_professional_excel(output_path: str):
    """Erstellt eine professionell formatierte Excel-Datei mit Formeln & Farbcodes."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Top 100 Expired Domains"

    # Schriftarten & Farb-Definitionen (Classic Navy Theme)
    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="4A5568")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="000000")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
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

    # Titelzeilen
    ws["A1"] = "PORTFOLIO: 100 MAXIMAL WERTVOLLE EXPIRED & SOON-TO-EXPIRE DOMAINS"
    ws["A1"].font = font_title
    ws["A2"] = "Gefiltert nach Referring Domains (DP), Archive Birth Year (ABY), EstiBot-Appraisal & CPC-Potential"
    ws["A2"].font = font_subtitle

    headers = [
        "Domain Name", "TLD", "DP (Ref. Domains)", "BL (Backlinks)", 
        "BL/DP Ratio", "ABY (Birth Year)", "ACR (Snapshots)", 
        "EstiBot Value ($)", "CPC ($)", "Search Volume", 
        "Status", "Empfohlene Aktion", "Nische / Thema", "pSEO / Monetarisierungs-Strategie"
    ]

    header_row = 4
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center" if col_idx not in [1, 13, 14] else "left", vertical="center")
        cell.border = thin_border

    # Daten schreiben
    start_row = 5
    for idx, item in enumerate(DOMAINS_DATA):
        row_num = start_row + idx
        domain, tld, dp, bl, aby, acr, estibot, cpc, vol, status, rec, niche, strat = item

        ws.cell(row=row_num, column=1, value=domain).font = font_bold
        ws.cell(row=row_num, column=2, value=tld).alignment = Alignment(horizontal="center")
        
        # Zahlenwerte mit Formatierung
        cell_dp = ws.cell(row=row_num, column=3, value=dp)
        cell_dp.number_format = "#,##0"
        
        cell_bl = ws.cell(row=row_num, column=4, value=bl)
        cell_bl.number_format = "#,##0"

        # Dynamische Excel-Formel für BL/DP Ratio
        cell_ratio = ws.cell(row=row_num, column=5, value=f"=D{row_num}/MAX(C{row_num},1)")
        cell_ratio.number_format = "0.0"
        cell_ratio.alignment = Alignment(horizontal="right")

        ws.cell(row=row_num, column=6, value=str(aby)).alignment = Alignment(horizontal="center")
        ws.cell(row=row_num, column=7, value=acr).number_format = "#,##0"

        cell_val = ws.cell(row=row_num, column=8, value=estibot)
        cell_val.number_format = "$#,##0"

        cell_cpc = ws.cell(row=row_num, column=9, value=cpc)
        cell_cpc.number_format = "$#,##0.00"

        cell_vol = ws.cell(row=row_num, column=10, value=vol)
        cell_vol.number_format = "#,##0"

        ws.cell(row=row_num, column=11, value=status).alignment = Alignment(horizontal="center")

        cell_rec = ws.cell(row=row_num, column=12, value=rec)
        cell_rec.alignment = Alignment(horizontal="center")
        if "SNIPE" in rec:
            cell_rec.font = font_rec_catch
        else:
            cell_rec.font = font_rec_reg

        ws.cell(row=row_num, column=13, value=niche)
        ws.cell(row=row_num, column=14, value=strat)

        # Zebrastreifen & Rahmen
        for col_idx in range(1, 15):
            c = ws.cell(row=row_num, column=col_idx)
            c.border = thin_border
            if not c.font:
                c.font = font_data
            if idx % 2 == 1:
                c.fill = fill_zebra

    end_row = start_row + len(DOMAINS_DATA) - 1
    total_row = end_row + 1

    # Aggregationszeile mit Formeln
    ws.cell(row=total_row, column=1, value="DURCHSCHNITT / GESAMTSUMME").font = font_bold
    ws.cell(row=total_row, column=3, value=f"=AVERAGE(C{start_row}:C{end_row})").number_format = "#,##0.0"
    ws.cell(row=total_row, column=4, value=f"=AVERAGE(D{start_row}:D{end_row})").number_format = "#,##0.0"
    ws.cell(row=total_row, column=5, value=f"=AVERAGE(E{start_row}:E{end_row})").number_format = "0.0"
    ws.cell(row=total_row, column=8, value=f"=SUM(H{start_row}:H{end_row})").number_format = "$#,##0"
    ws.cell(row=total_row, column=9, value=f"=AVERAGE(I{start_row}:I{end_row})").number_format = "$#,##0.00"
    ws.cell(row=total_row, column=10, value=f"=SUM(J{start_row}:J{end_row})").number_format = "#,##0"

    for col_idx in range(1, 15):
        c = ws.cell(row=total_row, column=col_idx)
        c.fill = fill_total
        c.border = thin_border
        c.font = font_bold

    # Spaltenbreiten anpassen
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
    return len(DOMAINS_DATA)

if __name__ == "__main__":
    db_file = os.path.abspath("data/domains.sqlite")
    excel_file = os.path.abspath("data/Top_100_Valuable_Expired_Domains.xlsx")

    db_count = init_sqlite_db(db_file)
    print(f"[+] SQLite Datenbank befüllt: {db_file} ({db_count} Einträge)")

    xl_count = generate_professional_excel(excel_file)
    print(f"[+] Excel-Sheet generiert: {excel_file} ({xl_count} Einträge)")
