"""
Orchestrator CLI & Autonomous Pipeline Runner
"""
import sys
import json
from typing import List
from .models import DomainMetrics, ActionRecommendation
from .recon import ReconEngine
from .valuation import FilterAndValuationEngine
from .sniper import DropCatchSniper
from .pseo import PSEODeploymentEngine

class DomainAcquisitionOrchestrator:
    def __init__(self):
        self.recon = ReconEngine()
        self.valuation = FilterAndValuationEngine()
        self.sniper = DropCatchSniper()
        self.pseo = PSEODeploymentEngine()

    def process_domain_candidate(self, metrics: DomainMetrics, auto_snipe: bool = False, generate_pseo: bool = False):
        # 1. DNS NXDOMAIN Check falls noch nicht gesetzt
        if not metrics.is_nxdomain:
            metrics.is_nxdomain = self.recon.is_nxdomain(metrics.domain)

        # 2. Bewertung & Spam Quality Gates
        eval_result = self.valuation.evaluate(metrics)

        # 3. Execution / Sniping
        snipe_result = None
        if eval_result.recommendation in [ActionRecommendation.IMMEDIATE_REGISTER, ActionRecommendation.DROP_CATCH_SNIPE]:
            if auto_snipe:
                snipe_result = self.sniper.execute_snipe(metrics.domain, dry_run=True)

        # 4. pSEO Assets
        pseo_html = None
        if generate_pseo and eval_result.recommendation != ActionRecommendation.IGNORE:
            niche = eval_result.suggested_niche or "Authority"
            keywords = [f"{metrics.domain.split('.')[0]} test", "vergleich", "preis", "erfahrungen"]
            pseo_html = self.pseo.build_pseo_page(metrics.domain, niche, keywords)

        return {
            "domain": metrics.domain,
            "metrics": metrics.model_dump(),
            "evaluation": eval_result.model_dump(),
            "snipe_status": snipe_result,
            "has_pseo_ready": pseo_html is not None
        }

def run_demo_pipeline():
    orchestrator = DomainAcquisitionOrchestrator()
    sample_leads = [
        DomainMetrics(domain="smart-security-advisor.com", tld="com", dp=68, bl=2400, aby=2015, acr=340, estibot_value=1250.0, is_nxdomain=True),
        DomainMetrics(domain="best-solar-calculator.org", tld="org", dp=42, bl=890, aby=2018, acr=110, estibot_value=720.0, is_nxdomain=False),
        DomainMetrics(domain="spammy-free-downloads.net", tld="net", dp=2, bl=18500, aby=2023, acr=12, estibot_value=50.0, is_nxdomain=True),
    ]

    print("\n" + "="*80)
    print("🚀 AUTONOMOUS EXPIRED DOMAIN ACQUISITION & SEO PIPELINE")
    print("="*80 + "\n")
    print(f"{'Domain':<30} | {'DP/BL':<12} | {'EstiBot':<9} | {'Spam':<6} | {'Empfohlene Aktion':<24}")
    print("-" * 88)

    for lead in sample_leads:
        res = orchestrator.process_domain_candidate(lead, auto_snipe=True, generate_pseo=True)
        m = res["metrics"]
        ev = res["evaluation"]
        dp_bl = f"{m['dp']}/{m['bl']}"
        val = f"${m['estibot_value']:.0f}"
        spam = f"{m['spam_score']:.2f}"
        rec = ev["recommendation"]
        print(f"{lead.domain:<30} | {dp_bl:<12} | {val:<9} | {spam:<6} | {rec:<24}")

    print("\n" + "="*80)
    print("✅ Pipeline-Durchlauf erfolgreich abgeschlossen.")
    print("="*80 + "\n")

if __name__ == "__main__":
    run_demo_pipeline()
