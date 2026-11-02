"""
Phase 2 & 3: Filtering, Spam Detection & Valuation (Appraisal Engine)
Bewertet DP vs. BL, Archive History, EstiBot Metriken & Traffic-Potential.
"""
from typing import Tuple, List
from .models import DomainMetrics, DomainEvaluation, ActionRecommendation

class FilterAndValuationEngine:
    def __init__(self, min_dp: int = 10, max_bl_dp_ratio: float = 180.0, min_age_years: int = 3):
        self.min_dp = min_dp
        self.max_bl_dp_ratio = max_bl_dp_ratio
        self.min_age_years = min_age_years

    def calculate_spam_score(self, metrics: DomainMetrics) -> Tuple[float, List[str]]:
        """
        Berechnet den Spam-Score (0.0 = clean, 1.0 = highly toxic).
        Kriterien: Unnatürliche Backlink-Verhältnisse (PBN/Spam-Blast),
        extrem junges Archive-Alter bei riesiger Link-Zahl oder fehlende History.
        """
        score = 0.0
        reasons = []

        if metrics.dp == 0 and metrics.bl > 50:
            score += 0.7
            reasons.append("Extreme Backlinks ohne Referring Domains (Ghost Links).")

        ratio = metrics.bl / max(metrics.dp, 1)
        if ratio > self.max_bl_dp_ratio:
            score += 0.5
            reasons.append(f"Verdächtiges BL/DP-Verhältnis ({ratio:.1f}x > {self.max_bl_dp_ratio}x).")

        if metrics.aby is not None:
            import datetime
            current_year = datetime.datetime.now().year
            domain_age = current_year - metrics.aby
            if domain_age < self.min_age_years and metrics.bl > 5000:
                score += 0.4
                reasons.append(f"Domain jünger als {self.min_age_years} Jahre bei massivem Backlink-Volumen.")
        elif metrics.bl > 200:
            score += 0.25
            reasons.append("Kein Archive Birth Year (ABY) bekannt trotz Backlinks.")

        # Cap score at 1.0
        return min(round(score, 2), 1.0), reasons

    def evaluate(self, metrics: DomainMetrics) -> DomainEvaluation:
        """
        Führt das finale Multi-Signal-Scoring durch und gibt eine präzise Handlungsempfehlung ab.
        """
        spam_score, reasons = self.calculate_spam_score(metrics)
        metrics.spam_score = spam_score
        
        # Basierend auf EstiBot / Link-Kapital
        estibot_val = metrics.estibot_value or 0.0
        
        # Entscheidungsmatrix
        if spam_score >= 0.6:
            rec = ActionRecommendation.IGNORE
            reasons.append("Spam-Score übersteigt Risikotoleranz.")
            confidence = 0.95
        elif metrics.dp < self.min_dp and estibot_val < 300:
            rec = ActionRecommendation.IGNORE
            reasons.append(f"Referring Domains ({metrics.dp}) unter Threshold {self.min_dp}.")
            confidence = 0.85
        elif metrics.is_nxdomain:
            if estibot_val > 1000 or metrics.dp >= 50:
                rec = ActionRecommendation.IMMEDIATE_REGISTER
                reasons.append("Hochwertige Domain ist frei verfügbar (NXDOMAIN). Sofortige Registrierung empfohlen.")
                confidence = 0.98
            else:
                rec = ActionRecommendation.IMMEDIATE_REGISTER
                reasons.append("Solides SEO-Kapital bei freier Verfügbarkeit.")
                confidence = 0.88
        else:
            # Domain ist noch im Drop-Prozess / Pending Delete / Auction
            if estibot_val >= 800 or metrics.dp >= 40:
                rec = ActionRecommendation.DROP_CATCH_SNIPE
                reasons.append("Hoher SEO- & Marktwert im Pending Delete Modus. EPP / DropCatch Snipe scharf schalten.")
                confidence = 0.92
            else:
                rec = ActionRecommendation.BACKORDER_AUCTION
                reasons.append("Moderater Wert. Günstiger Backorder-Bid empfohlen.")
                confidence = 0.80

        return DomainEvaluation(
            metrics=metrics,
            recommendation=rec,
            confidence=confidence,
            reasons=reasons,
            suggested_niche="General Authority / 301 Target" if metrics.dp > 50 else "Micro-Niche Affiliate"
        )
