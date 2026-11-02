"""
Datenmodelle für den Expired Domain Acquisition Lifecycle.
"""
from enum import Enum
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field

class DomainStatus(str, Enum):
    DISCOVERED = "DISCOVERED"
    FILTERED_SPAM = "FILTERED_SPAM"
    APPRAISED = "APPRAISED"
    QUEUED_CATCH = "QUEUED_CATCH"
    SNIPED = "SNIPED"
    REGISTERED = "REGISTERED"
    DEPLOYED = "DEPLOYED"
    IGNORED = "IGNORED"

class ActionRecommendation(str, Enum):
    IMMEDIATE_REGISTER = "IMMEDIATE_REGISTER"
    DROP_CATCH_SNIPE = "DROP_CATCH_SNIPE"
    BACKORDER_AUCTION = "BACKORDER_AUCTION"
    IGNORE = "IGNORE"

class DomainMetrics(BaseModel):
    domain: str
    tld: str
    dp: int = Field(0, description="Referring Domains (Domains pointing to target)")
    bl: int = Field(0, description="Total Backlinks")
    aby: Optional[int] = Field(None, description="Archive Birth Year (First Wayback capture)")
    acr: Optional[int] = Field(None, description="Archive Crawl Results count")
    estibot_value: Optional[float] = Field(None, description="Estimated USD appraisal from EstiBot")
    cpc: Optional[float] = Field(0.0, description="Average CPC for domain keyword")
    search_volume: Optional[int] = Field(0, description="Monthly search volume")
    spam_score: float = Field(0.0, description="Spam risk metric (0.0 to 1.0)")
    is_nxdomain: bool = Field(False, description="DNS resolution returns NXDOMAIN")
    drop_date: Optional[str] = Field(None, description="Predicted or scheduled drop timestamp")

class DomainEvaluation(BaseModel):
    metrics: DomainMetrics
    recommendation: ActionRecommendation
    confidence: float
    reasons: List[str]
    suggested_niche: Optional[str] = None
    target_pseo_template: Optional[str] = None
