from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class ComplaintRecord:
    """
    single processed complaint after SLA analysis.
    maps 1:1 to output rows in the dashboard.
    """
    complaint_id: int
    ro_code: int
    ro_name: str
    vendor_code: str
    assignment_time: datetime
    due_time: datetime
    close_time: Optional[datetime]
    duration_text: str
    delay_hours: float
    status: str            # Early / Delayed / On Time / Pending
    penalty: int
    sla_hours: float
    is_auto_closed: bool   # flagged if remarks contain "auto close"


@dataclass
class UploadSession:
    """
    tracks each file upload — keeps history so you
    can compare month-over-month later.
    """
    upload_id: int = 0
    filename: str = ""
    upload_time: datetime = field(default_factory=datetime.now)
    total_rows: int = 0
    early: int = 0
    delayed: int = 0
    on_time: int = 0
    pending: int = 0
    total_penalty: float = 0.0
