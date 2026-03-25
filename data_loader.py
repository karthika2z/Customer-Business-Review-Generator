"""
data_loader.py
Reads all CSVs for a customer and exposes structured properties.
Also provides extract_excel_to_dir() to convert an Excel workbook
(one sheet per dataset) into a directory of CSVs.
"""
import pandas as pd
from pathlib import Path


# Excel truncates sheet names to 31 characters; map those back to the full CSV filenames.
_SHEET_TO_CSV = {
    "Consumption Breakdown by Parame": "Consumption Breakdown by Parameter.csv",
    "List of Tickets Created (For In": "List of Tickets Created (For Internal use).csv",
    "Tickets Created in Past 12 Mont": "Tickets Created in Past 12 Months.csv",
}


def extract_excel_to_dir(xlsx_path, output_dir: Path) -> None:
    """
    Read every sheet of an Excel workbook and write each as a CSV file
    into output_dir.  Sheet names are mapped to CSV filenames via
    _SHEET_TO_CSV (for the three names Excel truncates at 31 chars);
    all others become '<sheet name>.csv'.
    """
    output_dir = Path(output_dir)
    xl = pd.ExcelFile(xlsx_path)
    for sheet in xl.sheet_names:
        csv_name = _SHEET_TO_CSV.get(sheet, f"{sheet}.csv")
        df = xl.parse(sheet)
        df.to_csv(output_dir / csv_name, index=False)


class CustomerData:
    LEVEL_LABELS = {
        1: "Critical",
        2: "Needs Improvement",
        3: "Healthy",
        4: "Best-In-Class",
    }

    def __init__(self, data_dir: Path):
        self.dir = Path(data_dir)
        self._load_all()
        self._extract_key_values()
        self.scores = self._compute_scores()

    # ------------------------------------------------------------------ #
    # CSV helpers
    # ------------------------------------------------------------------ #
    def _csv(self, filename):
        path = self.dir / filename
        if path.exists():
            try:
                return pd.read_csv(path)
            except Exception:
                return pd.DataFrame()
        return pd.DataFrame()

    def _get(self, df, col, default="", row=0):
        if df.empty or col not in df.columns:
            return default
        try:
            val = df.iloc[row][col]
            return default if pd.isna(val) else val
        except (IndexError, KeyError):
            return default

    def _int(self, df, col, default=0, row=0):
        try:
            return int(self._get(df, col, default, row))
        except (ValueError, TypeError):
            return default

    def _get_valid(self, df, col, default=0):
        """Return the first non-NaN value in col, scanning all rows."""
        if df.empty or col not in df.columns:
            return default
        vals = df[col].dropna()
        return float(vals.iloc[0]) if len(vals) else default

    def _uc_flag(self, df, flag_col, usage_col=None):
        """Return 'Y'/'N' for a use case field.
        Falls back to 'Y' if the flag is missing but usage count > 0
        (handles Excel exports where Y/N flags are omitted)."""
        val = str(self._get(df, flag_col, "")).strip().upper()
        if val == "Y":
            return "Y"
        if usage_col and self._int(df, usage_col, 0) > 0:
            return "Y"
        return "N"

    # ------------------------------------------------------------------ #
    # Load all CSVs
    # ------------------------------------------------------------------ #
    def _load_all(self):
        self.df_account         = self._csv("Account Scorecard.csv")
        self.df_features        = self._csv("Features and Usecases.csv")
        self.df_consumption     = self._csv("Consumption Breakdown by Parameter.csv")
        self.df_monthly_spend   = self._csv("Monthly Spend vs Usage.csv")
        self.df_billing_util    = self._csv("Billing Utilization.csv")
        self.df_contracted_nodes= self._csv("Contracted vs Consumed Nodes.csv")
        self.df_detail          = self._csv("Detail.csv")
        self.df_op_readiness    = self._csv("Operational Readiness Score.csv")
        self.df_copilot         = self._csv("Copilot Enabled Controllers.csv")
        self.df_supported_ctrl  = self._csv("Supported Controller Versions.csv")
        self.df_unsupported_ctrl= self._csv("Unsupported Controller Versions.csv")
        self.df_business_value  = self._csv("Business Value.csv")
        self.df_service_status  = self._csv("Service Status Report.csv")
        self.df_ps_util         = self._csv("PS Utilization Report.csv")
        self.df_ace_certs       = self._csv("Active ACE Certs.csv")
        self.df_ace_by_year     = self._csv("Ace Certs by Year.csv")
        self.df_last_ace        = self._csv("Last ACE Cert.csv")
        self.df_next_ace        = self._csv("Next ACE Expiration.csv")
        self.df_tickets_12m     = self._csv("Tickets Created in Past 12 Months.csv")
        self.df_tickets_list    = self._csv("List of Tickets Created (For Internal use).csv")
        self.df_ticket_requestor= self._csv("Ticket Count By Requestor.csv")
        self.df_upgrade_score   = self._csv("Product Upgrade Score.csv")
        self.df_upgrade_tickets = self._csv("Product Upgrade Tickets.csv")
        self.df_reliability_score = self._csv("Reliability & Support Score.csv")
        self.df_release_lifecycle = self._csv("Release Adoption Lifecycle.csv")
        self.df_feature_requests  = self._csv("Feature Requests.csv")
        self.df_fr_total        = self._csv("Feature Requests Total.csv")
        self.df_fr_completed    = self._csv("Feature Requests Completed.csv")
        self.df_fr_outstanding  = self._csv("Feature Requests Outstanding.csv")
        self.df_gateway_count   = self._csv("Gateway Count.csv")
        self.df_parameter_breakdown = self._csv("Parameter Breakdown.csv")

    # ------------------------------------------------------------------ #
    # Extract scalar values
    # ------------------------------------------------------------------ #
    def _extract_key_values(self):
        a = self.df_account

        # Account basics
        self.customer_name  = self._get(a, "Name", "Customer")
        self.arr            = float(self._get(a, "ARR (current mrr)", 0))
        self.account_owner  = self._get(a, "Account Owner Name C", "")
        self.ciso_name      = self._get(a, "Ciso Name", "")
        self.next_renewal   = str(self._get(a, "Next Renewal Date", ""))
        self.billing_type   = self._get(a, "Conversions C", "Term")
        self.account_health = self._get(a, "account_health", "Green")
        self.cnsf_cohort    = self._get(a, "Cnsf Cohort C", "")
        self.transit_gw     = self._int(a, "Transit Gw Sum")
        self.spoke_gw       = self._int(a, "Spoke Gw Sum")

        # Features enabled
        self.copilot_enabled_label = self._get(a, "Copilot Enabled Label", "N")
        self.dcf_enabled    = self._get(a, "DCF Enabled", "N")
        self.firenet_enabled= self._get(a, "Firenet Enabled", "N")
        self.cloudn_edge    = self._get(a, "CloudN/Edge Enabled", "N")
        self.multi_region   = self._get(a, "Multi-Region Transit", "N")

        # Use case usage counts
        self.uc_lateral_usage   = self._int(a, "Prevent Lateral Movement Usage")
        self.uc_3rd_party_usage = self._int(a, "Secure 3rd Party Access Usage")
        self.uc_zt_seg_usage    = self._int(a, "Zero Trust NW Segmentation Usage")
        self.uc_e2e_enc_usage   = self._int(a, "Enforce End-to-Encryption Usage")
        self.uc_exfil_usage     = self._int(a, "Block Data Exfiltration Usage")
        self.uc_dev_vel_usage   = self._int(a, "Accelerate Secure Developer Velocity Usage")

        # Use cases (Y / N) — _uc_flag infers Y from usage > 0 when flag is missing (e.g. Excel)
        self.uc_unified_nw   = self._uc_flag(a, "Unified Cloud NW Fabric")
        self.uc_lateral      = self._uc_flag(a, "Prevent Lateral Movement",              "Prevent Lateral Movement Usage")
        self.uc_zt_seg       = self._uc_flag(a, "Zero Trust NW Segmentation",            "Zero Trust NW Segmentation Usage")
        self.uc_3rd_party    = self._uc_flag(a, "Secure 3rd Party Access",               "Secure 3rd Party Access Usage")
        self.uc_e2e_enc      = self._uc_flag(a, "Enforce End-to-End Encryption",         "Enforce End-to-Encryption Usage")
        self.uc_block_exfil  = self._uc_flag(a, "Block Data Exfiltration",               "Block Data Exfiltration Usage")
        self.uc_dev_velocity = self._uc_flag(a, "Accelerate Secure Developer Velocity",  "Accelerate Secure Developer Velocity Usage")

        # Operational readiness
        op = self.df_op_readiness
        self.supported_controllers   = self._int(op, "Supported Controllers", 0)
        self.unsupported_controllers = self._int(op, "Unsupported Controllers", 0)
        cop = self.df_copilot
        copilot_from_cop = self._int(cop, "Copilot Deployed Sum", -1)
        self.copilot_deployed = (
            copilot_from_cop if copilot_from_cop >= 0
            else self._int(op, "Copilot Deployed Sum", 0)
        )
        self.op_readiness_score      = self._get(op, "Operational Readiness Score", "")

        # Controller detail
        d = self.df_detail
        self.controller_version  = self._get(d, "Controller Version", "")
        self.controller_release  = self._get(d, "Release No", "")
        self.controller_ip       = self._get(d, "Controller Ip", "")
        self.controller_copilot  = self._get(d, "Copilot Enabled", "N")
        self.controller_status   = self._get(d, "Version Support Status", "Supported")

        # Business value
        bv = self.df_business_value
        self.bv_status         = self._get(bv, "Business Value", "Not Initiated")
        self.cbr_status        = self._get(bv, "CBR Status", "")
        self.cnsf_pitch_status = self._get(bv, "CNSF Pitch Status", "")
        self.account_plan      = self._get(bv, "Account Plan", "")

        # Tickets
        t = self.df_tickets_12m
        self.total_tickets   = self._int(t, "Total Tickets")
        self.p1_tickets      = self._int(t, "Aviatrix Software Defects (P1's)")
        self.defect_tickets  = self._int(t, "Aviatrix Software Defects (exl. P1's)")
        self.other_tickets   = self._int(t, "Others/Non-Aviatrix Issues")

        # Reliability & upgrade
        self.reliability_level      = self._get(self.df_reliability_score, "Support Calls Level", "")
        self.upgrade_tickets_count  = self._int(self.df_upgrade_score, "Product Upgrade Tickets")
        self.gateway_sum            = self._int(self.df_upgrade_score, "Gateways Sum", 0)
        self.upgrade_level          = self._get(self.df_upgrade_score, "Product Upgrade Score", "")

        # ACE certifications
        self.active_ace_certs = self._int(self.df_ace_certs, "Total")
        last = self.df_last_ace
        if not last.empty and "Date Certified Max" in last.columns:
            dates = last["Date Certified Max"].dropna()
            self.last_ace_date = str(dates.max()) if len(dates) else ""
        else:
            self.last_ace_date = ""
        self.next_ace_expiry = str(self._get(self.df_next_ace, "Min Expiration Date per User", ""))

        # PS utilization
        self.ps_hours = 0
        ps = self.df_ps_util
        if not ps.empty and "Tracked Hours" in ps.columns:
            try:
                self.ps_hours = int(ps["Tracked Hours"].dropna().sum())
            except Exception:
                self.ps_hours = 0

        # TA / service status
        self.ta_engaged = False
        ss = self.df_service_status
        if not ss.empty and "Project Status" in ss.columns:
            self.ta_engaged = ss["Project Status"].str.lower().str.contains(
                "in progress|active", na=False
            ).any()

        # Feature requests
        # Column may be "Total" or "All Issues Count" depending on export
        _fr_col = lambda df: "Total" if "Total" in df.columns else "All Issues Count"
        self.fr_total_count       = self._int(self.df_fr_total,       _fr_col(self.df_fr_total))
        self.fr_completed_count   = self._int(self.df_fr_completed,   _fr_col(self.df_fr_completed))
        self.fr_outstanding_count = self._int(self.df_fr_outstanding, _fr_col(self.df_fr_outstanding))
        has_fr_rows = (
            not self.df_feature_requests.empty
            and len(self.df_feature_requests.dropna(how="all")) > 0
        )
        self.has_feature_requests = has_fr_rows or self.fr_total_count > 0

        # Billing utilization — scan all rows for first non-NaN (Excel has NaN in row 0)
        bu = self.df_billing_util
        self.consumption_pct = self._get_valid(bu, "Consumption Utilization %", 0)
        self.consumption_mrr = self._get_valid(bu, "Consumption MRR", 0)
        self.billing_mrr     = self._get_valid(bu, "Mrr", 0)

    # ------------------------------------------------------------------ #
    # Score computation (1–4 per pillar)
    # ------------------------------------------------------------------ #
    def _compute_scores(self):
        s = {}

        # 1. Product Utilization
        if self.unsupported_controllers > 0:
            pu = 1
        elif self.copilot_deployed > 0 and self.bv_status not in ("Not Initiated", ""):
            pu = 3
        else:
            pu = 2
        s["Product Utilization"] = pu

        # 2. Reliability & Support
        if self.p1_tickets > 0:
            rs = 1
        elif self.defect_tickets > 0:
            rs = 2
        elif self.total_tickets > 0:
            rs = 3
        else:
            rs = 4
        s["Reliability & Support"] = rs

        # 3. Use Case Adoption
        zt_net = self.uc_unified_nw == "Y" or self.uc_zt_seg == "Y"
        zt_wkl = (
            self.uc_lateral == "Y"
            or self.uc_block_exfil == "Y"
            or self.uc_dev_velocity == "Y"
        )
        if zt_net and zt_wkl:
            uca = 4
        elif zt_net or zt_wkl:
            uca = 3
        elif self.uc_3rd_party == "Y" or self.uc_e2e_enc == "Y":
            uca = 2
        else:
            uca = 1
        s["Use Case Adoption"] = uca

        # 4. Strategic Engagement
        has_active_certs = self.active_ace_certs > 0
        has_past_certs   = self.last_ace_date not in ("", "nan", "NaT")
        has_ps_hours     = self.ps_hours > 0
        # Level 3 (Healthy): active certs + TA engaged OR active certs + PS hours
        if has_active_certs and (self.ta_engaged or has_ps_hours):
            se = 3
        elif has_active_certs or (has_past_certs and (has_ps_hours or self.ta_engaged)):
            se = 2
        elif has_past_certs or self.ta_engaged:
            se = 2
        else:
            se = 1
        s["Strategic Engagement"] = se

        # ZTMM Alignment is Phase 2 — not scored automatically
        return s

    def score_label(self, score: int) -> str:
        return self.LEVEL_LABELS.get(score, "")

    # ------------------------------------------------------------------ #
    # Data helpers for chart builders
    # ------------------------------------------------------------------ #
    def get_consumption_chart_data(self):
        """Return (pivot_mrr, pivot_usage) DataFrames indexed by month."""
        df = self.df_consumption.copy()
        df["Account Name"] = df["Account Name"].astype(str)
        df = df[df["Account Name"].str.strip().str.lower() != "nan"]
        df = df[df["Account Name"].str.strip() != ""]
        df["Date Month"] = pd.to_datetime(df["Date Month"], errors="coerce")
        df = df.dropna(subset=["Date Month"])
        df = df.sort_values("Date Month")

        # Excel exports "Usage" as NaN and puts data in "Usage Sum" instead.
        # Use whichever column has actual values.
        if "Usage" in df.columns and df["Usage"].notna().any():
            usage_col = "Usage"
        elif "Usage Sum" in df.columns and df["Usage Sum"].notna().any():
            usage_col = "Usage Sum"
        else:
            usage_col = "Usage"   # fallback; will produce zeros

        pivot_mrr = df.pivot_table(
            index="Date Month", columns="Parameters",
            values="Consumption MRR", aggfunc="sum", fill_value=0
        )
        pivot_usage = df.pivot_table(
            index="Date Month", columns="Parameters",
            values=usage_col, aggfunc="mean", fill_value=0
        )
        return pivot_mrr, pivot_usage

    def get_tickets_for_display(self):
        """Return list of ticket dicts for the ticket table."""
        df = self.df_tickets_list
        if df.empty:
            return []
        wanted = ["Zendesk Created At Date", "Ticket ID", "Subject", "Product", "Status", "Resolution"]
        rows = []
        for _, row in df.iterrows():
            rows.append({c: str(row.get(c, "")) for c in wanted if c in df.columns})
        return rows

    def get_use_cases(self):
        """Return ordered list of (name, category, status, usage) tuples."""
        return [
            ("Unified Cloud NW Fabric",    "Zero Trust for Networking", self.uc_unified_nw,   None),
            ("Prevent Lateral Movement",   "Zero Trust for Workloads",  self.uc_lateral,       self.uc_lateral_usage),
            ("Zero Trust Segmentation",    "Zero Trust for Networking", self.uc_zt_seg,        self.uc_zt_seg_usage),
            ("Block Data Exfiltration",    "Zero Trust for Workloads",  self.uc_block_exfil,   None),
            ("Secure 3rd Party Access",    "Zero Trust for Networking", self.uc_3rd_party,     self.uc_3rd_party_usage),
            ("Secure Dev Velocity",        "Zero Trust for Workloads",  self.uc_dev_velocity,  None),
            ("End-to-End Encryption",      "Zero Trust for Networking", self.uc_e2e_enc,       None),
        ]
