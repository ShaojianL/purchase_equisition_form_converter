"""Field mappings and output columns for the UBC requisition template."""
from __future__ import annotations



FIELDS = {
    "Requester name": "requester_name",
    "Research group / Department": "department",
    "UBC email": "ubc_email",
    "Date requested": "date_requested",
    "Date required (if applicable)": "date_required",
    "Estimated total (CAD)": "estimated_total_cad",
    "Funding worktag / account": "funding_account",
    "Principal Investigator / Budget Owner": "principal_investigator",
    "Preferred delivery location / room": "delivery_location",
    "PI / Authorized Budget Owner Name": "approver_name",
    "Signature / Email Approval": "signature_or_email_approval",
    "Date": "approval_date",
    "Received by": "received_by",
    "Date received": "date_received",
    "Workday requisition / PO reference": "workday_po_reference",
    "Supplier order confirmation": "supplier_order_confirmation",
    "Expected delivery date": "expected_delivery_date",
    "Actual delivery date": "actual_delivery_date",
    "Received / verified by": "verified_by",
}
ITEM_FIELDS = ["item_number", "item_description", "supplier_product_link", "quantity", "estimated_unit_price"]
COLUMNS = ["source_file", *FIELDS.values(), "special_requirements", "additional_instructions",
           "approval", "order_status", "notes_discrepancies", *ITEM_FIELDS]
