"""
outlook_sender.py
-----------------
Opens a pre-filled Outlook email (draft, NOT sent) containing a mill's NEW
tags formatted in the tag-request column order, ready for a human to review
and send to the agent.

Uses Outlook COM automation, so it requires CLASSIC Outlook installed on the
machine that runs this (new Outlook lacks COM support). The draft opens on
THAT machine — here, roc-wfm-rv.

Called by the API's /workbooks/{code}/send-new endpoint (the "Send" button in
the New-tags view). The draft is displayed, not sent — the user sends at their
discretion.
"""

# The request columns the agent expects, in order. Some (Units, Frequency,
# Min, Max) aren't in the workbook, so they come out blank — that's expected.
REQUEST_COLUMNS = [
    "TableName", "TagName", "PLCIPAddress", "PLCPath", "DataType",
    "Units", "TransactionFrequency", "TagDescription", "Min", "Max", "NewTable?",
]

AGENT_ADDRESS = "WFM-OT-TagsRUsAgent@westfraser.com"


def _escape(text):
    """Minimal HTML escaping for cell values."""
    s = "" if text is None else str(text)
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))


def build_table_html(rows):
    """
    Build the HTML table in the request-column order from workbook New-tag rows.

    Each row is a dict from workbook_manager.read_rows (keys: table, tag_name,
    plc_tag, plc_name, plc_ip, data_type, tag_purpose, ...). We map those onto
    the request columns; fields the workbook doesn't carry are left blank.
    """
    head = "".join(f"<th>{c}</th>" for c in REQUEST_COLUMNS)
    body_rows = []
    for r in rows:
        # PLCIPAddress = PLC Name + PLC IP combined (space-separated), matching
        # the Tag Request Form convention. Handles either field being blank.
        plc_name = (r.get("plc_name") or "").strip()
        plc_ip   = (r.get("plc_ip") or "").strip()
        plc_ip_address = " ".join(p for p in (plc_name, plc_ip) if p)

        cells = [
            r.get("table", ""),          # TableName
            r.get("tag_name", ""),       # TagName
            plc_ip_address,              # PLCIPAddress = PLC Name + PLC IP
            r.get("plc_tag", ""),        # PLCPath  (workbook's PLC tag/path)
            r.get("data_type", ""),      # DataType
            "",                          # Units             (not in workbook)
            "",                          # TransactionFrequency (not in workbook)
            r.get("tag_purpose", ""),    # TagDescription
            "",                          # Min               (not in workbook)
            "",                          # Max               (not in workbook)
            "No",                        # NewTable?         (default No)
        ]
        tds = "".join(f"<td>{_escape(c)}</td>" for c in cells)
        body_rows.append(f"<tr>{tds}</tr>")

    return (
        '<table border="1" style="border-collapse:collapse;">'
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
    )


def open_draft(mill_code, mill_name, rows):
    """
    Open an Outlook draft addressed to the agent, with the New-tags table in the
    body. Returns the number of tags placed in the draft. Raises if Outlook COM
    isn't available (e.g. new Outlook, or not installed).
    """
    import win32com.client  # imported here so the module loads without pywin32

    if not rows:
        raise ValueError(f"No new tags to send for {mill_code}.")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)          # 0 = olMailItem
    mail.To = AGENT_ADDRESS
    mail.Subject = f"{mill_code} Tag Request: {len(rows)} new tags [Workbook]"

    mail.HTMLBody = build_table_html(rows)

    mail.Display()                        # opens the draft; does NOT send
    return len(rows)