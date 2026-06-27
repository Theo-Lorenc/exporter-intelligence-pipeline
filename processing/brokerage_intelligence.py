from processing.text_utils import clean_text


def assign_decision_category(row):
    quality = int(row.get("supplier_quality_score", 0))
    outreach = clean_text(row.get("outreach_ready", ""))
    website = clean_text(row.get("website", ""))
    emails = clean_text(row.get("emails", ""))

    # ✅ Strong signal → ready AND real contact channel
    if quality >= 8 and outreach == "Yes" and website:
        return "Contact Immediately"

    # ✅ Good but not perfect
    if quality >= 6 and outreach == "Yes":
        return "High Potential"

    # ✅ Needs more research
    if quality >= 4:
        return "Investigate Further"

    # ✅ Low signal
    if quality >= 2:
        return "Low Priority"

    # ✅ not useful
    return "Reject"