from processing.text_utils import clean_text


def assign_decision_category(row):
    quality = row.get("supplier_quality_score", 0)
    outreach = row.get("outreach_ready", "")

    if quality >= 7 and outreach == "Yes":
        return "Contact Immediately"
    elif quality >= 5:
        return "Investigate Further"
    elif quality >= 3:
        return "Low Priority"
    else:
        return "Reject"