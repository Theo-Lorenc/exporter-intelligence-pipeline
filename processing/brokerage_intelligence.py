from processing.text_utils import clean_text


def assign_decision_category(row):
    score = row.get("brokerage_fit_score", 0)
    outreach_ready = clean_text(row.get("outreach_ready", "")).lower()

    if score >= 12 and outreach_ready == "yes":
        return "Contact Immediately"
    elif score >= 8:
        return "Investigate Further"
    elif score >= 4:
        return "Low Priority"
    else:
        return "Reject"