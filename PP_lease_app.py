import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import streamlit as st
import calendar
import io
import json
import pdfplumber
import docx2txt
from google import genai

# ========== STREAMLIT APP ==========
st.set_page_config(page_title="Lease Schedule Generator", layout="centered")
st.title("📑 Lease Schedule Generator (Stub + Escalation + Monthly Discounting)")

MIN_DATE = datetime(1900, 1, 1).date()
MAX_DATE = datetime(2100, 12, 31).date()

# ===============================
# Helper: Extract lease fields using Gemini
# ===============================
def extract_lease_fields(file_bytes, file_type):

    extraction_prompt = """
You are a lease agreement analyser. Extract the following fields from the lease agreement provided.
Return ONLY a valid JSON object with exactly these keys and no other text, preamble or explanation:

{
    "lease_start_date": "YYYY-MM-DD or null",
    "lease_end_date": "YYYY-MM-DD or null",
    "lock_in_period_months": "integer or null",
    "is_cancellable": "true or false or null",
    "payment_timing": "Beginning or End or null",
    "rent_mode": "Single or Period-wise or null",
    "installment_amount": "float or null",
    "rent_inclusive_of_gst": "true or false or null",
    "escalation_type": "None or Percentage or Fixed Amount or null",
    "escalation_rate": "float or null",
    "escalation_amount": "float or null",
    "escalation_interval_months": "integer or null",
    "rent_periods": [
        {
            "duration_years": "integer",
            "monthly_rent": "float"
        }
    ],
    "gst_rate": "float or null",
    "additional_payments": [
        {
            "label": "string",
            "amount": "float",
            "date": "YYYY-MM-DD",
            "is_refundable": "true or false",
            "inclusive_of_gst": "true or false"
        }
    ],
    "has_purchase_option": "true or false or null",
    "purchase_option_price": "float or null"
}

Rules:
- If a field is not found or unclear, set it to null
- For rent_periods, only populate if rent varies by period, otherwise leave as empty array []
- For additional_payments, include ALL deposits and advances mentioned
- For is_refundable in additional_payments, set true if deposit is refundable, false if non-refundable
- For inclusive_of_gst in additional_payments, set true if agreement explicitly says amount includes GST
- For rent_inclusive_of_gst, set true ONLY if agreement explicitly says rent amount includes GST
- For gst_rate, extract the GST percentage if mentioned anywhere in the agreement
- For is_cancellable, set true if lease has early termination/cancellation clause
- For has_purchase_option, set true if agreement contains option to purchase asset at end of lease
- For purchase_option_price, extract the fixed purchase price mentioned
- Dates must be in YYYY-MM-DD format
- Return ONLY the JSON, no markdown, no explanation, no ```json fences
"""

    if file_type == "application/pdf":
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            text = "\n".join(
                page.extract_text() for page in pdf.pages
                if page.extract_text()
            )
    else:
        text = docx2txt.process(io.BytesIO(file_bytes))

    if not text.strip():
        raise ValueError("Could not extract text from the document. Please check the file.")

    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=extraction_prompt + "\n\nLease Agreement Text:\n" + text
    )

    raw = response.text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ===============================
# GST Helper
# ===============================
def net_of_gst(amount, gst_rate):
    return amount / (1 + gst_rate / 100)

# ===============================
# Section: Upload Lease Agreement
# ===============================
st.subheader("📄 Upload Lease Agreement (Optional)")
uploaded_file = st.file_uploader(
    "Upload lease agreement to auto-fill fields",
    type=["pdf", "docx"]
)

extracted = {}

if uploaded_file is not None:
    if st.button("🔍 Analyse Agreement"):
        with st.spinner("Analysing lease agreement using Gemini AI..."):
            try:
                file_bytes = uploaded_file.read()
                file_type = uploaded_file.type
                extracted = extract_lease_fields(file_bytes, file_type)
                st.session_state["extracted"] = extracted
                st.success("✅ Lease agreement analysed! Fields auto-filled below. Please review and adjust if needed.")

                field_labels = {
                    "lease_start_date": "Lease Start Date",
                    "lease_end_date": "Lease End Date",
                    "payment_timing": "Payment Timing",
                    "is_cancellable": "Cancellable/Non-Cancellable",
                    "rent_mode": "Rent Mode",
                    "installment_amount": "Installment Amount",
                }
                missing = [label for key, label in field_labels.items()
                           if extracted.get(key) is None]
                if missing:
                    st.warning(
                        f"⚠️ Could not extract: **{', '.join(missing)}** — please fill manually."
                    )

            except Exception as e:
                st.error(f"❌ Failed to analyse agreement: {str(e)}")

if "extracted" in st.session_state:
    extracted = st.session_state["extracted"]

def ex(key, default=None):
    val = extracted.get(key)
    return default if val is None else val

# ===============================
# ---- User Inputs ----
# ===============================
st.subheader("📋 Lease Details")

# Lease Start Date
default_start = datetime.today().date()
if ex("lease_start_date"):
    try:
        default_start = datetime.strptime(ex("lease_start_date"), "%Y-%m-%d").date()
    except:
        pass
lease_start = st.date_input(
    "Lease Start Date",
    value=default_start,
    min_value=MIN_DATE,
    max_value=MAX_DATE
)

# Lease End Date
default_end = datetime.today().date()
if ex("lease_end_date"):
    try:
        default_end = datetime.strptime(ex("lease_end_date"), "%Y-%m-%d").date()
    except:
        pass
lease_end_date = st.date_input(
    "Lease End Date",
    value=default_end,
    min_value=MIN_DATE,
    max_value=MAX_DATE
)

# Payment Timing
timing_options = ["Beginning", "End"]
default_timing_idx = timing_options.index(ex("payment_timing", "Beginning")) \
    if ex("payment_timing") in timing_options else 0
payment_timing = st.radio("Payments Timing", timing_options, index=default_timing_idx)

# Interest Rate
interest_rate = st.number_input("Annual Interest Rate (%)", min_value=0.0, step=0.1)

# ---- Cancellable or Non-Cancellable ----
cancel_options = ["Cancellable", "Non-Cancellable"]
default_cancel = "Cancellable" \
    if ex("is_cancellable", True) in [True, "true", "True"] \
    else "Non-Cancellable"
lease_cancellable = st.radio("Is the Lease Cancellable?", cancel_options,
                              index=cancel_options.index(default_cancel))

if lease_cancellable == "Cancellable":
    default_lockin = int(ex("lock_in_period_months", 12))
    lock_in_period = st.number_input("Lock-in Period (months)", min_value=1,
                                      step=1, value=default_lockin)
    lock_in_end = lease_start + relativedelta(months=int(lock_in_period)) - timedelta(days=1)
    st.info(f"🔒 Lock-in End Date: {lock_in_end.strftime('%Y-%m-%d')} | "
            f"📅 Lease End Date (Reference): {lease_end_date.strftime('%Y-%m-%d')}")
else:
    lock_in_period = (
        (lease_end_date.year - lease_start.year) * 12 +
        (lease_end_date.month - lease_start.month) +
        (1 if lease_end_date.day >= lease_start.day else 0)
    )
    st.info(f"📅 Lease End Date: {lease_end_date.strftime('%Y-%m-%d')} | "
            f"🔒 Derived Lease Term: {lock_in_period} months")

# ---- GST Configuration ----
st.subheader("🧾 GST Configuration")

extracted_gst_rate = ex("gst_rate", None)
extracted_rent_gst = ex("rent_inclusive_of_gst", False)

rent_gst_default = extracted_rent_gst in [True, "true", "True"]
rent_inclusive_of_gst = st.checkbox(
    "Rent amount mentioned in agreement is inclusive of GST",
    value=rent_gst_default
)

gst_rate = 0.0
if rent_inclusive_of_gst:
    if extracted_gst_rate is not None:
        st.info(f"ℹ️ GST rate extracted from agreement: **{extracted_gst_rate}%**")
        gst_rate = st.number_input(
            "GST Rate (%) — extracted from agreement, edit if needed",
            min_value=0.0, step=0.5,
            value=float(extracted_gst_rate)
        )
    else:
        st.warning("⚠️ GST rate not found in agreement — please enter manually.")
        gst_rate = st.number_input(
            "GST Rate (%)", min_value=0.0, step=0.5, value=18.0
        )

# ---- Purchase Option ----
st.subheader("🛒 Purchase Option at End of Lease")

has_purchase_option = ex("has_purchase_option", False) in [True, "true", "True"]
purchase_option_exists = st.checkbox(
    "Agreement has a Purchase Option at end of lease",
    value=has_purchase_option
)

exercise_purchase_option = False
purchase_option_price = 0.0
asset_life_months = 0

if purchase_option_exists:
    default_purchase_price = float(ex("purchase_option_price", 0.0))
    purchase_option_price = st.number_input(
        "Purchase Option Price (as per agreement)",
        min_value=0.0, step=100.0,
        value=default_purchase_price
    )
    exercise_purchase_option = st.radio(
        "Do you intend to exercise the Purchase Option?",
        ["Yes — I will buy the asset", "No — I will return the asset"],
        index=0
    ) == "Yes — I will buy the asset"

    if exercise_purchase_option:
        st.success("✅ Purchase price will be included in PV calculation and lease liability.")
        st.info("ℹ️ ROU asset will be amortized over the life of the asset.")
        asset_life_months = st.number_input(
            "Life of Asset (months)", min_value=1, step=1,
            value=int(lock_in_period)
        )

# ---- Rent Configuration ----
st.subheader("🏷️ Rent Configuration")

rent_mode_options = ["Single Installment Amount", "Period-wise Rent"]
extracted_rent_mode = ex("rent_mode", "Single")
default_rent_mode = "Period-wise Rent" \
    if extracted_rent_mode == "Period-wise" and ex("rent_periods") \
    else "Single Installment Amount"
rent_mode = st.radio("Rent Input Mode", rent_mode_options,
                      index=rent_mode_options.index(default_rent_mode))

installment_amount = 0.0
rent_periods = []

if rent_mode == "Single Installment Amount":
    default_amt = float(ex("installment_amount", 0.0))
    installment_amount = st.number_input(
        "Installment Amount (as per agreement, inclusive of GST if applicable)",
        min_value=0.0, step=100.0, value=default_amt
    )

    esc_options = ["None", "Percentage", "Fixed Amount"]
    default_esc = ex("escalation_type", "None")
    default_esc = default_esc if default_esc in esc_options else "None"
    escalation_type = st.selectbox("Escalation Type", esc_options,
                                    index=esc_options.index(default_esc))

    escalation_rate = 0.0
    escalation_amount = 0.0

    if escalation_type == "Percentage":
        escalation_rate = st.number_input(
            "Escalation Rate (%)", min_value=0.0, step=0.5,
            value=float(ex("escalation_rate", 0.0))
        )
    elif escalation_type == "Fixed Amount":
        escalation_amount = st.number_input(
            "Escalation Amount", min_value=0.0, step=100.0,
            value=float(ex("escalation_amount", 0.0))
        )

    escalation_interval = st.number_input(
        "Escalation Interval (months)", min_value=1,
        value=int(ex("escalation_interval_months", 12))
    )

else:
    escalation_type = "None"
    escalation_rate = 0.0
    escalation_amount = 0.0
    escalation_interval = 12

    extracted_periods = ex("rent_periods", [])
    num_periods = st.number_input("Number of Rent Periods", min_value=1,
                                   step=1, value=max(1, len(extracted_periods)))
    total_defined_months = 0

    for p in range(int(num_periods)):
        st.markdown(f"**Rent Period #{p+1}**")
        col1, col2 = st.columns(2)

        default_years = int(extracted_periods[p]["duration_years"]) \
            if p < len(extracted_periods) else 1
        default_rent = float(extracted_periods[p]["monthly_rent"]) \
            if p < len(extracted_periods) else 0.0

        with col1:
            period_years = st.number_input(
                f"Duration (years) #{p+1}", min_value=1, step=1,
                value=default_years, key=f"period_years_{p}"
            )
        with col2:
            period_rent = st.number_input(
                f"Monthly Rent #{p+1} (inclusive of GST if applicable)",
                min_value=0.0, step=100.0,
                value=default_rent, key=f"period_rent_{p}"
            )
        period_months = period_years * 12
        total_defined_months += period_months
        rent_periods.append({"months": period_months, "rent": period_rent})

    if int(lock_in_period) > 0 and total_defined_months > int(lock_in_period):
        st.info(
            f"ℹ️ Total defined period months ({total_defined_months}) exceeds "
            f"Lock-in Period ({int(lock_in_period)} months). "
            f"Periods will be automatically trimmed."
        )
    elif int(lock_in_period) > 0 and total_defined_months < int(lock_in_period):
        st.warning(
            f"⚠️ Total defined period months ({total_defined_months}) is less than "
            f"Lock-in Period ({int(lock_in_period)} months). Please add more periods."
        )

# ---- Additional Payments ----
st.subheader("💰 Additional Payments (Deposits / Advances)")

extracted_additional = ex("additional_payments", [])

non_refundable_extracted = [
    ap for ap in extracted_additional
    if ap.get("is_refundable") not in [True, "true", "True"]
]

if extracted_additional:
    refundable_count = len(extracted_additional) - len(non_refundable_extracted)
    if refundable_count > 0:
        st.info(f"ℹ️ {refundable_count} refundable deposit(s) found — excluded from schedule automatically.")

num_additional = st.number_input(
    "Number of Additional Payments (non-refundable only)",
    min_value=0, step=1,
    value=len(non_refundable_extracted)
)

additional_payments = []
for j in range(int(num_additional)):
    st.markdown(f"**Additional Payment #{j+1}**")
    col1, col2, col3 = st.columns(3)

    ap_ex = non_refundable_extracted[j] if j < len(non_refundable_extracted) else {}

    default_label = ap_ex.get("label", f"Deposit {j+1}")
    default_ap_amt = float(ap_ex.get("amount", 0.0))
    default_ap_date = datetime.today().date()
    if ap_ex.get("date"):
        try:
            default_ap_date = datetime.strptime(ap_ex["date"], "%Y-%m-%d").date()
        except:
            pass

    ap_gst_default = ap_ex.get("inclusive_of_gst", False) in [True, "true", "True"]

    with col1:
        ap_label = st.text_input(f"Label #{j+1}", value=default_label,
                                  key=f"ap_label_{j}")
    with col2:
        ap_amount = st.number_input(
            f"Amount #{j+1} (gross)",
            min_value=0.0, step=100.0,
            value=default_ap_amt, key=f"ap_amount_{j}"
        )
    with col3:
        ap_date = st.date_input(
            f"Payment Date #{j+1}",
            value=default_ap_date,
            min_value=MIN_DATE,
            max_value=MAX_DATE,
            key=f"ap_date_{j}"
        )

    ap_inclusive_gst = st.checkbox(
        f"Amount #{j+1} is inclusive of GST",
        value=ap_gst_default,
        key=f"ap_gst_{j}"
    )

    if ap_inclusive_gst and gst_rate > 0:
        net_amt = net_of_gst(ap_amount, gst_rate)
        st.caption(f"Net amount (excl. GST @ {gst_rate}%): ₹{net_amt:,.2f}")
    else:
        net_amt = ap_amount

    additional_payments.append({
        "label": ap_label,
        "amount": net_amt,
        "date": ap_date
    })

# ===============================
# Generate Schedule
# ===============================
if st.button("Generate Schedule"):

    # ✅ Short-term lease check
    if int(lock_in_period) <= 12:
        st.warning(
            "⚠️ This is a **Short-Term Lease** (lock-in period is 12 months or less). "
            "As per Ind AS 116 / IFRS 16, no lease schedule is required. "
            "Rentals may be recognised as expense on a straight-line basis."
        )
        st.stop()

    # ✅ Interest rate validation
    if interest_rate == 0.0:
        st.warning(
            "⚠️ Please enter a valid **Annual Interest Rate** before generating the schedule. "
            "Interest rate cannot be zero."
        )
        st.stop()

    # ---- Validate period-wise rent ----
    if rent_mode == "Period-wise Rent":
        total_defined_months = sum(p["months"] for p in rent_periods)
        if total_defined_months < int(lock_in_period):
            st.error(
                f"❌ Total defined period months ({total_defined_months}) "
                f"is less than Lock-in Period ({int(lock_in_period)} months). "
                f"Please add more periods."
            )
            st.stop()

    t0 = lease_start
    annual_rate = interest_rate / 100
    monthly_rate = annual_rate / 12

    # ===============================
    # Step 1: Build monthly rent schedule (net of GST)
    # ===============================
    monthly_rents = []

    def apply_gst(amount):
        if rent_inclusive_of_gst and gst_rate > 0:
            return net_of_gst(amount, gst_rate)
        return amount

    if rent_mode == "Single Installment Amount":
        base_rent = apply_gst(installment_amount)
        current_rent = base_rent
        for m in range(1, int(lock_in_period) + 1):
            if escalation_type != "None" and m > 1 and (m - 1) % escalation_interval == 0:
                if escalation_type == "Percentage":
                    current_rent = current_rent * (1 + escalation_rate / 100)
                elif escalation_type == "Fixed Amount":
                    current_rent = current_rent + apply_gst(escalation_amount)
            monthly_rents.append(current_rent)
    else:
        months_remaining = int(lock_in_period)
        for period in rent_periods:
            if months_remaining <= 0:
                break
            months_to_use = min(period["months"], months_remaining)
            net_rent = apply_gst(period["rent"])
            for _ in range(months_to_use):
                monthly_rents.append(net_rent)
            months_remaining -= months_to_use

    # ===============================
    # Step 2: Build Payment Dates with Stub
    # ===============================
    dates = []
    amounts = []
    labels = []

    days_in_month = calendar.monthrange(t0.year, t0.month)[1]
    end_of_month = datetime(t0.year, t0.month, days_in_month).date()

    has_stub = (t0.day != 1 and t0.day != calendar.monthrange(t0.year, t0.month)[1])

    if has_stub:
        stub_days = (end_of_month - t0).days + 1
        prorated = monthly_rents[0] * (stub_days / days_in_month)
        dates.append(end_of_month if payment_timing.lower() == "end" else t0)
        amounts.append(prorated)
        labels.append("Lease Rental")
        start_for_regular = end_of_month + timedelta(days=1)
    else:
        dates.append(t0 if payment_timing.lower() == "beginning" else end_of_month)
        amounts.append(monthly_rents[0])
        labels.append("Lease Rental")
        start_for_regular = t0 + relativedelta(months=1)

    for i in range(1, int(lock_in_period) + (1 if has_stub else 0)):
        rent_idx = min(i, len(monthly_rents) - 1)
        pay_month = start_for_regular + relativedelta(months=i - 1)
        dim = calendar.monthrange(pay_month.year, pay_month.month)[1]

        if payment_timing.lower() == "beginning":
            pay_date = datetime(pay_month.year, pay_month.month, 1).date()
        else:
            pay_date = datetime(pay_month.year, pay_month.month, dim).date()

        dates.append(pay_date)
        amounts.append(monthly_rents[rent_idx])
        labels.append("Lease Rental")

    # Last Payment Stub Adjustment
    lease_end = t0 + relativedelta(months=int(lock_in_period)) - timedelta(days=1)
    if lease_end.day != calendar.monthrange(lease_end.year, lease_end.month)[1]:
        dim = calendar.monthrange(lease_end.year, lease_end.month)[1]
        last_rent = monthly_rents[-1]
        prorated = last_rent * (lease_end.day / dim)
        dates[-1] = lease_end
        amounts[-1] = prorated

    # Add Purchase Option
    if exercise_purchase_option and purchase_option_price > 0:
        dates.append(lease_end)
        amounts.append(purchase_option_price)
        labels.append("Purchase Option")

    # Merge Additional Payments
    for ap in additional_payments:
        dates.append(ap["date"])
        amounts.append(ap["amount"])
        labels.append(ap["label"])

    # Sort all by date
    combined = sorted(zip(dates, amounts, labels), key=lambda x: x[0])
    dates, amounts, labels = zip(*combined)
    dates = list(dates)
    amounts = list(amounts)
    labels = list(labels)

    # ===============================
    # Step 3: PV Calculation
    # ===============================
    days_in_start_month = calendar.monthrange(t0.year, t0.month)[1]
    stub_days = (end_of_month - t0).days + 1 if has_stub else days_in_start_month
    discount_stub_fraction = stub_days / days_in_start_month

    pv_list = []
    lease_rental_indices = [i for i, l in enumerate(labels) if l == "Lease Rental"]
    last_lease_rental_idx = lease_rental_indices[-1]

    for i in range(len(dates)):
        amt = amounts[i]
        months_diff = (dates[i].year - t0.year) * 12 + (dates[i].month - t0.month)

        if dates[i] < t0:
            period = 0

        elif dates[i] == t0:
            period = 0

        elif i == last_lease_rental_idx:
            period = int(lock_in_period)

        elif months_diff == 0:
            period = discount_stub_fraction

        else:
            if payment_timing.lower() == "beginning":
                period = (months_diff - 1) + discount_stub_fraction
            else:
                period = months_diff + (discount_stub_fraction if has_stub else 0)

        pv = amt / ((1 + monthly_rate) ** period)
        pv_list.append(pv)

    lease_liability = sum(pv_list)
    ROU_opening = lease_liability

    # ===============================
    # ROU Amortization Period
    # ===============================
    if exercise_purchase_option and asset_life_months > 0:
        amortization_months = int(asset_life_months)
        st.info(f"ℹ️ ROU asset amortized over **{amortization_months} months** (Life of Asset)")
    else:
        amortization_months = int(lock_in_period)

    amortization_per_month = ROU_opening / amortization_months

    # ===============================
    # Step 4: Build Schedule
    # ===============================
    rows = []
    opening_balance = lease_liability
    rou_balance = ROU_opening

    i = 0
    sl_no = 1

    while i < len(dates):
        # Group all payments on same date
        same_date_indices = [i]
        while i + 1 < len(dates) and dates[i + 1] == dates[i]:
            i += 1
            same_date_indices.append(i)

        current_date = dates[same_date_indices[0]]
        total_payments_today = sum(amounts[k] for k in same_date_indices)
        months_diff = (current_date.year - t0.year) * 12 + (current_date.month - t0.month)

        # ===============================
        # Interest Calculation
        # ===============================
        if current_date < t0:
            # Prior deposit — lease not yet commenced, no interest
            interest = 0.0

        elif has_stub and months_diff == 0 and same_date_indices[0] == 0:
            # First stub period payment
            if payment_timing.lower() == "beginning":
                interest = (opening_balance - total_payments_today) * monthly_rate * discount_stub_fraction
            else:
                interest = opening_balance * monthly_rate * discount_stub_fraction

        elif last_lease_rental_idx in same_date_indices and has_stub:
            # Last partial (stub tail) period
            dim_last = calendar.monthrange(current_date.year, current_date.month)[1]
            last_fraction = current_date.day / dim_last
            if payment_timing.lower() == "beginning":
                interest = (opening_balance - total_payments_today) * monthly_rate * last_fraction
            else:
                interest = opening_balance * monthly_rate * last_fraction

        else:
            # All regular full months including t0 no-stub first payment
            if payment_timing.lower() == "beginning":
                # Payment made on day 1; interest accrues on post-payment balance for the month
                interest = (opening_balance - total_payments_today) * monthly_rate
            else:
                # Payment at end; interest accrues on full opening balance
                interest = opening_balance * monthly_rate

        # ===============================
        # Apply payments row by row
        # ===============================
        interest_applied = False

        for idx in same_date_indices:
            payment = amounts[idx]
            label = labels[idx]

            if not interest_applied:
                row_interest = interest
                if payment_timing.lower() == "beginning":
                    balance_after_interest = (opening_balance - total_payments_today) + interest + payment
                else:
                    balance_after_interest = opening_balance + interest
                interest_applied = True
            else:
                row_interest = 0.0
                balance_after_interest = opening_balance

            closing_balance = balance_after_interest - payment

            # ROU Amortization
            if label not in ["Lease Rental", "Purchase Option"]:
                rou_amort = 0.0
            elif label == "Purchase Option":
                rou_amort = 0.0
            elif idx == 0:
                rou_amort = amortization_per_month * discount_stub_fraction
            elif idx == last_lease_rental_idx and has_stub:
                rou_amort = amortization_per_month * (1 - discount_stub_fraction)
            else:
                rou_amort = amortization_per_month

            rou_balance -= rou_amort

            rows.append([
                sl_no,
                current_date.strftime("%Y-%m-%d"),
                label,
                round(payment, 2),
                round(pv_list[idx], 2),
                round(opening_balance, 2),
                round(payment, 2),
                round(row_interest, 2),
                round(closing_balance, 2),
                round(rou_balance + rou_amort, 2),
                round(rou_amort, 2),
                round(rou_balance, 2)
            ])

            opening_balance = closing_balance
            sl_no += 1

        i += 1

    df = pd.DataFrame(rows, columns=[
        "Sl No.", "Installment Date", "Payment Type", "Amount Paid", "PV of Installment",
        "Opening Lease Liability", "Payment", "Interest", "Closing Lease Liability",
        "ROU Opening", "Amortization", "ROU Closing"
    ])

    if rent_inclusive_of_gst and gst_rate > 0:
        st.info(f"ℹ️ All amounts shown are net of GST @ {gst_rate}%")

    st.success("✅ Lease Schedule Generated Successfully")
    st.dataframe(df, use_container_width=True)

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=buffer,
        file_name="Lease_Schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
