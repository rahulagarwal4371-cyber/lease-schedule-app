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
# Helper: Compute lock-in months correctly accounting for stub months
# ===============================
def compute_lock_in_months(lease_start, lease_end_date):
    """
    Count the number of schedule months between lease_start and lease_end_date.

    When lease_start is NOT the 1st of a month (stub month), month 1 is the
    partial period from lease_start to the end of that calendar month.  Month 2
    is the next full calendar month, and so on.  The final schedule month is
    whichever calendar month contains lease_end_date.

    Example: Mar 10 2024 → Mar 9 2026
      Month 1  = Mar 10-31 (stub)
      Month 2  = Apr 2024
      ...
      Month 24 = Feb 2026
      Month 25 = Mar 2026  (contains Mar 9, the lease end)
      → returns 25
    """
    t0 = lease_start
    days_in_month = calendar.monthrange(t0.year, t0.month)[1]
    has_stub = (t0.day != 1 and t0.day != days_in_month)

    if not has_stub:
        # No stub: straightforward month count
        return (
            (lease_end_date.year - lease_start.year) * 12 +
            (lease_end_date.month - lease_start.month) +
            (1 if lease_end_date.day >= lease_start.day else 0)
        )
    else:
        # Stub present: month 1 is partial, so every subsequent month is one
        # calendar month further than the raw date-diff would suggest.
        end_of_stub = datetime(t0.year, t0.month, days_in_month).date()
        first_full = end_of_stub + timedelta(days=1)          # start of month 2
        full_months = (
            (lease_end_date.year - first_full.year) * 12 +
            (lease_end_date.month - first_full.month)
        )
        # 1 (stub month) + full calendar months between month-2-start and
        # lease_end_date's month + 1 (the month that contains lease_end_date)
        return 1 + full_months + 1


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


def compute_true_lease_months(start, end):
    return (
        (end.year - start.year) * 12 +
        (end.month - start.month)
    )

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

# ---- Payment Frequency ----
st.subheader("🗓️ Payment Frequency")

frequency_options = {
    "Monthly (every 1 month)": 1,
    "Bi-Monthly (every 2 months)": 2,
    "Quarterly (every 3 months)": 3,
    "Half-Yearly (every 6 months)": 6,
    "Annually (every 12 months)": 12,
}
selected_frequency_label = st.selectbox(
    "Payment Frequency",
    list(frequency_options.keys()),
    index=0,
    help="How often rent payments are made. The schedule remains monthly; payments are aggregated per frequency."
)
payment_frequency = frequency_options[selected_frequency_label]

# Starting month of payment (only relevant when frequency > 1)
payment_start_month = 1
if payment_frequency > 1:
    st.info(
        f"ℹ️ Payments are made every **{payment_frequency} months**. "
        f"Please select which month (within the cycle) the first payment falls in."
    )
    payment_start_month = st.number_input(
        f"Starting Month of First Payment (1 to {payment_frequency})",
        min_value=1,
        max_value=payment_frequency,
        value=1,
        step=1,
        help=(
            f"E.g. if frequency is quarterly (3 months) and you set starting month = 2, "
            f"payments will fall in months 2, 5, 8, 11, ... of the lease."
        )
    )
    st.caption(
        f"📅 Payments will occur in lease months: "
        f"{', '.join(str(payment_start_month + i * payment_frequency) for i in range(min(5, (int((ex('lock_in_period_months', 36) or 36) - payment_start_month) // payment_frequency) + 1)))}..."
    )

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
    # Use compute_lock_in_months to correctly handle stub-month boundary shifts
    lock_in_period = compute_lock_in_months(lease_start, lease_end_date)
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

if payment_frequency > 1:
    st.info(
        f"ℹ️ **Frequency note:** Enter the **monthly rent** amount below. "
        f"The schedule will aggregate payments every {payment_frequency} months "
        f"(i.e. each payment = {payment_frequency} × monthly rent, adjusted for stubs)."
    )

installment_amount = 0.0
rent_periods = []

if rent_mode == "Single Installment Amount":
    default_amt = float(ex("installment_amount", 0.0))
    installment_amount = st.number_input(
        "Monthly Installment Amount (as per agreement, inclusive of GST if applicable)",
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

    # ---- Validate starting month ----
    if payment_frequency > 1 and payment_start_month > payment_frequency:
        st.error(
            f"❌ Starting month ({payment_start_month}) cannot exceed "
            f"the payment frequency ({payment_frequency} months)."
        )
        st.stop()

    t0 = lease_start
    annual_rate = interest_rate / 100
    monthly_rate = annual_rate / 12

    # -----------------------------------------------------------------------
    # Pre-compute stub / tail-stub metadata
    # -----------------------------------------------------------------------
    days_in_first_month = calendar.monthrange(t0.year, t0.month)[1]
    end_of_month = datetime(t0.year, t0.month, days_in_first_month).date()

    # A stub exists when lease starts after the 1st and before the last day
    has_stub = (t0.day != 1 and t0.day != days_in_first_month)

    lock_in = int(lock_in_period)

    # Actual lease end date:
    #   Non-cancellable → use the user-supplied lease_end_date directly
    #   Cancellable     → derive from lock_in (user entered months manually)
    if lease_cancellable == "Non-Cancellable":
        lease_end = lease_end_date
    else:
        lease_end = t0 + relativedelta(months=lock_in) - timedelta(days=1)

    # ✅ ALWAYS define (outside if-else)
    true_lease_months = compute_true_lease_months(t0, lease_end)

    # Stub fraction: proportion of the first calendar month covered
    stub_days_count = (end_of_month - t0).days + 1 if has_stub else days_in_first_month
    discount_stub_fraction = stub_days_count / days_in_first_month

    # Tail-stub fraction: proportion of the last calendar month covered
    last_dim = calendar.monthrange(lease_end.year, lease_end.month)[1]
    last_month_fraction = lease_end.day / last_dim
    has_tail_stub = (lease_end.day != last_dim)

    # -----------------------------------------------------------------------
    # Step 1: Build monthly rent schedule (net of GST, length = lock_in)
    # -----------------------------------------------------------------------
    monthly_rents = []

    def apply_gst(amount):
        if rent_inclusive_of_gst and gst_rate > 0:
            return net_of_gst(amount, gst_rate)
        return amount

    if rent_mode == "Single Installment Amount":
        base_rent = apply_gst(installment_amount)
        current_rent = base_rent
        for m in range(1, lock_in + 1):
            if escalation_type != "None" and m > 1 and (m - 1) % escalation_interval == 0:
                if escalation_type == "Percentage":
                    current_rent = current_rent * (1 + escalation_rate / 100)
                elif escalation_type == "Fixed Amount":
                    current_rent = current_rent + apply_gst(escalation_amount)
            monthly_rents.append(current_rent)
    else:
        months_remaining = lock_in
        for period in rent_periods:
            if months_remaining <= 0:
                break
            months_to_use = min(period["months"], months_remaining)
            net_rent = apply_gst(period["rent"])
            for _ in range(months_to_use):
                monthly_rents.append(net_rent)
            months_remaining -= months_to_use

    # -----------------------------------------------------------------------
    # Step 2: Determine payment months (based on frequency and starting month)
    # -----------------------------------------------------------------------
    payment_months_set = set()
    m = payment_start_month
    while m <= lock_in:
        payment_months_set.add(m)
        m += payment_frequency

    # -----------------------------------------------------------------------
    # Step 3: Build payment_buckets
    #   Maps payment_month_num → list of lease month numbers whose rent is
    #   collected in that payment.
    #
    #   The last lease month (lock_in) always gets its own dedicated bucket
    #   so it is never silently merged into an earlier one.
    # -----------------------------------------------------------------------
    first_pay_month = payment_start_month
    payment_buckets = {}

    for lm in range(1, lock_in + 1):
        if lm == lock_in:
            assigned_pay_month = lock_in
        else:
            if lm <= first_pay_month:
                assigned_pay_month = first_pay_month
            else:
                steps = ((lm - first_pay_month - 1) // payment_frequency) + 1
                assigned_pay_month = first_pay_month + steps * payment_frequency
            assigned_pay_month = min(assigned_pay_month, lock_in)

        if assigned_pay_month not in payment_buckets:
            payment_buckets[assigned_pay_month] = []
        payment_buckets[assigned_pay_month].append(lm)

    # -----------------------------------------------------------------------
    # Helper: calendar date on which a bucket's payment is made
    #
    #   Rules for the last bucket when has_tail_stub is True:
    #     • Beginning timing → 1st of the month containing lease_end
    #                          (tenant pays at the START of the partial month;
    #                           amount is prorated but date is always the 1st)
    #     • End timing       → actual lease_end date
    #                          (tenant pays when the period closes)
    #
    #   For all other buckets:
    #     • Beginning → 1st of the relevant calendar month
    #     • End       → last day of the relevant calendar month
    # -----------------------------------------------------------------------
    def get_payment_date_for_bucket(pay_month_num, bucket_months):
        # ── Stub month (month 1) ──────────────────────────────────────────
        if pay_month_num == 1 and has_stub:
            return t0 if payment_timing.lower() == "beginning" else end_of_month

        # ── Determine the calendar month this bucket maps to ──────────────
        if has_stub:
            month_offset = pay_month_num - 2   # month 2 → offset 0
            base = end_of_month + timedelta(days=1)
        else:
            month_offset = pay_month_num - 1   # month 1 → offset 0
            base = t0

        pay_month_date = base + relativedelta(months=month_offset)
        dim = calendar.monthrange(pay_month_date.year, pay_month_date.month)[1]

        if payment_timing.lower() == "beginning":
            # ── Beginning timing ──────────────────────────────────────────
            # Always the 1st of the month — even for the tail-stub last month.
            # (The rent amount is prorated via last_month_fraction, but the
            #  payment DATE is the start of that calendar month.)
            return datetime(pay_month_date.year, pay_month_date.month, 1).date()
        else:
            # ── End timing ────────────────────────────────────────────────
            # Tail-stub last month: pay on the actual lease end date.
            # All other months: pay on the last day of the calendar month.
            if pay_month_num == lock_in and has_tail_stub:
                return lease_end
            return datetime(pay_month_date.year, pay_month_date.month, dim).date()

    # -----------------------------------------------------------------------
    # Helper: prorated (or full) rent for a given lease month number
    # -----------------------------------------------------------------------
    def get_month_rent(lm):
        idx = min(lm - 1, len(monthly_rents) - 1)
        rent = monthly_rents[idx]
        if lm == 1 and has_stub:
            return rent * discount_stub_fraction
        if lm == lock_in and has_tail_stub:
            return rent * last_month_fraction
        return rent

    # -----------------------------------------------------------------------
    # Build the flat lists of (date, amount, label) from buckets
    # -----------------------------------------------------------------------
    dates, amounts, labels = [], [], []

    for pay_month_num in sorted(payment_buckets.keys()):
        bucket_months = payment_buckets[pay_month_num]
        total_rent = sum(get_month_rent(lm) for lm in bucket_months)
        pay_date = get_payment_date_for_bucket(pay_month_num, bucket_months)
        dates.append(pay_date)
        amounts.append(total_rent)
        labels.append("Lease Rental")

    # Purchase Option
    if exercise_purchase_option and purchase_option_price > 0:
        dates.append(lease_end)
        amounts.append(purchase_option_price)
        labels.append("Purchase Option")

    # Additional / deposit payments
    for ap in additional_payments:
        dates.append(ap["date"])
        amounts.append(ap["amount"])
        labels.append(ap["label"])

    # Sort everything by date
    combined = sorted(zip(dates, amounts, labels), key=lambda x: x[0])
    dates, amounts, labels = zip(*combined)
    dates, amounts, labels = list(dates), list(amounts), list(labels)

    # -----------------------------------------------------------------------
    # Step 4: PV calculation
    #   Each payment is discounted by its exact month-offset from t0.
    # -----------------------------------------------------------------------
    def get_period_for_pv(pay_date, label, true_lease_months):
        months_diff = (pay_date.year - t0.year) * 12 + (pay_date.month - t0.month)

        if pay_date <= t0:
            return 0  # Same day or prior deposit → no discounting

        if months_diff == 0:
            # Same calendar month as t0 (stub payment within month 1)
            return discount_stub_fraction

        if payment_timing.lower() == "beginning":
            period = (months_diff - 1) + discount_stub_fraction
        else:
            if has_stub:
                period = months_diff + discount_stub_fraction
            else:
                period = months_diff

        return min(period, true_lease_months)

    # Last lease rental gets discounted over the full lock_in horizon
    lease_rental_indices = [i for i, l in enumerate(labels) if l == "Lease Rental"]
    last_lease_rental_idx = lease_rental_indices[-1] if lease_rental_indices else -1

    pv_list = []
    for i, (pay_date, amt, lbl) in enumerate(zip(dates, amounts, labels)):
        if i == last_lease_rental_idx:
            period = true_lease_months
        else:
            period = get_period_for_pv(pay_date, lbl, true_lease_months)
        pv = amt / ((1 + monthly_rate) ** period)
        pv_list.append(pv)

    lease_liability = sum(pv_list)
    ROU_opening = lease_liability

    # -----------------------------------------------------------------------
    # ROU amortization period
    # -----------------------------------------------------------------------
    if exercise_purchase_option and asset_life_months > 0:
        amortization_months = int(asset_life_months)
        st.info(f"ℹ️ ROU asset amortized over **{amortization_months} months** (Life of Asset)")
    else:
        amortization_months = true_lease_months

    amortization_per_month = ROU_opening / amortization_months

    # -----------------------------------------------------------------------
    # Step 5: Build month-by-month schedule
    # -----------------------------------------------------------------------
    payment_map = {}
    for i, (d, amt, lbl) in enumerate(zip(dates, amounts, labels)):
        payment_map.setdefault(d, []).append(
            {"amount": amt, "label": lbl, "pv": pv_list[i]}
        )

    rows = []
    opening_balance = lease_liability
    rou_balance = ROU_opening
    sl_no = 1

    for lm in range(1, lock_in + 1):

        # ── Calendar boundaries for this lease month ──────────────────────
        if lm == 1 and has_stub:
            month_start = t0
            month_end_date = end_of_month
        elif has_stub:
            month_offset = lm - 2
            base_month = end_of_month + timedelta(days=1) + relativedelta(months=month_offset)
            dim = calendar.monthrange(base_month.year, base_month.month)[1]
            month_start = datetime(base_month.year, base_month.month, 1).date()
            month_end_date = datetime(base_month.year, base_month.month, dim).date()
        else:
            base_month = t0 + relativedelta(months=lm - 1)
            dim = calendar.monthrange(base_month.year, base_month.month)[1]
            month_start = datetime(base_month.year, base_month.month, 1).date()
            month_end_date = datetime(base_month.year, base_month.month, dim).date()

        # ── Fraction of month (for interest and ROU amortization) ─────────
        if lm == 1:
            first_payment_date = get_payment_date_for_bucket(1, payment_buckets.get(1, []))

            # ✅ If no time gap → no interest
            if lm == 1:
                days_in_month = calendar.monthrange(t0.year, t0.month)[1]
                days_used = (end_of_month - t0).days + 1

                # ✅ Only 1-day case (31st scenario)
                if days_used <= 1:
                    month_fraction = 0
                else:
                    month_fraction = days_used / days_in_month

        elif lm == lock_in and has_tail_stub:
            month_fraction = last_month_fraction

        else:
            month_fraction = 1.0

        # ── Cash payments that fall within this lease month's window ───────
        cash_payments_this_month = [
            (pay_date, p)
            for pay_date, pay_list in payment_map.items()
            if month_start <= pay_date <= month_end_date
            for p in pay_list
        ]
        cash_payments_this_month.sort(key=lambda x: x[0])

        total_cash_this_month = sum(p["amount"] for _, p in cash_payments_this_month)

        # ── Interest ───────────────────────────────────────────────────────
        if payment_timing.lower() == "beginning":
            interest = (opening_balance - total_cash_this_month) * monthly_rate * month_fraction
        else:
            interest = opening_balance * monthly_rate * month_fraction

        # ── ROU amortization ───────────────────────────────────────────────
        # Force the last lease month to consume the exact remaining ROU balance
        # so the closing figure is precisely 0.00 with no floating-point residual.
        if lm == lock_in:
            rou_amort = rou_balance
        else:
            rou_amort = amortization_per_month * month_fraction

        # ── Emit rows ──────────────────────────────────────────────────────
        if cash_payments_this_month:
            interest_applied = False
            for pay_date, p in cash_payments_this_month:
                payment = p["amount"]
                lbl = p["label"]

                if lbl == "Purchase Option":
                    row_rou_amort = 0.0
                elif not interest_applied:
                    row_rou_amort = rou_amort
                else:
                    row_rou_amort = 0.0

                row_interest = interest if not interest_applied else 0.0
                interest_applied = True

                closing_balance = opening_balance + row_interest - payment
                rou_balance -= row_rou_amort

                rows.append([
                    sl_no,
                    pay_date.strftime("%Y-%m-%d"),
                    lbl,
                    round(payment, 2),
                    round(p["pv"], 2),
                    round(opening_balance, 2),
                    round(payment, 2),
                    round(row_interest, 2),
                    round(closing_balance, 2),
                    round(rou_balance + row_rou_amort, 2),
                    round(row_rou_amort, 2),
                    round(rou_balance, 2),
                ])

                opening_balance = closing_balance
                sl_no += 1

        else:
            # No cash payment this month — accrue interest and amortize ROU
            closing_balance = opening_balance + interest
            rou_balance -= rou_amort

            rows.append([
                sl_no,
                month_end_date.strftime("%Y-%m-%d"),
                "Interest Accrual",
                0.0,
                0.0,
                round(opening_balance, 2),
                0.0,
                round(interest, 2),
                round(closing_balance, 2),
                round(rou_balance + rou_amort, 2),
                round(rou_amort, 2),
                round(rou_balance, 2),
            ])

            opening_balance = closing_balance
            sl_no += 1

    # -----------------------------------------------------------------------
    # Display
    # -----------------------------------------------------------------------
    df = pd.DataFrame(rows, columns=[
        "Sl No.", "Installment Date", "Payment Type", "Amount Paid", "PV of Installment",
        "Opening Lease Liability", "Payment", "Interest", "Closing Lease Liability",
        "ROU Opening", "Amortization", "ROU Closing",
    ])

    if rent_inclusive_of_gst and gst_rate > 0:
        st.info(f"ℹ️ All amounts shown are net of GST @ {gst_rate}%")

    if payment_frequency > 1:
        st.info(
            f"ℹ️ Payment frequency: every **{payment_frequency} months** "
            f"(starting lease month {payment_start_month}). "
            f"Months without cash payments are shown as **Interest Accrual** rows."
        )

    st.success("✅ Lease Schedule Generated Successfully")
    st.dataframe(df, use_container_width=True)

    total_payments = df[df["Payment Type"] == "Lease Rental"]["Amount Paid"].sum()
    total_interest = df["Interest"].sum()
    st.markdown(f"""
    **Schedule Summary:**
    - Initial Lease Liability (PV): ₹{lease_liability:,.2f}
    - Total Lease Rent Payments: ₹{total_payments:,.2f}
    - Total Interest Accrued: ₹{total_interest:,.2f}
    - Payment Frequency: {selected_frequency_label}
    """)

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=buffer,
        file_name="Lease_Schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
