import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import streamlit as st
import calendar

# ========== STREAMLIT APP ==========
st.set_page_config(page_title="Lease Schedule Generator", layout="centered")

st.title("📑 Lease Schedule Generator (with Stub Periods)")

# ---- User Inputs ----
lease_start = st.date_input("Lease Start Date", datetime.today())
lease_term = st.number_input("Lease Term (months)", min_value=1, step=1)
installment_amount = st.number_input("Installment Amount", min_value=0.0, step=100.0)
payment_timing = st.radio("Payments Timing", ["Beginning", "End"])
interest_rate = st.number_input("Annual Interest Rate (%)", min_value=0.0, step=0.1)

if st.button("Generate Schedule"):
    # Convert lease start date
    t0 = lease_start
    annual_rate = interest_rate / 100
    daily_rate = annual_rate / 365

    # Step 1: Build Payment Dates with Stub
    dates = []
    amounts = []

    # --- First Payment (stub) ---
    days_in_month = calendar.monthrange(t0.year, t0.month)[1]
    end_of_month = datetime(t0.year, t0.month, days_in_month).date()

    if t0.day != 1 and t0.day != days_in_month:  # stub needed
        stub_days = (end_of_month - t0).days + 1
        prorated = installment_amount * (stub_days / days_in_month)
        dates.append(end_of_month if payment_timing.lower() == "end" else t0)
        amounts.append(prorated)
        start_for_regular = end_of_month + timedelta(days=1)
    else:
        # no stub, first payment is regular
        dates.append(t0 if payment_timing.lower() == "beginning" else end_of_month)
        amounts.append(installment_amount)
        start_for_regular = t0 + relativedelta(months=1)

    # --- Regular Payments ---
    for i in range(1, lease_term):
        pay_month = start_for_regular + relativedelta(months=i-1)
        days_in_month = calendar.monthrange(pay_month.year, pay_month.month)[1]
        if payment_timing.lower() == "beginning":
            pay_date = datetime(pay_month.year, pay_month.month, 1).date()
        else:
            pay_date = datetime(pay_month.year, pay_month.month, days_in_month).date()
        dates.append(pay_date)
        amounts.append(installment_amount)

    # --- Last Payment (stub if not aligned) ---
    lease_end = t0 + relativedelta(months=lease_term) - timedelta(days=1)
    if lease_end.day != calendar.monthrange(lease_end.year, lease_end.month)[1]:
        days_in_month = calendar.monthrange(lease_end.year, lease_end.month)[1]
        stub_days = lease_end.day
        prorated = installment_amount * (stub_days / days_in_month)
        dates[-1] = lease_end  # replace last date with actual end
        amounts[-1] = prorated

    # Step 2: PV Calculation
    pv_list = []
    for d, amt in zip(dates, amounts):
        days = (d - t0).days
        pv = amt / ((1 + daily_rate) ** days) if days > 0 else amt
        pv_list.append(pv)

    lease_liability = sum(pv_list)
    ROU_opening = lease_liability
    amortization = ROU_opening / lease_term

    # Step 3: Build Schedule
    rows = []
    opening_balance = lease_liability
    rou_balance = ROU_opening

    for i in range(len(dates)):
        interest = opening_balance * daily_rate * ((dates[i] - (dates[i-1] if i > 0 else t0)).days)
        if i == 0 and payment_timing.lower() == "beginning":
            interest = 0
        payment = amounts[i]
        closing_balance = opening_balance + interest - payment
        rou_balance -= amortization

        rows.append([
            i+1,
            dates[i].strftime("%Y-%m-%d"),
            round(payment, 2),
            round(pv_list[i], 2),
            round(opening_balance, 2),
            round(payment, 2),
            round(interest, 2),
            round(closing_balance, 2),
            round(ROU_opening, 2),
            round(amortization, 2),
            round(rou_balance, 2)
        ])

        opening_balance = closing_balance

    df = pd.DataFrame(rows, columns=[
        "Sl No.", "Installment Date", "Amount Paid", "PV of Installment",
        "Opening Lease Liability", "Payment", "Interest", "Closing Lease Liability",
        "ROU Opening", "Amortization", "ROU Closing"
    ])

    st.success("✅ Lease Schedule Generated with Stub Periods")
    st.dataframe(df, use_container_width=True)

    # Download option
    st.download_button(
        label="📥 Download Excel",
        data=df.to_excel("Lease_Schedule.xlsx", index=False, engine="openpyxl"),
        file_name="Lease_Schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
