import datetime as dt
import io

import streamlit as st

from data_loader import load_database, NO_DATA
from pricing_engine import LineItem, build_package_totals
from excel_generator import generate_media_package_excel

st.set_page_config(page_title="OOH Media Package Generator", layout="wide")

DB_PATH = "Database.xlsx"
TEMPLATE_PATH = "changan_template.xlsx"  # master template shipped alongside app.py


@st.cache_resource(show_spinner="Loading media database...")
def get_database():
    return load_database(DB_PATH)


def main():
    st.title("📦 OOH Media Package Generator")

    db = get_database()

    if "package_items" not in st.session_state:
        st.session_state.package_items = []  # list[LineItem]

    st.header("1. Client & Package Info")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client / Product name", value="")
    with col2:
        package_date = st.date_input("Package date", value=dt.date.today())

    st.header("2. Add Media")
    c1, c2, c3 = st.columns(3)
    with c1:
        type_of_media = st.selectbox("Type of Media", sorted(db.keys()))
    records = db[type_of_media]
    media_options = sorted(set(r.media_base for r in records))
    with c2:
        media_choice = st.selectbox("Media", media_options if media_options else ["(no data)"])
    matching = [r for r in records if r.media_base == media_choice]

    location_options = sorted(set(r.location for r in matching if r.location != NO_DATA))
    chosen_record = None
    with c3:
        if location_options:
            loc_choice = st.selectbox("Location", location_options)
            loc_matches = [r for r in matching if r.location == loc_choice]
            chosen_record = loc_matches[0] if loc_matches else None
        else:
            chosen_record = matching[0] if matching else None
            st.write("Location: (n/a)")

    if chosen_record:
        with st.expander("Preview database values for this selection", expanded=False):
            st.json({
                "Quantity": chosen_record.quantity,
                "Detail (Min/Loop)": chosen_record.min_loop,
                "Rate Card": chosen_record.rate_card,
                "Agency Rate": chosen_record.agency_rate,
                "Production": chosen_record.production,
                "TAX3 Q1-Q4": [chosen_record.tax3_q1, chosen_record.tax3_q2,
                               chosen_record.tax3_q3, chosen_record.tax3_q4],
                "Remark": chosen_record.remark,
            })

        li_dummy = LineItem(record=chosen_record, period="", duration_months=1)
        if li_dummy.needs_manual_rate_review:
            st.warning(
                "⚠️ This item's rate can't be read as a single number "
                f"(source value: '{chosen_record.rate_card}'). It will be "
                "added with rate cells shown as '-' — resolve the correct "
                "tier manually before sending to the client."
            )

        c4, c5 = st.columns(2)
        with c4:
            period = st.text_input("Period (e.g. 1 Oct - 31 Dec'26)", key="period_input")
        with c5:
            duration = st.selectbox("Duration (Month)", list(range(1, 13)), index=2, key="duration_input")

        if st.button("➕ Add to package", type="primary"):
            if not period:
                st.error("Period is required — Sales must enter this manually.")
            else:
                st.session_state.package_items.append(
                    LineItem(record=chosen_record, period=period, duration_months=duration)
                )
                st.rerun()

    st.header("3. Current Package")
    if not st.session_state.package_items:
        st.info("No media added yet.")
    else:
        for i, li in enumerate(st.session_state.package_items):
            d = li.to_row_dict()
            cols = st.columns([3, 2, 1, 1, 1, 1, 0.5])
            cols[0].write(f"**{d['type_of_media']}** — {d['media']}")
            cols[1].write(d["period"])
            cols[2].write(f"{d['duration']} mo")
            cols[3].write(f"AR: {d['agency_rate']}")
            cols[4].write(f"Cost: {d['total_media_cost']}")
            cols[5].write(f"Invest: {d['total_media_investment']}")
            if cols[6].button("🗑️", key=f"del_{i}"):
                st.session_state.package_items.pop(i)
                st.rerun()

        totals = build_package_totals(st.session_state.package_items)
        st.subheader("Totals")
        t = totals["totals"]
        st.write(
            f"Total Media Cost: **{t['total_media_cost']:,.0f}** THB · "
            f"Production: **{t['production']:,.0f}** THB · "
            f"Total Media Investment: **{t['total_media_investment']:,.0f}** THB"
        )
        if totals["unresolved"]:
            st.warning(
                f"{len(totals['unresolved'])} item(s) excluded from the totals above "
                "because their rate needs manual review (see warnings when added)."
            )

        st.header("4. Generate Excel")
        if st.button("📊 Generate Media Package Excel", type="primary"):
            if not client_name:
                st.error("Client / Product name is required.")
            else:
                out_path = f"/tmp/Media_Package_{client_name}_{package_date.isoformat()}.xlsx"
                generate_media_package_excel(
                    TEMPLATE_PATH, out_path,
                    st.session_state.package_items,
                    client_name=client_name,
                    package_date=package_date,
                )
                with open(out_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Media Package.xlsx",
                        data=f.read(),
                        file_name=f"Media Package_{client_name}_{package_date.strftime('%d.%m.%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                st.success("Generated. Click above to download.")


if __name__ == "__main__":
    main()
