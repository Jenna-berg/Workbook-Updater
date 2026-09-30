        with st.expander("Files read  ({} uploaded, {} hotel(s) found)".format(
                len(pl_files), len(pl_parsed)), expanded=len(pl_parsed) != 1):
            st.dataframe(pd.DataFrame(pl_rows), use_container_width=True, hide_index=True)
            if len(pl_parsed) > 1:
                st.warning(
                    "More than one hotel name was found, so the years are split "
                    "between them. If these are all the same property, the report "
                    "headers differ - check the Hotel column above."
                )

        for msg in pl_problems:
            st.warning("Skipped - " + msg)

        if not pl_parsed:
            st.error("No statements could be read.")
        else:
            hotel_pick = st.selectbox("Hotel", sorted(pl_parsed), key="pl_hotel")
            per_year = pl_parsed[hotel_pick]
            years = sorted(per_year)
            years = years[-PL.MAX_YEARS:]
            per_year = {y: per_year[y] for y in years}
            st.success("**{}** - {} year(s): {}".format(
                hotel_pick, len(years), ", ".join(str(y) for y in years)))
            if len(years) == 1:
                st.info(
                    "Each operating statement covers **one year**, so one file gives "
                    "one column. To get 10 years, upload 10 statements for this hotel "
                    "- one per year end. Check the 'Files read' table above to confirm "
                    "every file was picked up and resolved to the year you expect."
                )

            WANT = ["Total Revenue", "Room", "Food & Beverage", "Miscellaneous",
                    "Rental Income", "Operating Profit or Loss",
                    "Net Income or Loss", "A.D.R.", "Occupancy", "REV PAR"]
            recs = []
            for yr in years:
                row = {"Year": yr}
                for ln in per_year[yr]:
                    if ln.page == "Summary" and ln.label in WANT and ln.label not in row:
                        row[ln.label] = ln.act
                        row[ln.label + " (Budget)"] = ln.bud
                recs.append(row)
            pl_df = pd.DataFrame(recs).set_index("Year")

            latest = years[-1]
            mcols = st.columns(4)

            def _pl_metric(col, label, field, money=True):
                if field not in pl_df.columns:
                    return
                act = pl_df.loc[latest, field]
                bkey = field + " (Budget)"
                bud = pl_df.loc[latest, bkey] if bkey in pl_df.columns else None
                delta = None
                if bud not in (None, 0):
                    delta = "{:+.1%} vs budget".format((act - bud) / bud)
                shown = "${:,.0f}".format(act) if money else "{:,.2f}".format(act)
                col.metric(label, shown, delta)

            _pl_metric(mcols[0], "Total Revenue {}".format(latest), "Total Revenue")
            _pl_metric(mcols[1], "Operating Profit {}".format(latest),
                       "Operating Profit or Loss")
            _pl_metric(mcols[2], "Net Income {}".format(latest), "Net Income or Loss")
            _pl_metric(mcols[3], "ADR {}".format(latest), "A.D.R.", money=False)

            with st.expander("Show the underlying numbers"):
                st.dataframe(pl_df.style.format("{:,.2f}"), use_container_width=True)

            st.divider()
            if st.button("Build P&L Workbook", key="pl_build", type="primary"):
                try:
                    buf = io.BytesIO()
                    PL.build_workbook(hotel_pick, per_year, buf, {})
                    safe_name = re.sub(r"[^A-Za-z0-9 -]", "", hotel_pick).strip() or "Hotel"
                    st.download_button(
                        "Download P&L Workbook",
                        data=buf.getvalue(),
                        file_name=safe_name + " - P&L.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="pl_dl",
                    )
                    st.success("Built. Tabs: Summary, Rooms, Food, Beverage, "
                               "Miscellaneous, Fixed Expenses.")
                except Exception as exc:
                    st.error("Build error: {}".format(exc))


# ══════════════════════════════════════════════════════════════════════════════
# 1-Year Projection
# ══════════════════════════════════════════════════════════════════════════════
with tab_projection:
    st.divider()
    # Streamlit runs every tab body on every interaction, so importing the
    # projector here would pull altair and XlsxWriter — about 55 MB — into
    # memory on each page view, whether or not anyone opens this tab. This app
    # already runs close to the limit on Streamlit Cloud, so the import waits
    # behind a click and only happens for someone actually using the tool.
    if not st.session_state.get("projector_open"):
        st.header("1-Year Projection")
        st.caption(
            "Day-by-day rooms and ADR budget for the year ahead, built from a "
            "segmentation pivot export."
        )
        if st.button("Open the Budget Projector", key="projector_open_btn",
                     type="primary"):
            st.session_state["projector_open"] = True
            st.rerun()
    else:
        try:
            from projector import ui as projector_ui
            projector_ui.render()
        except Exception as exc:
            st.error(f"1-Year Projection failed to load: {exc}")
            st.caption(
                "It needs altair, numpy and XlsxWriter — check they installed "
                "with the rest of requirements.txt."
            )


# Everything the inactive section drew went into this placeholder; clear it so
# only the open section reaches the page. Must stay the last statement in the
# file — anything added after it would render into the void.
_offstage.empty()
