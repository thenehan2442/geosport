import streamlit as st
import pandas as pd
import calendar

from pathlib import Path
from datetime import datetime
from urllib.parse import quote


# ============================================================
# HTML HELPER
# ============================================================
#
# Streamlit's markdown renderer parses standard Markdown before it
# passes raw HTML through. Any line indented 4+ spaces (which happens
# naturally when an HTML string is built from indented f-strings /
# loops) gets treated as an indented Markdown code block, especially
# once a blank line breaks the current paragraph. That escapes the
# HTML instead of rendering it, which is why sections of the page
# were showing literal "<div>...</div>" text instead of styled cards.
#
# Every raw HTML string is passed through this before st.markdown so
# indentation/blank lines can never trigger that behavior.

def render_html(html_str):

    cleaned = "\n".join(
        line.strip()
        for line in html_str.strip("\n").splitlines()
        if line.strip()
    )

    st.markdown(
        cleaned,
        unsafe_allow_html=True
    )


def clean_text(text_str):
    # Same indentation problem as render_html, but for plain markdown
    # text (e.g. st.success/st.error bodies) where blank lines should
    # be kept as paragraph breaks - only the leading indentation from
    # the source code needs to go.
    return "\n".join(
        line.strip()
        for line in text_str.strip("\n").splitlines()
    )


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="GeoSports League",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = Path("geosports_history.xlsx")


# ============================================================
# STYLE
# ============================================================

render_html(
    """
    <style>

    /* -------------------------------------------------------
       PAGE
    ------------------------------------------------------- */

    .stApp {
        background: #f4f6f9;
    }

    .block-container {
        max-width: 1180px;
        padding-top: 1.2rem;
        padding-bottom: 4rem;
    }

    /* -------------------------------------------------------
       HEADER
    ------------------------------------------------------- */

    .league-header {
        background: linear-gradient(135deg, #081a33 0%, #102b50 100%);
        padding: 28px 32px;
        border-radius: 14px;
        margin-bottom: 22px;
        box-shadow: 0 4px 14px rgba(0,0,0,.12);
    }

    .league-title {
        color: white;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
    }

    .league-subtitle {
        color: #b7c6db;
        font-size: 0.92rem;
        margin-top: 6px;
    }

    /* -------------------------------------------------------
       SECTION HEADERS
    ------------------------------------------------------- */

    .section-title {
        color: #102b50;
        font-size: 1.35rem;
        font-weight: 750;
        margin-top: 14px;
        margin-bottom: 4px;
    }

    .section-subtitle {
        color: #6b7280;
        font-size: .88rem;
        margin-bottom: 16px;
    }

    /* -------------------------------------------------------
       KPI CARDS
    ------------------------------------------------------- */

    .kpi-card {
        background: white;
        border: 1px solid #e3e8ef;
        border-radius: 12px;
        padding: 18px 20px;
        min-height: 115px;
        box-shadow: 0 2px 8px rgba(16,43,80,.06);
    }

    .kpi-label {
        color: #7b8492;
        text-transform: uppercase;
        letter-spacing: .6px;
        font-size: .72rem;
        font-weight: 700;
    }

    .kpi-value {
        color: #102b50;
        font-size: 1.75rem;
        font-weight: 800;
        margin-top: 5px;
    }

    .kpi-note {
        color: #8b93a1;
        font-size: .78rem;
        margin-top: 4px;
    }

    /* -------------------------------------------------------
       PLAYER HEADER
    ------------------------------------------------------- */

    .player-header {
        background: white;
        padding: 24px;
        border: 1px solid #e3e8ef;
        border-radius: 14px;
        box-shadow: 0 2px 8px rgba(16,43,80,.06);
        margin-bottom: 18px;
    }

    .player-name {
        color: #102b50;
        font-size: 2rem;
        font-weight: 800;
        margin: 0;
    }

    /* -------------------------------------------------------
       PERFORMANCE BADGES
    ------------------------------------------------------- */

    .badge-good {
        background: #dcfce7;
        color: #166534;
        border-radius: 999px;
        padding: 4px 10px;
        font-weight: 700;
        font-size: .8rem;
    }

    .badge-average {
        background: #fef3c7;
        color: #92400e;
        border-radius: 999px;
        padding: 4px 10px;
        font-weight: 700;
        font-size: .8rem;
    }

    .badge-poor {
        background: #fee2e2;
        color: #991b1b;
        border-radius: 999px;
        padding: 4px 10px;
        font-weight: 700;
        font-size: .8rem;
    }

    /* -------------------------------------------------------
       CALENDAR
    ------------------------------------------------------- */

    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        gap: 7px;
        margin-top: 12px;
    }

    .calendar-header {
        text-align: center;
        font-size: .72rem;
        font-weight: 800;
        text-transform: uppercase;
        color: #788392;
        padding-bottom: 4px;
    }

    .calendar-day {
        min-height: 92px;
        padding: 9px;
        border-radius: 9px;
        border: 1px solid rgba(0,0,0,.10);
        text-decoration: none !important;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: .15s;
        box-shadow: 0 1px 4px rgba(0,0,0,.05);
    }

    .calendar-day:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 10px rgba(0,0,0,.12);
    }

    .calendar-empty {
        min-height: 92px;
    }

    .calendar-number {
        color: #111827;
        font-size: .82rem;
        font-weight: 800;
    }

    .calendar-value {
        color: #111827;
        font-size: .76rem;
        font-weight: 750;
        text-align: center;
    }

    .calendar-label {
        color: #374151;
        font-size: .62rem;
        text-align: center;
    }

    /* -------------------------------------------------------
       STREAMLIT CONTROLS
    ------------------------------------------------------- */

    div[data-testid="stRadio"] > div {
        background: white;
        border: 1px solid #e3e8ef;
        padding: 6px;
        border-radius: 10px;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid #e1e6ed;
        border-radius: 10px;
        overflow: hidden;
    }

    hr {
        border: none;
        border-top: 1px solid #e1e6ed;
    }

    /* -------------------------------------------------------
       MOBILE
    ------------------------------------------------------- */

    @media(max-width: 700px) {

        .block-container {
            padding-left: .7rem;
            padding-right: .7rem;
        }

        .league-header {
            padding: 21px;
        }

        .league-title {
            font-size: 1.55rem;
        }

        .calendar-grid {
            gap: 3px;
        }

        .calendar-day,
        .calendar-empty {
            min-height: 67px;
        }

        .calendar-day {
            padding: 5px;
        }

        .calendar-value {
            font-size: .62rem;
        }

        .calendar-label {
            display: none;
        }

    }

    </style>
    """)


# ============================================================
# LOAD DATA
# ============================================================

# Cache key is the file's last-modified time, not a fixed timer.
# Geosport.py fully rewrites geosports_history.xlsx on every scrape,
# so its mtime only changes when there's actually new data - this
# way the (slower, Excel-parsing) reload only happens when the file
# has genuinely changed, instead of every 60 seconds regardless.
# max_entries keeps memory bounded as the file changes day after day.

@st.cache_data(max_entries=3)
def load_data(file_mtime):

    summary = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Summary"
    )

    scores = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Scores",
        index_col=0
    )

    percentiles = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Percentiles",
        index_col=0
    )

    scores.index = scores.index.astype(str)
    percentiles.index = percentiles.index.astype(str)

    scores = scores.apply(pd.to_numeric, errors="coerce")
    percentiles = percentiles.apply(pd.to_numeric, errors="coerce")

    return summary, scores, percentiles


if not EXCEL_FILE.exists():
    st.error("League data is currently unavailable.")
    st.stop()


summary, scores, percentiles = load_data(EXCEL_FILE.stat().st_mtime)


# ============================================================
# DATE HELPERS
# ============================================================

def parse_date(value):
    return pd.to_datetime(value, errors="coerce")


def display_date(value):

    d = parse_date(value)

    if pd.isna(d):
        return str(value)

    return f"{d.month}/{d.day}/{str(d.year)[2:]}"


valid_dates = sorted(
    [
        c for c in scores.columns
        if not pd.isna(parse_date(c))
    ],
    key=parse_date
)

scores = scores[valid_dates]

percentile_dates = sorted(
    [
        c for c in percentiles.columns
        if not pd.isna(parse_date(c))
    ],
    key=parse_date
)

percentiles = percentiles[percentile_dates]


if len(valid_dates) == 0:
    st.error("No league results are currently available.")
    st.stop()


LATEST_COL = valid_dates[-1]
LATEST_DATE = parse_date(LATEST_COL)


# ============================================================
# COLORS
# ============================================================

def score_color(value, low, high):

    if pd.isna(value):
        return "#e5e7eb"

    if high == low:
        pct = .5
    else:
        pct = (value - low) / (high - low)

    pct = max(0, min(1, pct))

    # red -> gold
    if pct < .5:

        t = pct / .5

        start = (248, 113, 113)
        end = (250, 204, 21)

    # gold -> green
    else:

        t = (pct - .5) / .5

        start = (250, 204, 21)
        end = (74, 222, 128)

    rgb = tuple(
        int(start[i] + (end[i] - start[i]) * t)
        for i in range(3)
    )

    return f"rgb{rgb}"


def percentile_badge(value):

    if value >= 75:
        return "badge-good"

    if value >= 40:
        return "badge-average"

    return "badge-poor"


# ============================================================
# HEADER
# ============================================================

render_html(
    f"""
    <div class="league-header">
        <div class="league-title">
            GeoSports League
        </div>

        <div class="league-subtitle">
            Official league standings • Updated through
            {LATEST_DATE.strftime("%B %d, %Y")}
        </div>
    </div>
    """)


# ============================================================
# NAVIGATION
# ============================================================

query = st.query_params

requested_view = query.get(
    "view",
    "Home"
)

pages = [
    "Home",
    "League Calendar",
    "Player Profiles"
]

if requested_view not in pages:
    requested_view = "Home"


page = st.radio(
    "Navigation",
    pages,
    index=pages.index(requested_view),
    horizontal=True,
    label_visibility="collapsed"
)


# ============================================================
# HOME
# ============================================================

if page == "Home":

    # --------------------------------------------------------
    # TODAY
    # --------------------------------------------------------

    render_html(
        '<div class="section-title">Today\'s Standings</div>')

    render_html(
        f"""
        <div class="section-subtitle">
            {LATEST_DATE.strftime("%A, %B %d, %Y")}
        </div>
        """)


    today = pd.DataFrame({
        "Score": scores[LATEST_COL],
        "Percentile": percentiles[LATEST_COL]
    })

    today = (
        today
        .dropna(subset=["Score"])
        .sort_values("Score", ascending=False)
    )

    today.insert(
        0,
        "Rank",
        range(1, len(today) + 1)
    )

    today = today.reset_index().rename(
        columns={"index": "Player"}
    )


    # --------------------------------------------------------
    # TOP PLAYER CARD
    # --------------------------------------------------------

    if not today.empty:

        leader = today.iloc[0]

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            render_html(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">Today's Leader</div>
                    <div class="kpi-value">{leader["Player"]}</div>
                    <div class="kpi-note">1st place today</div>
                </div>
                """)

        with c2:
            render_html(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">Top Score</div>
                    <div class="kpi-value">{leader["Score"]:.0f}</div>
                    <div class="kpi-note">Raw score</div>
                </div>
                """)

        with c3:
            render_html(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">Percentile</div>
                    <div class="kpi-value">{leader["Percentile"]:.1f}</div>
                    <div class="kpi-note">Difficulty adjusted</div>
                </div>
                """)

        with c4:

            league_avg = today["Score"].mean()

            render_html(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">League Average</div>
                    <div class="kpi-value">{league_avg:.1f}</div>
                    <div class="kpi-note">Today's field</div>
                </div>
                """)


    st.write("")


    st.dataframe(
        today,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Rank": st.column_config.NumberColumn(
                "Rank",
                format="%d"
            ),
            "Score": st.column_config.NumberColumn(
                "Score",
                format="%.0f"
            ),
            "Percentile": st.column_config.NumberColumn(
                "Percentile",
                format="%.1f"
            )
        }
    )


    # --------------------------------------------------------
    # WEEKLY + ALL TIME
    # --------------------------------------------------------

    render_html(
        '<div class="section-title">League Leaders</div>')

    render_html(
        """
        <div class="section-subtitle">
            Weekly and all-time performance summaries.
        </div>
        """)


    seven_day_start = (
        LATEST_DATE
        -
        pd.Timedelta(days=6)
    )


    week_cols = [
        c for c in scores.columns
        if seven_day_start <= parse_date(c) <= LATEST_DATE
    ]


    week_pct_cols = [
        c for c in percentiles.columns
        if seven_day_start <= parse_date(c) <= LATEST_DATE
    ]


    league = pd.DataFrame(
        index=scores.index
    )

    league["7 Day Avg Score"] = (
        scores[week_cols].mean(axis=1)
    )

    league["7 Day Avg Percentile"] = (
        percentiles[week_pct_cols].mean(axis=1)
    )

    league["All-Time Avg Score"] = (
        scores.mean(axis=1)
    )

    league["All-Time Avg Percentile"] = (
        percentiles.mean(axis=1)
    )

    league["Days Played"] = (
        scores.notna().sum(axis=1)
    )


    league = (
        league
        .sort_values(
            "All-Time Avg Percentile",
            ascending=False
        )
        .round(1)
        .reset_index()
        .rename(columns={"index": "Player"})
    )

    league.insert(
        0,
        "Rank",
        range(1, len(league) + 1)
    )


    st.dataframe(
        league,
        hide_index=True,
        use_container_width=True
    )


# ============================================================
# LEAGUE CALENDAR
# ============================================================

elif page == "League Calendar":

    render_html(
        '<div class="section-title">League Calendar</div>')

    render_html(
        """
        <div class="section-subtitle">
            Calendar color represents average league score.
            Select any completed date to view the full results.
        </div>
        """)


    daily_average = scores.mean(axis=0)

    daily_average.index = pd.to_datetime(
        daily_average.index
    )

    daily_average = daily_average.sort_index()


    months = sorted(
        {
            (x.year, x.month)
            for x in daily_average.index
        }
    )


    selected_month = st.selectbox(
        "Season Month",
        months,
        index=len(months) - 1,
        format_func=lambda x:
            datetime(
                x[0],
                x[1],
                1
            ).strftime("%B %Y")
    )


    year, month = selected_month

    low = daily_average.min()
    high = daily_average.max()


    html = '<div class="calendar-grid">'


    for name in [
        "MON",
        "TUE",
        "WED",
        "THU",
        "FRI",
        "SAT",
        "SUN"
    ]:

        html += (
            f'<div class="calendar-header">'
            f'{name}'
            f'</div>'
        )


    cal = calendar.Calendar(
        firstweekday=0
    )


    for week in cal.monthdayscalendar(
        year,
        month
    ):

        for day in week:

            if day == 0:

                html += (
                    '<div class="calendar-empty"></div>'
                )

                continue


            d = pd.Timestamp(
                year,
                month,
                day
            )


            if d in daily_average.index:

                avg = daily_average.loc[d]

                bg = score_color(
                    avg,
                    low,
                    high
                )

                date_query = d.strftime(
                    "%Y-%m-%d"
                )

                html += f"""
                <a
                    class="calendar-day"
                    style="background:{bg};"
                    href="?view=League%20Calendar&date={date_query}"
                    target="_self"
                >
                    <div class="calendar-number">
                        {day}
                    </div>

                    <div>
                        <div class="calendar-value">
                            {avg:.0f}
                        </div>

                        <div class="calendar-label">
                            AVG SCORE
                        </div>
                    </div>
                </a>
                """

            else:

                html += f"""
                <div
                    class="calendar-day"
                    style="background:#edf0f4;"
                >
                    <div class="calendar-number">
                        {day}
                    </div>
                </div>
                """


    html += "</div>"


    render_html(
        html)


    c1, c2, c3 = st.columns(3)

    c1.caption("Lower scoring day")
    c2.caption("Average difficulty")
    c3.caption("Higher scoring day")


    # --------------------------------------------------------
    # CLICKED DAY
    # --------------------------------------------------------

    selected_date = query.get("date")


    if selected_date:

        selected_ts = pd.to_datetime(
            selected_date
        )

        matching = [
            c for c in scores.columns
            if parse_date(c) == selected_ts
        ]


        if matching:

            col = matching[0]

            day_results = pd.DataFrame({
                "Score": scores[col],
                "Percentile": percentiles[col]
            })

            day_results = (
                day_results
                .dropna(subset=["Score"])
                .sort_values(
                    "Score",
                    ascending=False
                )
            )

            day_results.insert(
                0,
                "Rank",
                range(
                    1,
                    len(day_results) + 1
                )
            )

            day_results = (
                day_results
                .reset_index()
                .rename(
                    columns={"index": "Player"}
                )
            )


            st.divider()

            render_html(
                f"""
                <div class="section-title">
                    {selected_ts.strftime("%B %d, %Y")}
                </div>
                """)


            c1, c2, c3 = st.columns(3)

            c1.metric(
                "Players",
                len(day_results)
            )

            c2.metric(
                "Average Score",
                f"{day_results['Score'].mean():.1f}"
            )

            c3.metric(
                "Winning Score",
                f"{day_results['Score'].max():.0f}"
            )


            st.dataframe(
                day_results,
                hide_index=True,
                use_container_width=True
            )


# ============================================================
# PLAYER PROFILES
# ============================================================

elif page == "Player Profiles":

    render_html(
        '<div class="section-title">Player Profiles</div>')

    render_html(
        """
        <div class="section-subtitle">
            Career performance, consistency and daily results.
        </div>
        """)


    players = sorted(
        scores.index.tolist()
    )


    player = st.selectbox(
        "Select Player",
        players
    )


    p_scores = (
        scores.loc[player]
        .dropna()
    )

    p_percentiles = (
        percentiles.loc[player]
        .dropna()
    )


    avg_score = p_scores.mean()

    avg_pct = p_percentiles.mean()

    score_std = p_scores.std(ddof=0)

    pct_std = p_percentiles.std(ddof=0)


    best_date = p_percentiles.idxmax()

    worst_date = p_percentiles.idxmin()

    best_pct = p_percentiles.loc[
        best_date
    ]

    worst_pct = p_percentiles.loc[
        worst_date
    ]


    render_html(
        f"""
        <div class="player-header">
            <div class="player-name">
                {player}
            </div>

            <div style="
                color:#7b8492;
                margin-top:5px;
                font-size:.85rem;
            ">
                {len(p_scores)} recorded league performances
            </div>
        </div>
        """)


    # --------------------------------------------------------
    # CORE METRICS
    # --------------------------------------------------------

    c1, c2, c3, c4 = st.columns(4)


    metrics = [
        (
            c1,
            "Average Score",
            f"{avg_score:.1f}",
            "Career"
        ),
        (
            c2,
            "Average Percentile",
            f"{avg_pct:.1f}",
            "Difficulty adjusted"
        ),
        (
            c3,
            "Score Std Dev",
            f"{score_std:.1f}",
            "Consistency"
        ),
        (
            c4,
            "Percentile Std Dev",
            f"{pct_std:.1f}",
            "Consistency"
        )
    ]


    for column, label, value, note in metrics:

        with column:

            render_html(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">
                        {label}
                    </div>

                    <div class="kpi-value">
                        {value}
                    </div>

                    <div class="kpi-note">
                        {note}
                    </div>
                </div>
                """)


    st.write("")


    # --------------------------------------------------------
    # BEST / WORST
    # --------------------------------------------------------

    c1, c2 = st.columns(2)


    with c1:

        st.markdown(
            "### Best Performance"
        )

        st.success(
            clean_text(f"""
            {best_pct:.1f} percentile

            Score: {p_scores[best_date]:.0f}

            Date: {display_date(best_date)}
            """)
        )


    with c2:

        st.markdown(
            "### Lowest Performance"
        )

        st.error(
            clean_text(f"""
            {worst_pct:.1f} percentile

            Score: {p_scores[worst_date]:.0f}

            Date: {display_date(worst_date)}
            """)
        )


    # --------------------------------------------------------
    # TREND
    # --------------------------------------------------------

    render_html(
        '<div class="section-title">Performance Trend</div>')


    trend = pd.DataFrame({
        "Date": pd.to_datetime(
            p_percentiles.index
        ),
        "Percentile": p_percentiles.values
    })


    trend = trend.sort_values(
        "Date"
    )


    st.line_chart(
        trend,
        x="Date",
        y="Percentile"
    )


    # --------------------------------------------------------
    # PLAYER CALENDAR
    # --------------------------------------------------------

    render_html(
        '<div class="section-title">Performance Calendar</div>')

    render_html(
        """
        <div class="section-subtitle">
            Each date is colored by the player's percentile
            performance that day.
        </div>
        """)


    player_calendar = p_percentiles.copy()

    player_calendar.index = pd.to_datetime(
        player_calendar.index
    )


    months = sorted(
        {
            (x.year, x.month)
            for x in player_calendar.index
        }
    )


    selected_player_month = st.selectbox(
        "Month",
        months,
        index=len(months) - 1,
        format_func=lambda x:
            datetime(
                x[0],
                x[1],
                1
            ).strftime("%B %Y"),
        key="player_calendar_month"
    )


    year, month = selected_player_month


    html = '<div class="calendar-grid">'


    for name in [
        "MON",
        "TUE",
        "WED",
        "THU",
        "FRI",
        "SAT",
        "SUN"
    ]:

        html += (
            f'<div class="calendar-header">'
            f'{name}'
            f'</div>'
        )


    cal = calendar.Calendar(
        firstweekday=0
    )


    for week in cal.monthdayscalendar(
        year,
        month
    ):

        for day in week:

            if day == 0:

                html += (
                    '<div class="calendar-empty"></div>'
                )

                continue


            d = pd.Timestamp(
                year,
                month,
                day
            )


            if d in player_calendar.index:

                pct = player_calendar.loc[d]

                bg = score_color(
                    pct,
                    0,
                    100
                )

                html += f"""
                <div
                    class="calendar-day"
                    style="background:{bg};"
                >
                    <div class="calendar-number">
                        {day}
                    </div>

                    <div>
                        <div class="calendar-value">
                            {pct:.0f}
                        </div>

                        <div class="calendar-label">
                            PERCENTILE
                        </div>
                    </div>
                </div>
                """

            else:

                html += f"""
                <div
                    class="calendar-day"
                    style="background:#edf0f4;"
                >
                    <div class="calendar-number">
                        {day}
                    </div>
                </div>
                """


    html += "</div>"


    render_html(
        html)


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "GeoSports League Analytics • "
    "Percentiles are normalized against each day's "
    "league scoring distribution."
)