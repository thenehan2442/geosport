# ============================================================
# GEOSPORTS DAILY SCRAPER
#
# Excel workbook contains:
#
#   Summary
#   Scores
#   Percentiles
#   Daily Summary
#
# IMPORTANT:
# If the script runs twice on the same day,
# today's Scores column, Percentiles column,
# and Daily Summary row are REPLACED.
#
# There will only ever be ONE column per date.
# ============================================================


# ============================================================
# IMPORTS
# ============================================================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException
)

from bs4 import BeautifulSoup as Soup
from scipy.stats import norm

import pandas as pd

from pathlib import Path
from datetime import date, datetime

import subprocess
import time
from zoneinfo import ZoneInfo


# ============================================================
# SETTINGS
# ============================================================

URL = "https://geosports.app/groups/HXG2BB"

# Excel file
EXCEL_FILE = Path("geosports_history.xlsx")

# Chrome profile containing your saved GeoSports/Google login
CHROME_PROFILE = r"C:\selenium\geosports"

# Chrome debugging port
DEBUG_PORT = 9222

# Today's date
GEOSPORTS_TIMEZONE = ZoneInfo("America/New_York")

SCRAPE_TIME = datetime.now(GEOSPORTS_TIMEZONE)

TODAY = SCRAPE_TIME.date()

TODAY_ISO = TODAY.isoformat()

TODAY_COL = (
    f"{TODAY.month}/"
    f"{TODAY.day}/"
    f"{str(TODAY.year)[2:]}"
)

print(
    "GeoSports date:",
    TODAY_COL
)

print(
    "Scraped at:",
    SCRAPE_TIME.strftime(
        "%m/%d/%Y %I:%M:%S %p %Z"
    )
)


# ============================================================
# HELPER: MAKE DATE HEADERS CONSISTENT
#
# Handles either:
#
# 8/31/26
# 2026-08-31
# Excel datetime objects
#
# and converts them all to:
#
# 8/31/26
# ============================================================

def normalize_date_header(value):

    try:

        parsed = pd.to_datetime(value)

        return (
            f"{parsed.month}/"
            f"{parsed.day}/"
            f"{str(parsed.year)[2:]}"
        )

    except Exception:

        return str(value)


# ============================================================
# HELPER: SORT DATE COLUMNS CHRONOLOGICALLY
# ============================================================

def sort_date_columns(dataframe):

    if dataframe.empty:
        return dataframe

    ordered_columns = sorted(
        dataframe.columns,
        key=lambda x: pd.to_datetime(x)
    )

    return dataframe.reindex(
        columns=ordered_columns
    )


# ============================================================
# 1. FIND GOOGLE CHROME
# ============================================================

possible_chrome_paths = [

    Path(
        r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    ),

    Path(
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ),
]

CHROME_PATH = None

for path in possible_chrome_paths:

    if path.exists():

        CHROME_PATH = str(path)

        break


if CHROME_PATH is None:

    raise FileNotFoundError(
        "Could not find Google Chrome."
    )


# ============================================================
# 2. CONNECT TO CHROME
#
# If your Selenium Chrome is already open:
#     connect to it
#
# If it is closed:
#     start Chrome automatically using the profile
#     containing your saved login.
# ============================================================

def connect_to_chrome():

    options = webdriver.ChromeOptions()

    options.debugger_address = (
        f"127.0.0.1:{DEBUG_PORT}"
    )

    # --------------------------------------------------------
    # Try connecting to an already-running Chrome first
    # --------------------------------------------------------

    try:

        driver = webdriver.Chrome(
            options=options
        )

        print(
            "Connected to existing Chrome."
        )

        return driver


    except WebDriverException:

        print(
            "Chrome debugging session is not running."
        )

        print(
            "Starting Chrome automatically..."
        )


    # --------------------------------------------------------
    # Launch Chrome with saved profile
    # --------------------------------------------------------

    subprocess.Popen(
        [
            CHROME_PATH,
            f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={CHROME_PROFILE}",
            "--headless=new",
            "--disable-gpu"
        ]
    )


    # --------------------------------------------------------
    # Wait until Chrome is ready
    # --------------------------------------------------------

    for attempt in range(30):

        time.sleep(0.5)

        options = webdriver.ChromeOptions()

        options.debugger_address = (
            f"127.0.0.1:{DEBUG_PORT}"
        )

        try:

            driver = webdriver.Chrome(
                options=options
            )

            print(
                "Chrome started successfully."
            )

            return driver


        except WebDriverException:

            pass


    raise RuntimeError(
        "Chrome opened, but Selenium "
        "could not connect to it."
    )


driver = connect_to_chrome()

wait = WebDriverWait(
    driver,
    20
)


# ============================================================
# 3. OPEN GEOSPORTS GROUP
# ============================================================

driver.get(URL)


# ============================================================
# 4. WAIT FOR STANDINGS PANEL
#
# If this times out on the first-ever run,
# log into GeoSports manually in the Chrome window,
# then rerun the program.
# ============================================================

try:

    wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                "league-period-panel"
            )
        )
    )


except TimeoutException:

    raise RuntimeError(
        "\nGeoSports standings did not load.\n"
        "If this is the first run with this Chrome profile, "
        "log into GeoSports manually in the Chrome window "
        "and then run the program again."
    )


# ============================================================
# 5. EXPLICITLY CLICK TODAY
#
# Important because Chrome may remember that Week
# or Month was selected previously.
# ============================================================

today_button = wait.until(
    EC.element_to_be_clickable(
        (
            By.ID,
            "league-period-tab-today"
        )
    )
)


try:

    today_button.click()

except Exception:

    driver.execute_script(
        "arguments[0].click();",
        today_button
    )


# ============================================================
# 6. VERIFY TODAY IS SELECTED
# ============================================================

wait.until(
    lambda d:
        d.find_element(
            By.ID,
            "league-period-tab-today"
        ).get_attribute(
            "aria-selected"
        ) == "true"
)


print("Today tab selected.")

time.sleep(1)


# ============================================================
# 7. WAIT FOR TODAY'S TABLE
# ============================================================

wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            "#league-period-panel table"
        )
    )
)


# ============================================================
# 8. CLICK "VIEW ALL X MEMBERS"
#
# Works regardless of whether it says:
#
# View all 15 members
# View all 20 members
# View all 57 members
# ============================================================

VIEW_ALL_XPATH = (
    '//*[@id="league-period-panel"]//button'
    '[contains(normalize-space(.), "View all") '
    'and contains(normalize-space(.), "members")]'
)

MEMBERS_BUTTON_XPATH = (
    '//*[@id="league-period-panel"]//button'
    '[contains(normalize-space(.), "members")]'
)


try:

    view_all_button = WebDriverWait(
        driver,
        5
    ).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                VIEW_ALL_XPATH
            )
        )
    )


    print(
        "Clicking:",
        view_all_button.text
    )


    try:

        view_all_button.click()

    except Exception:

        driver.execute_script(
            "arguments[0].click();",
            view_all_button
        )


    # --------------------------------------------------------
    # Wait for table to expand
    # --------------------------------------------------------

    try:

        wait.until(
            lambda d:
                d.find_element(
                    By.XPATH,
                    MEMBERS_BUTTON_XPATH
                ).get_attribute(
                    "aria-expanded"
                ) == "true"
        )

    except TimeoutException:

        # Some versions of the page may rerender the button.
        # Give it a moment before scraping.
        pass


    time.sleep(1)


except TimeoutException:

    print(
        "No View All button found. "
        "All players may already be visible."
    )


# ============================================================
# 9. GET FULL RENDERED HTML
# ============================================================

bs = Soup(
    driver.page_source,
    "html.parser"
)


# ============================================================
# 10. FIND STANDINGS TABLE
# ============================================================

table = bs.select_one(
    "#league-period-panel table"
)


if table is None:

    raise ValueError(
        "Could not find today's standings table."
    )


# ============================================================
# 11. GET TABLE HEADERS
# ============================================================

headers = [

    th.get_text(
        " ",
        strip=True
    )

    for th in table.select(
        "thead th"
    )
]


print(
    "\nColumns found:"
)

print(headers)


# ============================================================
# 12. GET TABLE ROWS
# ============================================================

data = []


for row in table.select(
    "tbody tr"
):

    cells = row.find_all(
        [
            "th",
            "td"
        ]
    )

    values = [

        cell.get_text(
            " ",
            strip=True
        )

        for cell in cells
    ]


    # Only use rows that line up correctly
    if len(values) == len(headers):

        data.append(values)


# ============================================================
# 13. CREATE RAW DATAFRAME
# ============================================================

df = pd.DataFrame(
    data,
    columns=headers
)


print(
    "\nRaw dataframe:"
)

print(df)


# ============================================================
# 14. CLEAN RANK COLUMN
#
# Remove rows where # is not numeric.
# ============================================================

df["#"] = pd.to_numeric(
    df["#"],
    errors="coerce"
)


df = df.dropna(
    subset=["#"]
).copy()


df["#"] = (
    df["#"]
    .astype(int)
)


# ============================================================
# 15. CLEAN PLAYER
#
# Take the first whitespace-separated string.
#
# Example:
#
# Tomh You
#
# becomes
#
# Tomh
# ============================================================

df["Player"] = (
    df["Player"]
    .astype(str)
    .str.split()
    .str[0]
)


# ============================================================
# 16. CLEAN SCORE
#
# Take first whitespace-separated value
# and convert to number.
# ============================================================

df["Score"] = (
    df["Score"]
    .astype(str)
    .str.split()
    .str[0]
)


df["Score"] = pd.to_numeric(
    df["Score"],
    errors="coerce"
)


# ============================================================
# 17. DROP DAYS WIN
# ============================================================

df = df.drop(
    columns=[
        "Days win"
    ],
    errors="ignore"
)


# ============================================================
# 18. REMOVE ROWS WITHOUT SCORE
# ============================================================

df = df.dropna(
    subset=["Score"]
).copy()


# ============================================================
# 19. CHECK PLAYER NAMES ARE UNIQUE
#
# Because Player becomes the permanent Excel index.
# ============================================================

duplicate_players = (
    df.loc[
        df["Player"].duplicated(
            keep=False
        ),
        "Player"
    ]
    .unique()
)


if len(duplicate_players) > 0:

    raise ValueError(
        "The first-word Player cleanup produced "
        "duplicate player names:\n"
        f"{duplicate_players}\n\n"
        "We would need to change the Player-name "
        "cleaning before saving."
    )


# ============================================================
# 20. MAKE # THE INDEX FOR TODAY'S DISPLAY
# ============================================================

df = df.set_index("#")


print(
    "\nCleaned today's standings:"
)

print(df)


# ============================================================
# 21. TODAY'S SCORE DISTRIBUTION
# ============================================================

mean_score = (
    df["Score"]
    .mean()
)


# Population standard deviation
std_score = (
    df["Score"]
    .std(
        ddof=0
    )
)


print(
    "\nToday's average score:",
    round(
        mean_score,
        2
    )
)


print(
    "Today's standard deviation:",
    round(
        std_score,
        2
    )
)


# ============================================================
# 22. CALCULATE NORMAL-DISTRIBUTION PERCENTILES
#
#              Score - Daily Mean
#     Z = -------------------------------
#              Daily Standard Deviation
#
#
#     Percentile = Normal CDF(Z) * 100
# ============================================================

if (
    pd.isna(std_score)
    or
    std_score == 0
):

    df["Percentile"] = 50.0


else:

    z_score = (
        (
            df["Score"]
            -
            mean_score
        )
        /
        std_score
    )


    df["Percentile"] = (
        norm.cdf(
            z_score
        )
        *
        100
    )


df["Percentile"] = (
    df["Percentile"]
    .round(1)
)


print(
    "\nToday's percentile results:"
)

print(
    df[
        [
            "Player",
            "Score",
            "Percentile"
        ]
    ]
)


# ============================================================
# 23. CREATE TODAY'S SCORE SERIES
#
# Player
# Tomh      829
# Kunal     824
# ZZ        743
# ============================================================

today_scores = (
    df
    .reset_index()
    .set_index(
        "Player"
    )["Score"]
)


today_scores.index = (
    today_scores.index
    .astype(str)
)


# ============================================================
# 24. CREATE TODAY'S PERCENTILE SERIES
# ============================================================

today_percentiles = (
    df
    .reset_index()
    .set_index(
        "Player"
    )["Percentile"]
)


today_percentiles.index = (
    today_percentiles.index
    .astype(str)
)


# ============================================================
# 25. LOAD OLD SCORES SHEET
#
# If workbook/sheet is empty or doesn't exist,
# start from scratch.
# ============================================================

if EXCEL_FILE.exists():

    try:

        scores = pd.read_excel(
            EXCEL_FILE,
            sheet_name="Scores",
            index_col=0
        )

        print(
            "\nLoaded existing Scores sheet."
        )

    except Exception:

        scores = pd.DataFrame()

        print(
            "\nNo usable Scores sheet found. "
            "Starting a new one."
        )


else:

    scores = pd.DataFrame()


scores.index = (
    scores.index
    .astype(str)
)


scores.index.name = "Player"


# Normalize old date columns
scores.columns = [

    normalize_date_header(column)

    for column in scores.columns
]


# ============================================================
# 26. LOAD OLD PERCENTILES SHEET
# ============================================================

if EXCEL_FILE.exists():

    try:

        percentiles = pd.read_excel(
            EXCEL_FILE,
            sheet_name="Percentiles",
            index_col=0
        )

        print(
            "Loaded existing Percentiles sheet."
        )

    except Exception:

        percentiles = pd.DataFrame()

        print(
            "No usable Percentiles sheet found. "
            "Starting a new one."
        )


else:

    percentiles = pd.DataFrame()


percentiles.index = (
    percentiles.index
    .astype(str)
)


percentiles.index.name = "Player"


percentiles.columns = [

    normalize_date_header(column)

    for column in percentiles.columns
]


# ============================================================
# 27. REMOVE TODAY'S OLD SCORE COLUMN
#
# THIS IS WHAT MAKES RERUNNING SAFE.
#
# If 8/31/26 already exists:
#
#     delete old 8/31/26
#     insert new 8/31/26
#
# There can only be ONE column for the date.
# ============================================================

if TODAY_COL in scores.columns:

    print(
        f"\nReplacing old Scores column "
        f"for {TODAY_COL}."
    )

    scores = scores.drop(
        columns=[
            TODAY_COL
        ]
    )


# ============================================================
# 28. REMOVE TODAY'S OLD PERCENTILE COLUMN
# ============================================================

if TODAY_COL in percentiles.columns:

    print(
        f"Replacing old Percentiles column "
        f"for {TODAY_COL}."
    )

    percentiles = percentiles.drop(
        columns=[
            TODAY_COL
        ]
    )


# ============================================================
# 29. BUILD COMPLETE PLAYER LIST
#
# This keeps:
#
# - old players
# - today's players
# - players who join in the future
#
# A player who joins later simply has blank historical cells.
# ============================================================

existing_players = []


for player in scores.index:

    if player not in existing_players:

        existing_players.append(
            player
        )


for player in percentiles.index:

    if player not in existing_players:

        existing_players.append(
            player
        )


for player in today_scores.index:

    if player not in existing_players:

        print(
            f"New player detected: {player}"
        )

        existing_players.append(
            player
        )


all_players = existing_players


# ============================================================
# 30. REINDEX BOTH HISTORICAL TABLES
# ============================================================

scores = scores.reindex(
    all_players
)


percentiles = percentiles.reindex(
    all_players
)


scores.index.name = "Player"

percentiles.index.name = "Player"


# ============================================================
# 31. INSERT TODAY'S NEW SCORES
#
# Because the old date was deleted first,
# this is ALWAYS the only TODAY_COL.
# ============================================================

scores[TODAY_COL] = (
    today_scores
    .reindex(
        all_players
    )
)


# ============================================================
# 32. INSERT TODAY'S NEW PERCENTILES
# ============================================================

percentiles[TODAY_COL] = (
    today_percentiles
    .reindex(
        all_players
    )
)


# ============================================================
# 33. SORT DATE COLUMNS
# ============================================================

scores = sort_date_columns(
    scores
)


percentiles = sort_date_columns(
    percentiles
)


percentiles = (
    percentiles
    .round(1)
)


# ============================================================
# 34. LOAD DAILY SUMMARY
# ============================================================

DAILY_SUMMARY_COLUMNS = [

    "Date",
    "Players",
    "Average_Score",
    "Std_Score",
    "Median_Score",
    "High_Score",
    "Low_Score"
]


if EXCEL_FILE.exists():

    try:

        daily_summary = pd.read_excel(
            EXCEL_FILE,
            sheet_name="Daily Summary"
        )

    except Exception:

        daily_summary = pd.DataFrame(
            columns=DAILY_SUMMARY_COLUMNS
        )


else:

    daily_summary = pd.DataFrame(
        columns=DAILY_SUMMARY_COLUMNS
    )


# ============================================================
# 35. NORMALIZE EXISTING DAILY SUMMARY DATES
# ============================================================

if not daily_summary.empty:

    daily_summary["Date"] = (
        daily_summary["Date"]
        .apply(
            normalize_date_header
        )
    )


# ============================================================
# 36. DELETE TODAY'S OLD DAILY SUMMARY ROW
#
# Same overwrite behavior as Scores/Percentiles.
# ============================================================

if (
    not daily_summary.empty
    and
    TODAY_COL in daily_summary["Date"].values
):

    print(
        f"Replacing old Daily Summary row "
        f"for {TODAY_COL}."
    )


daily_summary = (
    daily_summary[
        daily_summary["Date"]
        !=
        TODAY_COL
    ]
    .copy()
)


# ============================================================
# 37. CREATE TODAY'S DAILY SUMMARY
# ============================================================

today_daily_summary = pd.DataFrame(
    {
        "Date": [
            TODAY_COL
        ],

        "Players": [
            len(df)
        ],

        "Average_Score": [
            mean_score
        ],

        "Std_Score": [
            std_score
        ],

        "Median_Score": [
            df["Score"].median()
        ],

        "High_Score": [
            df["Score"].max()
        ],

        "Low_Score": [
            df["Score"].min()
        ]
    }
)


# ============================================================
# 38. ADD TODAY'S DAILY SUMMARY
# ============================================================

daily_summary = pd.concat(
    [
        daily_summary,
        today_daily_summary
    ],
    ignore_index=True
)


# ============================================================
# 39. SORT DAILY SUMMARY BY DATE
# ============================================================

daily_summary["_sort_date"] = (
    pd.to_datetime(
        daily_summary["Date"]
    )
)


daily_summary = (
    daily_summary
    .sort_values(
        "_sort_date"
    )
    .drop(
        columns=[
            "_sort_date"
        ]
    )
    .reset_index(
        drop=True
    )
)


daily_summary[
    [
        "Average_Score",
        "Std_Score",
        "Median_Score",
        "High_Score",
        "Low_Score"
    ]
] = (
    daily_summary[
        [
            "Average_Score",
            "Std_Score",
            "Median_Score",
            "High_Score",
            "Low_Score"
        ]
    ]
    .round(1)
)


# ============================================================
# 40. BUILD SUMMARY FROM SCORES + PERCENTILES
#
# There is NO separate History sheet anymore.
# ============================================================

summary = pd.DataFrame(
    index=all_players
)


summary.index.name = "Player"


# ------------------------------------------------------------
# Participation
# ------------------------------------------------------------

summary["Days_Played"] = (
    scores
    .notna()
    .sum(
        axis=1
    )
)


# ------------------------------------------------------------
# Percentile statistics
# ------------------------------------------------------------

summary["Average_Percentile"] = (
    percentiles
    .mean(
        axis=1,
        skipna=True
    )
)


summary["Median_Percentile"] = (
    percentiles
    .median(
        axis=1,
        skipna=True
    )
)


summary["Best_Percentile"] = (
    percentiles
    .max(
        axis=1,
        skipna=True
    )
)


summary["Worst_Percentile"] = (
    percentiles
    .min(
        axis=1,
        skipna=True
    )
)


summary["Total_Percentile_Points"] = (
    percentiles
    .sum(
        axis=1,
        skipna=True
    )
)


# ------------------------------------------------------------
# Raw score statistics
# ------------------------------------------------------------

summary["Average_Score"] = (
    scores
    .mean(
        axis=1,
        skipna=True
    )
)


summary["Best_Score"] = (
    scores
    .max(
        axis=1,
        skipna=True
    )
)


# ============================================================
# 41. ROUND SUMMARY
# ============================================================

summary_columns_to_round = [

    "Average_Percentile",

    "Median_Percentile",

    "Best_Percentile",

    "Worst_Percentile",

    "Total_Percentile_Points",

    "Average_Score",

    "Best_Score"
]


summary[
    summary_columns_to_round
] = (
    summary[
        summary_columns_to_round
    ]
    .round(1)
)


# ============================================================
# 42. SORT SUMMARY
#
# For now, overall leaderboard is based on
# Average Percentile.
#
# Total_Percentile_Points is also available
# if we decide to reward participation more heavily later.
# ============================================================

summary = (
    summary
    .reset_index()
    .sort_values(
        [
            "Average_Percentile",
            "Days_Played"
        ],
        ascending=[
            False,
            False
        ]
    )
    .reset_index(
        drop=True
    )
)


# ============================================================
# 43. ADD OVERALL RANK
# ============================================================

summary.insert(
    0,
    "Overall_Rank",
    range(
        1,
        len(summary) + 1
    )
)


# ============================================================
# 44. FINAL DUPLICATE-DATE SAFETY CHECK
#
# There must be exactly one today's column
# in Scores and Percentiles.
# ============================================================

if (
    list(
        scores.columns
    ).count(
        TODAY_COL
    )
    !=
    1
):

    raise ValueError(
        f"Scores should contain exactly one "
        f"{TODAY_COL} column."
    )


if (
    list(
        percentiles.columns
    ).count(
        TODAY_COL
    )
    !=
    1
):

    raise ValueError(
        f"Percentiles should contain exactly one "
        f"{TODAY_COL} column."
    )


if (
    list(
        daily_summary["Date"]
    ).count(
        TODAY_COL
    )
    !=
    1
):

    raise ValueError(
        f"Daily Summary should contain exactly "
        f"one {TODAY_COL} row."
    )


# ============================================================
# 45. SAVE EXCEL
#
# Workbook now contains ONLY:
#
# Summary
# Scores
# Percentiles
# Daily Summary
#
# Close the Excel workbook before running the script.
# ============================================================

try:

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl"
    ) as writer:


        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )


        scores.to_excel(
            writer,
            sheet_name="Scores"
        )


        percentiles.to_excel(
            writer,
            sheet_name="Percentiles"
        )


        daily_summary.to_excel(
            writer,
            sheet_name="Daily Summary",
            index=False
        )


except PermissionError:

    raise PermissionError(
        "\nCould not update geosports_history.xlsx.\n"
        "Make sure the Excel workbook is closed "
        "and then run the program again."
    )


# ============================================================
# 46. FINISHED
# ============================================================

print(
    "\n======================================"
)

print(
    "SUCCESS"
)

print(
    "======================================"
)


print(
    f"\nExcel file:\n{EXCEL_FILE.resolve()}"
)


print(
    f"\nDate added/replaced: {TODAY_COL}"
)


print(
    f"Players scraped today: {len(df)}"
)


print(
    f"Total players stored: {len(scores)}"
)


print(
    f"Total dates stored: {len(scores.columns)}"
)


print(
    f"\n{TODAY_COL} appears exactly once "
    "in Scores."
)


print(
    f"{TODAY_COL} appears exactly once "
    "in Percentiles."
)


print(
    f"{TODAY_COL} appears exactly once "
    "in Daily Summary."
)


print(
    "\nCurrent leaderboard:"
)


print(
    summary[
        [
            "Overall_Rank",
            "Player",
            "Days_Played",
            "Average_Percentile",
            "Median_Percentile",
            "Best_Percentile",
            "Total_Percentile_Points"
        ]
    ]
)
driver.quit()