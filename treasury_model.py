import warnings
import datetime as dt
import calendar
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from pandas.tseries.offsets import *

pd.options.mode.chained_assignment = None
pd.set_option("display.expand_frame_repr", False)
pd.set_option("display.max_rows", None)


def normalize_date_formats(security):
    """
    Helper function that takes in a row representing a single security and normalizes date formatting
    :param:
        security: A singular DataFrame row containing the security to be formatted
    :return:
        security: The same security row with cleaned dates
    """

    # Clean issue date
    if "Issue Date" in security.index:
        issue_date = security["Issue Date"]
        if not pd.isna(issue_date) and not isinstance(issue_date, str):
            security["Issue Date"] = issue_date.date()

    # Clean maturity date
    if "Maturity Date" in security.index:
        maturity_date = security["Maturity Date"]
        if not pd.isna(maturity_date) and not isinstance(maturity_date, str):
            security["Maturity Date"] = maturity_date.date()
        elif not pd.isna(maturity_date) and maturity_date[0].isdigit():
            security["Maturity Date"] = pd.to_datetime(maturity_date).date()

    # Clean coupon payment date(s)
    if "Interest Payable Dates" in security.index:
        interest_dates = security["Interest Payable Dates"]
        if not pd.isna(interest_dates) and not isinstance(interest_dates, str):
            security["Interest Payable Dates"] = interest_dates.date()
        elif not pd.isna(interest_dates) and interest_dates[0].isdigit():
            security["Interest Payable Dates"] = interest_dates.split()

    return security


def read_mspd(file_path, include_bills=True):
    """
    Takes in an MSPD Excel spreadsheet and extracts outstanding debt information into a DataFrame
    :param:
        file_path: String containing path to target MSPD file (.xls)
        include_bills; Boolean identifying whether to read outstanding bills or not; defaults to True
    :return:
        full_mspd_df: DataFrame with all outstanding debt information containing normalized dates/units
        mspd_pub_date: Datetime value with MSPD publication date
    """
    full_mspd_df = pd.read_excel(
        file_path, sheet_name="Marketable", usecols="B:E, G:J, L, N, P", header=None
    )
    full_mspd_df.columns = [
        "Category",
        "CUSIP",
        "Interest Rate",
        "Yield",
        "Issue Date",
        "Maturity Date",
        "Interest Payable Dates",
        "Issued",
        "Inflation Adj",
        "Redeemed",
        "Outstanding",
    ]

    # Extract MSPD publication date from the top of the 'Marketable' tab
    mspd_pub_date = pd.to_datetime(full_mspd_df["Category"][0][54:]).date()

    # Find indices of each security type in MSPD and separate into individual data frames
    debt_cats = (
        "Treasury Bills",
        "Treasury Notes",
        "Treasury Bonds",
        "Treasury Inflation-Protected Securities",
        "Treasury Floating Rate Notes",
    )
    debt_cat_indices = []
    is_next_start = (
        True  # Boolean used to skip over subtotal and grand total rows in MSPD
    )
    for row in range(full_mspd_df.shape[0]):
        # Clean date formats
        full_mspd_df.iloc[row] = normalize_date_formats(full_mspd_df.iloc[row])

        temp_label = full_mspd_df.iat[row, 0]
        if isinstance(temp_label, str):
            if temp_label.startswith(debt_cats) and is_next_start:
                debt_cat_indices.append(row + 2)
                is_next_start = False
            elif temp_label.startswith("Total") and not is_next_start:
                debt_cat_indices.append(row)
                is_next_start = True

    # Filter full MSPD down to just relevant outstanding debt categories based on indices found above
    full_mspd_df = full_mspd_df.iloc[
        np.r_[
            debt_cat_indices[0] : debt_cat_indices[1],
            debt_cat_indices[2] : debt_cat_indices[3],
            debt_cat_indices[4] : debt_cat_indices[5],
            debt_cat_indices[6] : debt_cat_indices[7],
            debt_cat_indices[8] : debt_cat_indices[9],
        ]
    ].reset_index(drop=True)

    # Add category labels to each security
    for row in range(full_mspd_df.shape[0]):
        if row < debt_cat_indices[1] - debt_cat_indices[0]:
            full_mspd_df.iloc[row, 0] = "Bills"
        elif row < debt_cat_indices[3] - debt_cat_indices[0] - (
            debt_cat_indices[2] - debt_cat_indices[1]
        ):
            full_mspd_df.iloc[row, 0] = "Notes"
        elif row < debt_cat_indices[5] - debt_cat_indices[0] - (
            debt_cat_indices[2] - debt_cat_indices[1]
        ) - (debt_cat_indices[4] - debt_cat_indices[3]):
            full_mspd_df.iloc[row, 0] = "Bonds"
        elif row < debt_cat_indices[7] - debt_cat_indices[0] - (
            debt_cat_indices[2] - debt_cat_indices[1]
        ) - (debt_cat_indices[4] - debt_cat_indices[3]) - (
            debt_cat_indices[6] - debt_cat_indices[5]
        ):
            full_mspd_df.iloc[row, 0] = "TIPS"
        else:
            full_mspd_df.iloc[row, 0] = "FRNs"

    # Based on include_bills boolean, drop bills
    if not include_bills:
        full_mspd_df = full_mspd_df[full_mspd_df["Category"] != "Bills"]

    # Standardize units to $ billions
    full_mspd_df["Issued"] = full_mspd_df["Issued"] / 1000
    full_mspd_df["Inflation Adj"] = full_mspd_df["Inflation Adj"] / 1000
    full_mspd_df["Redeemed"] = full_mspd_df["Redeemed"] / 1000
    full_mspd_df["Outstanding"] = full_mspd_df["Outstanding"] / 1000

    # Convert all maturity dates to next business day to account for Treasury weekend/holiday settlement rules
    full_mspd_df["Maturity Date"] = (full_mspd_df["Maturity Date"] + 0 * BDay()).dt.date

    return full_mspd_df, mspd_pub_date


def read_soma(file_path, include_bills=True):
    """
    Takes in a SOMA Excel spreadsheet and extracts outstanding SOMA holdings into a DataFrame
    :param:
        file_path: String containing path to target SOMA file (.csv)
        include_bills; Boolean identifying whether to read SOMA bill holdings or not; defaults to True
    :return:
        full_soma_df: DataFrame with all outstanding SOMA holdings (normalized dates, units, etc.)
        soma_pub_date: Single datetime value with SOMA publication date
    """
    full_soma_df = pd.read_csv(file_path, header=0, usecols=[*range(0, 13)])

    # Extract SOMA publication date using first value of 'As Of Date' column
    soma_pub_date = pd.to_datetime(full_soma_df.iat[0, 0]).date()

    # Based on include_bills boolean, drop bills
    if include_bills:
        full_soma_df = full_soma_df[
            full_soma_df["Security Type"].isin(["Bills", "NotesBonds", "FRNs", "TIPS"])
        ].reset_index(drop=True)
    else:
        full_soma_df = full_soma_df[
            full_soma_df["Security Type"].isin(["NotesBonds", "FRNs", "TIPS"])
        ].reset_index(drop=True)

    # Standardize units to $ billions
    full_soma_df["Par Value"] = full_soma_df["Par Value"] / 1000000000
    full_soma_df["Inflation Compensation"] = (
        full_soma_df["Inflation Compensation"] / 1000000000
    )

    # Remove surrounding quotation marks around CUSIP
    full_soma_df["CUSIP"] = full_soma_df["CUSIP"].str.strip("'")

    return full_soma_df, soma_pub_date


def read_issuance(file_path):
    """
    Reads in issuance table from the discretionary inputs spreadsheet
    :param:
        file_path: String containing path to target assumptions file
    :return:
        full_iss_df: DataFrame containing the issuance table
    """
    # Note: If any new securities are added to the issuance table, the usecols field must be updated in the line below
    full_iss_df = pd.read_excel(
        file_path, sheet_name="Issuance Table", usecols="B:N", skiprows=1, header=0
    )

    full_iss_df.fillna(0, inplace=True)  # Blank cells indicate 0 issuance

    # Replace first two columns (MM/YY) of issuance table with single datetime object
    full_iss_df.rename(columns={"MM": "Month", "YY": "Year"}, inplace=True)
    full_iss_df.insert(
        0, "Date", pd.to_datetime(full_iss_df[["Year", "Month"]].assign(DAY=15))
    )
    full_iss_df.drop(["Year", "Month"], axis=1, inplace=True)

    # Generate monthly totals = sum of all coupon issuance each month
    full_iss_df["Monthly"] = full_iss_df.iloc[:, 1:].sum(axis=1)

    return full_iss_df


def read_disc_assumptions(file_path, date_col):
    """
        Reads in all discretionary assumptions (except issuance table, which is read separately above) from Excel input
        spreadsheet, and does some basic error checking to make sure user inputs are formatted correctly.
    :param:
        file_path: String containing path to target input file
        date_col: DataFrame column containing list of all months in model horizon
    :return:
        fy_funding_need: Array of funding needs for each fiscal year over the full model duration
        qe_path: Array of QE purchases by month over the full model duration
        runoff_caps: Array of monthly SOMA runoff cap for each month over the full model duration
    """
    fy_funding_table = pd.read_excel(
        file_path, sheet_name="FY Funding Needs", usecols="B:C", skiprows=3, header=0
    )
    qe_path = pd.read_excel(
        file_path, sheet_name="QE Path", usecols="B:C", skiprows=3, header=0
    )
    runoff_caps = pd.read_excel(
        file_path, sheet_name="SOMA Runoff Caps", usecols="B:C", skiprows=3, header=0
    )

    # Clean formatting of DataFrames read in from Excel
    fy_funding_table.dropna(axis=0, how="all", inplace=True)
    qe_path.dropna(axis=0, how="all", inplace=True)
    runoff_caps.dropna(axis=0, how="all", inplace=True)
    fy_funding_table["FY"] = fy_funding_table["FY"].astype("int")

    # Store model end date for subsequent error checking
    temp_model_end_date = date_col.iat[-1].date()

    # Error checks for FY funding needs: input file length, non-negative funding needs, date formatting
    if temp_model_end_date > dt.date(int(fy_funding_table["FY"].iat[-1]), 9, 15):
        warnings.warn(
            "Warning: Funding need assumptions have not been provided for all necessary fiscal years. Update "
            "the file 'discretionary_assumptions.xlsx' to ensure a funding need is provided for each fiscal "
            "year of the issuance table and restart the model.",
            stacklevel=10,
        )

    if (fy_funding_table["Funding Need ($ bn)"].values < 0).any():
        warnings.warn(
            "One or more FY funding needs has been entered as a negative number. Check to ensure this is "
            "intended - generally, Treasury borrowing needs should be specified using positive numbers.",
            stacklevel=10,
        )

    if not pd.api.types.is_numeric_dtype(fy_funding_table["FY"]):
        warnings.warn(
            "One or more dates in the FY Funding Need assumptions sheet has been entered incorrectly. Update "
            "this sheet in the file 'discretionary_assumptions.xlsx' and restart the model.",
            stacklevel=10,
        )

    # Error checks for QE path: input file length, non-negative values, date formatting
    if temp_model_end_date > qe_path["Month"].iat[-1]:
        warnings.warn(
            "Warning: QE path has not been provided for all necessary months. Update these inputs in the file"
            " 'discretionary_assumptions.xlsx' to ensure each month has an associated QE path assumption and "
            "restart the model.",
            stacklevel=10,
        )

    if (qe_path["QE Purchases ($ bn)"].values < 0).any():
        warnings.warn(
            "Warning: QE purchases path cannot contain negative values. For QT, use the SOMA Runoff Caps "
            "sheet instead to parametrize the pace of balance sheet reduction. Update the QE path in the "
            "file 'discretionary_assumptions.xlsx' and restart the model.",
            stacklevel=10,
        )

    if not pd.api.types.is_datetime64_any_dtype(qe_path["Month"]):
        warnings.warn(
            "One or more years dates in the QE Path assumptions sheet has been entered incorrectly. Update "
            "this sheet in the file 'discretionary_assumptions.xlsx' and restart the model.",
            stacklevel=10,
        )

    if not (
        qe_path["Month"].dt.day[0] == 15
        and (qe_path["Month"].dt.day == qe_path["Month"].dt.day[0]).all()
    ):
        warnings.warn(
            "One or more months in the QE path assumptions has been entered using a non mid-month date. All "
            "months should specifically be entered as the 15th of that month; update this in the file "
            "'discretionary_inputs.xlsx' and restart the model.",
            stacklevel=10,
        )

    # Error checks for SOMA Runoff Caps: input file length, non-negative values, date formatting
    if temp_model_end_date > runoff_caps["Month"].iat[-1]:
        warnings.warn(
            "Warning: SOMA runoff caps have not been provided for all necessary months. Update these inputs "
            "in the file 'discretionary_assumptions.xlsx' to ensure each month has an associated runoff cap "
            "and restart the model.",
            stacklevel=10,
        )

    if (runoff_caps["Runoff Cap ($ bn)"].values < 0).any():
        warnings.warn(
            "Warning: SOMA Runoff Caps cannot contain negative values. For QE, use the QE Path sheet instead "
            "to specify the pace of QE purchases. Update the runoff caps in the file "
            "'discretionary_assumptions.xlsx' and restart the model.",
            stacklevel=10,
        )

    if not pd.api.types.is_datetime64_any_dtype(runoff_caps["Month"]):
        warnings.warn(
            "One or more years dates in the SOMA Runoff Caps assumptions sheet has been entered incorrectly. "
            "Update this sheet in the file 'discretionary_assumptions.xlsx' and restart the model.",
            stacklevel=10,
        )

    if not (
        runoff_caps["Month"].dt.day[0] == 15
        and (runoff_caps["Month"].dt.day == runoff_caps["Month"].dt.day[0]).all()
    ):
        warnings.warn(
            "One or more months in the SOMA Runoff Caps path assumptions has been entered using a non "
            "mid-month date. All months should specifically be entered as the 15th of that month; update this"
            " in the file 'discretionary_inputs.xlsx' and restart the model.",
            stacklevel=10,
        )

    return fy_funding_table, qe_path, runoff_caps


def read_third_party_inputs(file_path, start_date, end_date, infl_rate_type="Mid"):
    """
        Read in inflation spot rates and forward curve from third party data input spreadsheet
    :param:
        file_path: String containing path to target file
        start_date: Date object containing model jump off point
        end_date: Date object containing end date of model horizon
        infl_rate_type: Must be "Bid"/"Mid"/"Ask", optional parameter to select which spot rate to use; defaults to Mid
    :return:
        spot_rates_dict: Dictionary containing yearly spot rates
        forward_curve: DataFrame containing forward curve matrix
        monthly_borrowing_distribution: DataFrame containing average monthly distribution of fiscal funding
    """
    # Checks if input date matches MSPD date to prevent time-lag errors and throw warning if not
    temp_date = (
        pd.read_excel(file_path, sheet_name="Forward Curve", usecols="C", nrows=3)
        .iat[2, 0]
        .date()
    )
    if not temp_date == start_date:
        warnings.warn(
            "Warning: Input file date does not match MSPD date. Update this value in the file "
            "'market_data.xlsx' now and then restart the program.",
            stacklevel=10,
        )

    # Read each sheet of the inputted file
    forward_curve = pd.read_excel(
        file_path, sheet_name="Forward Curve", usecols="B:T", skiprows=5, nrows=12
    )
    term_premium = pd.read_excel(
        file_path, sheet_name="Term Premium", usecols="B:T", skiprows=5, nrows=12
    )
    infl_swaps_df = pd.read_excel(
        file_path, sheet_name="Inflation Swaps", usecols="B:E", skiprows=5, nrows=16
    )
    monthly_borrowing_distribution = pd.read_excel(
        file_path,
        sheet_name="Monthly Borrowing Distribution",
        usecols="B:C",
        skiprows=3,
        nrows=12,
    )

    # Error checking for market data inputs: N/A values
    if (
        np.concatenate(
            [
                forward_curve[col].astype("str").str.contains(r"#N/A", na=False)
                for col in forward_curve
            ]
        ).sum()
        > 0
    ):
        warnings.warn(
            "One or more entries in the Forward Curve input file is returning an N/A value. Update this"
            "value in the file 'market_data.xlsx' now and then restart the program.",
            stacklevel=10,
        )

    if (
        np.concatenate(
            [
                infl_swaps_df[col].astype("str").str.contains(r"#N/A", na=False)
                for col in infl_swaps_df
            ]
        ).sum()
        > 0
    ):
        warnings.warn(
            "One or more entries in the Inflation Swaps input sheet is returning an N/A value. Update this"
            " value in the file 'market_data.xlsx' now and then restart the program.",
            stacklevel=10,
        )

    # Error checking for monthly borrowing distribution: sum must equal 1
    if monthly_borrowing_distribution["Avg % of Fiscal Funding Need Filled"].sum() != 1:
        warnings.warn(
            "Monthly borrowing distribution of FY funding needs does not sum to 1. This is causing funding "
            "needs to be assigned incorrectly over the course of the model. Update these values in the file "
            "'market_data.xlsx' now and then restart the program.",
            stacklevel=10,
        )

    # Add term premium to forward curve
    forward_curve.iloc[:, 1:] = forward_curve.iloc[:, 1:].add(term_premium.iloc[:, 1:])

    # Pull spot rates into a dictionary based on zero coupon inflation swaps data
    spot_rates_dict = {}
    temp_model_length_yrs = end_date.year - start_date.year + 1
    for row in range(temp_model_length_yrs):
        spot_rates_dict[infl_swaps_df.iloc[row]["Tenor"]] = infl_swaps_df.iloc[row][
            infl_rate_type
        ]

    return spot_rates_dict, forward_curve, monthly_borrowing_distribution


def gen_monthly_borrowing(borrowing_yearly, date_col, monthly_dist):
    """
        Apportion FY borrowing assumptions into monthly borrowing needs based on historic seasonal distribution input
    :param:
        borrowing_yearly: DataFrame of yearly borrowing amounts by fiscal year
        date_col: Column of all months in the model horizon
        monthly_dist: DataFrame containing monthly seasonal distributions of the borrowing need (as %'s)
    :return:
        borrowing_monthly_df: DataFrame of monthly borrowing need
    """
    borrowing_monthly = []

    # Adjust current FY distribution percentages to reflect only the remaining portion of the FY
    first_year_dist = monthly_dist.copy()
    temp_start = first_year_dist.index[
        first_year_dist["Month"] == calendar.month_abbr[date_col[0].month]
    ].tolist()[0]
    first_year_dist = first_year_dist[temp_start:].reset_index(drop=True)
    first_year_dist.iloc[:, 1] = (
        first_year_dist.iloc[:, 1] / first_year_dist.iloc[:, 1].sum()
    )

    for date in date_col:
        # Current FY borrowing needs scaled using previously adjusted current FY distributions
        if date <= dt.date(borrowing_yearly.iat[0, 0], 9, 15):
            temp_month_abbr = calendar.month_abbr[date.month]
            monthly_adj = first_year_dist[
                first_year_dist["Month"] == temp_month_abbr
            ].iat[0, 1]
            borrowing_monthly.append(borrowing_yearly.iat[0, 1] * monthly_adj)
            continue

        # Remaining FYs assign monthly funding need proportional to unadjusted seasonal distribution
        for i in range(1, borrowing_yearly.shape[0]):
            start_date = dt.date(borrowing_yearly.iat[i - 1, 0], 9, 30)
            end_date = dt.date(borrowing_yearly.iat[i, 0], 9, 30)

            if start_date < date <= end_date:
                temp_month_abbr = calendar.month_abbr[date.month]
                monthly_adj = monthly_dist[
                    monthly_dist["Month"] == temp_month_abbr
                ].iat[0, 1]
                borrowing_monthly.append(borrowing_yearly.iat[i, 1] * monthly_adj)

    borrowing_monthly_df = pd.DataFrame(borrowing_monthly, columns=["Funding Need"])

    return borrowing_monthly_df


def gen_fwd_inflation_rates(inflation_spot_rates):
    """
        Compute Ny1Y forward rates for inflation implied by a given set of spot rates; used to scale TIPS each period
    :param:
        inflation_spot_rates: Dictionary of spot rates over the full model horizon
    :return:
        monthly_adj: Array of total inflation change from model jump-off point up to that given month
        monthly_deltas: Array of specific monthly inflation rates (e.g. monthly deltas of the monthly_adj array)
    """
    result_rates = pd.DataFrame.from_dict(
        inflation_spot_rates, orient="index", columns=["Spot"]
    )
    result_rates["Spot"] = result_rates["Spot"].div(
        100
    )  # Convert inputs to percentages
    result_rates["Ny1y Fwd"] = 0.00

    # Set up first year values separately to allow for computation of total inflation in monthly_adj
    temp_monthly_val = (1 + result_rates.iat[0, 0]) ** (1 / 12)
    monthly_adj = [temp_monthly_val]
    monthly_deltas = [temp_monthly_val]

    # First year inflation calculated using 1y Spot rate
    for i in range(11):
        # Convert annualized rates to monthly rates
        temp_monthly_val = (1 + result_rates.iat[0, 0]) ** (1 / 12)

        monthly_adj.append(monthly_adj[-1] * temp_monthly_val)
        monthly_deltas.append(temp_monthly_val)

    # Remaining years using Ny1y Fwd Rates
    for i in range(0, result_rates.shape[0] - 1):
        # Generate Ny1y Fwd Rate using corresponding spot rates
        result_rates.iat[i, 1] = (i + 2) * result_rates.iat[i + 1, 0] - (
            i + 1
        ) * result_rates.iat[i, 0]

        # Convert annualized rates to monthly rates
        temp_monthly_val = (1 + result_rates.iat[i, 1]) ** (1 / 12)

        for j in range(12):
            monthly_adj.append(monthly_adj[-1] * temp_monthly_val)
            monthly_deltas.append(temp_monthly_val)

    return monthly_adj, monthly_deltas


def gen_monthly_coupons(forward_curve, gross_issuance):
    """
    Takes in a matrix representing the forward curve and returns a DataFrame containing implied monthly coupons for
    each future security to be generated implied from the issuance table.
    :param:
        forward_curve: DataFrame representing the forward curve
        gross_issuance: DataFrame containing issuance table
    :return:
        coupon_table: DataFrame in same shape as issuance table containing monthly coupons for each security.
        frn_path: Array containing front of curve values at each month (used to adjust FRNs each period)
    """
    coupon_table = gross_issuance.copy()

    # Drop monthly summation column
    coupon_table.drop("Monthly", axis=1, inplace=True)

    # Generate table of coupon values for each tenor at each month of forward issuance
    for col in range(1, coupon_table.shape[1]):
        temp_coupon_rates = [np.nan] * (
            coupon_table.shape[0] + 1
        )  # +1 to allow for interpolation of final year values

        # Pull forward rates for the specific tenor; if FRNs, pull forward rates for front of curve instead
        if coupon_table.columns[col].split(" ")[-1] == "FRN":
            temp_fwd_rates = forward_curve[forward_curve["Tenor"] == "1M"]
        else:
            temp_fwd_rates = forward_curve[
                forward_curve["Tenor"] == coupon_table.columns[col].split(" ")[0]
            ]

        # Create an array of monthly values (matching issuance length) based on interpolation of forward rates
        temp_loc = 0
        for i in range(1, temp_fwd_rates.shape[1]):
            if temp_fwd_rates.columns[i][-1] == "M":
                temp_loc = int(temp_fwd_rates.columns[i][:-1])
            elif temp_fwd_rates.columns[i][-1] == "Y":
                temp_loc = int(temp_fwd_rates.columns[i][:-1]) * 12

            if temp_loc < len(temp_coupon_rates):
                temp_coupon_rates[temp_loc] = temp_fwd_rates.iat[0, i]
            else:
                break

        temp_coupon_rates = (
            pd.Series(temp_coupon_rates).interpolate().to_numpy().tolist()[:-1]
        )

        coupon_table.iloc[:, col] = temp_coupon_rates

    # Store FRN path to return at end of method before subsequent calculations
    frn_path = coupon_table.iloc[:, -1].copy()

    # Adjust coupon table to account for reopenings and month/tenor combinations with 0 issuance
    for col in range(5, coupon_table.shape[1]):
        for row in range(coupon_table.shape[0]):
            # If there is no issuance for a given tenor/period combination, set coupon to N/A
            if gross_issuance.iat[row, col] == 0:
                coupon_table.iat[row, col] = np.nan
                continue

            # Assign coupons for reopenings to match the originally issued security's coupon
            temp_month = coupon_table.iat[row, 0].date().month

            # 10Y, 20Y, 30Y reopenings
            if len(coupon_table.columns[col]) == 3:
                if row > 1 and (temp_month % 3 == 0 or temp_month % 3 == 1):
                    coupon_table.iat[row, col] = coupon_table.iat[row - 1, col]

            # TIPS reopenings
            elif coupon_table.columns[col][-1] == "P":
                if row > 1 and (
                    temp_month == 3
                    or temp_month == 6
                    or temp_month == 9
                    or temp_month == 12
                ):
                    coupon_table.iat[row, col] = coupon_table.iat[row - 2, col]
                elif row > 3 and (temp_month == 5 or temp_month == 11):
                    coupon_table.iat[row, col] = coupon_table.iat[row - 4, col]
                elif row > 5 and temp_month == 8:
                    coupon_table.iat[row, col] = coupon_table.iat[row - 6, col]

            # FRNs reopenings
            elif coupon_table.columns[col][-1] == "N":
                if row > 1 and (temp_month % 3 == 0 or temp_month % 3 == 2):
                    coupon_table.iat[row, col] = coupon_table.iat[row - 1, col]

    return coupon_table, frn_path


def join_with_soma(os, soma_holdings):
    """
    Joins SOMA data with MSPD data on CUSIP to consolidate into one DataFrame
    :param:
        os: DataFrame containing outstanding stock of debt from MSPD
        soma: DataFrame containing current SOMA holdings from Fed report
    :return:
        results: DataFrame joining MSPD and SOMA based on CUSIP, simplified to only show desired columns
        bills_row: Synthetic row containing total outstanding stock of bills as of the model jump-off point
    """

    # Group MSPD by CUSIP and sum amounts outstanding to combine reopenings into a single security
    simple_os = pd.DataFrame(
        os.groupby(["Category", "CUSIP"], sort=False).agg(
            {
                "Interest Rate": ["first"],
                "Issue Date": ["first"],
                "Maturity Date": ["first"],
                "Issued": ["sum"],
                "Outstanding": ["sum"],
            }
        )
    ).reset_index()

    # Add SOMA inflation compensation field onto par value
    soma_holdings["Par Value"] = soma_holdings[
        ["Par Value", "Inflation Compensation"]
    ].sum(axis=1)

    # Simplify both tables to only include desired columns before joining
    simple_os.columns = [
        "Category",
        "CUSIP",
        "Interest Rate",
        "Issue Date",
        "Maturity Date",
        "Issued",
        "Amt Outstanding",
    ]
    simple_soma = soma_holdings[["CUSIP", "Par Value", "Percent Outstanding"]]

    # Merge SOMA and MSPD DataFrames (outer join to preserve securities not held by Fed)
    results = simple_os.merge(simple_soma, on="CUSIP", how="outer").rename(
        columns={
            "Par Value": "SOMA Par Value",
            "Percent Outstanding": "SOMA Percent Outstanding",
        }
    )

    # Drop securities that existed in SOMA file but not in MSPD (already matured holdings not updated in SOMA yet)
    # For example, given a Dec 30 SOMA and Dec 31 MSPD, drop all securities in SOMA maturing on Dec 31 itself
    results = results[results["Category"].notna()]
    results.fillna(0, inplace=True)

    results["Amt Ex-SOMA"] = results["Amt Outstanding"] - results["SOMA Par Value"]

    # Aggregate all bills into a singular "synthetic bills" row, then drop bills from joined table
    bills_row = pd.DataFrame(
        results[results["Category"] == "Bills"].groupby(["Category"], sort=False).sum()
    ).reset_index()

    results = results[results["Category"] != "Bills"].reset_index(drop=True)

    return results, bills_row


def gen_maturity_table(gross_issuance):
    """
    Takes in a projected issuance table and generates a corresponding table of maturity dates for each security.
    Maturity dates are adjusted to reflect Treasury's rules for weekend/holiday settlements.
    :param:
        gross_issuance: DataFrame containing issuance table
    :return:
        maturities: DataFrame in same shape as issuance table containing maturity dates of each future security
    """
    maturities = gross_issuance.copy()

    # Drop summation column
    maturities.drop("Monthly", axis=1, inplace=True)

    # Loops over issuance table by security tenor
    for col in range(1, maturities.shape[1]):
        for row in range(maturities.shape[0]):
            temp_date = maturities.iat[row, 0].date()
            temp_month = temp_date.month

            # Generate maturity dates only where there is issuance
            # Nested if/else statements account for reopenings (if reopening, assign original issue's maturity date)
            if not gross_issuance.iat[row, col] == 0:

                # Notes; 2Y/5Y/7Y mature EOM and 3Y matures mid-month
                if len(gross_issuance.columns[col]) == 2:
                    temp_mat = dt.date(
                        year=(temp_date.year + int(gross_issuance.columns[col][0])),
                        month=temp_date.month,
                        day=temp_date.day,
                    )
                    if gross_issuance.columns[col][0] != "3":
                        temp_mat = dt.date(
                            year=temp_mat.year,
                            month=temp_mat.month,
                            day=(temp_mat + MonthEnd()).day,
                        )

                # Bonds; 20Y matures EOM and 10Y/30Y mature mid-month
                elif len(gross_issuance.columns[col]) == 3:
                    if temp_month % 3 == 0:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row == 0:
                            if temp_date.month > 1:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:2])
                                    ),
                                    month=(temp_date.month - 1),
                                    day=temp_date.day,
                                )
                            else:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:2])
                                    ),
                                    month=(temp_date.month - 1 + 12),
                                    day=temp_date.day,
                                )
                            if gross_issuance.columns[col][0] == 2:
                                temp_mat = dt.date(
                                    year=temp_mat.year,
                                    month=temp_mat.month,
                                    day=(temp_mat + MonthEnd()).day,
                                )
                        else:
                            temp_mat = maturities.iat[row - 1, col]
                    elif temp_month % 3 == 1:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row == 0:
                            if temp_date.month > 2:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:2])
                                    ),
                                    month=(temp_date.month - 2),
                                    day=temp_date.day,
                                )
                            else:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:2])
                                    ),
                                    month=(temp_date.month - 2 + 12),
                                    day=temp_date.day,
                                )
                            if gross_issuance.columns[col][0] == 2:
                                temp_mat = dt.date(
                                    year=temp_mat.year,
                                    month=temp_mat.month,
                                    day=(temp_mat + MonthEnd()).day,
                                )
                        else:
                            temp_mat = maturities.iat[row - 1, col]
                    else:
                        temp_mat = dt.date(
                            year=(
                                temp_date.year + int(gross_issuance.columns[col][:2])
                            ),
                            month=temp_date.month,
                            day=temp_date.day,
                        )
                        if gross_issuance.columns[col][0] == 2:
                            temp_mat = dt.date(
                                year=temp_mat.year,
                                month=temp_mat.month,
                                day=(temp_mat + MonthEnd()).day,
                            )

                # TIPS; all TIPS mature mid-month
                elif gross_issuance.columns[col][-1] == "P":
                    if temp_month % 3 == 0:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row < 2:
                            temp_mat = dt.date(
                                year=(
                                    temp_date.year
                                    + int(gross_issuance.columns[col][:-5])
                                ),
                                month=(temp_date.month - 2),
                                day=temp_date.day,
                            )
                        else:
                            temp_mat = maturities.iat[row - 2, col]
                    elif temp_month == 5 or temp_month == 11:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row < 4:
                            temp_mat = dt.date(
                                year=(
                                    temp_date.year
                                    + int(gross_issuance.columns[col][:-5])
                                ),
                                month=(temp_date.month - 4),
                                day=temp_date.day,
                            )
                        else:
                            temp_mat = maturities.iat[row - 4, col]
                    elif temp_month == 8:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row < 6:
                            temp_mat = dt.date(
                                year=(
                                    temp_date.year
                                    + int(gross_issuance.columns[col][:-5])
                                ),
                                month=(temp_date.month - 6),
                                day=temp_date.day,
                            )
                        else:
                            temp_mat = maturities.iat[row - 6, col]
                    else:
                        temp_mat = dt.date(
                            year=(
                                temp_date.year + int(gross_issuance.columns[col][:-5])
                            ),
                            month=temp_date.month,
                            day=temp_date.day,
                        )

                # FRNs; all FRNs mature EOM
                else:
                    if temp_month % 3 == 0:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row == 0:
                            if temp_date.month > 2:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:-5])
                                    ),
                                    month=(temp_date.month - 2),
                                    day=temp_date.day,
                                )
                            else:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:-5])
                                    ),
                                    month=(temp_date.month - 2 + 12),
                                    day=temp_date.day,
                                )
                            temp_mat = dt.date(
                                year=temp_mat.year,
                                month=temp_mat.month,
                                day=(temp_mat + MonthEnd()).day,
                            )
                        else:
                            temp_mat = maturities.iat[row - 1, col]

                    elif temp_month % 3 == 2:
                        # Handles case where first occurrence of a security in the issuance table is a reopening
                        if row == 0:
                            if temp_date.month > 1:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:-5])
                                    ),
                                    month=(temp_date.month - 1),
                                    day=temp_date.day,
                                )
                            else:
                                temp_mat = dt.date(
                                    year=(
                                        temp_date.year
                                        + int(gross_issuance.columns[col][:-5])
                                    ),
                                    month=(temp_date.month - 1 + 12),
                                    day=temp_date.day,
                                )
                            temp_mat = dt.date(
                                year=temp_mat.year,
                                month=temp_mat.month,
                                day=(temp_mat + MonthEnd()).day,
                            )
                        else:
                            temp_mat = maturities.iat[row - 1, col]
                    else:
                        temp_mat = dt.date(
                            year=(
                                temp_date.year + int(gross_issuance.columns[col][:-5])
                            ),
                            month=temp_date.month,
                            day=temp_date.day,
                        )
                        temp_mat = dt.date(
                            year=temp_mat.year,
                            month=temp_mat.month,
                            day=(temp_mat + MonthEnd()).day,
                        )

                temp_mat = (
                    temp_mat + 0 * BDay()
                ).date()  # Set maturity date to next business day if weekend/holiday
                maturities.iat[row, col] = temp_mat

    # For months with 0 issuance, set maturity date to N/A
    maturities.replace(0, np.nan, inplace=True)

    return maturities


def gen_future_debt_stock(gross_issuance, coupon_table):
    """
    Takes in a table of projected issuance and generates a table of all the implied future securities. These are
    formatted to match the structure of securities in the MSPD.
    :param:
        gross_issuance: DataFrame containing issuance table
        coupon_table: DataFrame in same shape as issuance table containing coupons to be assigned to each tenor
    :return:
        future_stock_df: DataFrame containing all future securities implied by the issuance table in MSPD form
    """
    future_stock = []
    maturity_table = gen_maturity_table(gross_issuance)

    # Loops over issuance table by column/security type and generate a new CUSIP security row for each entry
    for col in range(1, gross_issuance.shape[1] - 1):
        for row in range(gross_issuance.shape[0]):
            temp_date = gross_issuance.iat[row, 0].date()
            temp_month = temp_date.month

            temp_issue = temp_date
            temp_cusip = "CUSIP" + str(row) + str(col)
            temp_coupon = coupon_table.iat[row, col]
            temp_mat = maturity_table.iat[row, col]

            # Generate future securities only where there is issuance
            # Nested if/else statements account for reopenings (if reopening, assign original issue's CUSIP/maturity)
            if not (gross_issuance.iat[row, col] == 0):

                # Notes (2Y, 5Y, 7Y); 3Y notes are issued mid-month, 2Y/5Y/7Y notes are issued EOM
                if len(gross_issuance.columns[col]) == 2:
                    temp_cat = "Notes"
                    if gross_issuance.columns[col][0] != "3":
                        temp_issue = dt.date(
                            year=temp_date.year,
                            month=temp_date.month,
                            day=(temp_date + MonthEnd()).day,
                        )
                    temp_issue = (temp_issue + 0 * BDay()).date()

                # Bonds (10Y, 20Y, 30Y); 10Y/30Y bonds are issued mid-month, 20Y bonds are issued EOM
                elif len(gross_issuance.columns[col]) == 3:
                    temp_cat = "Bonds"
                    # Handles case where first occurrence of nominal bond is a reopening
                    if temp_month % 3 == 0 or temp_month % 3 == 1:
                        temp_cusip = future_stock[-1].get("CUSIP")

                    if gross_issuance.columns[col][:2] == "20":
                        temp_issue = dt.date(
                            year=temp_date.year,
                            month=temp_date.month,
                            day=(temp_date + MonthEnd()).day,
                        )
                    temp_issue = (temp_issue + 0 * BDay()).date()

                # TIPS; all TIPS are issued on the last business day of the month
                elif gross_issuance.columns[col][-1] == "P":
                    temp_cat = "TIPS"
                    # Handles case where first occurrence of a TIPS is a reopening
                    if not (temp_month == 1 or temp_month == 2 or temp_month == 7):
                        temp_cusip = future_stock[-1].get("CUSIP")

                    temp_issue = pd.date_range(temp_date, periods=1, freq="BM")[
                        0
                    ].date()

                # FRNs; FRNs are issued EOM for new issues and reopened on the last business Friday of the month
                else:
                    temp_cat = "FRNs"
                    temp_issue = dt.date(
                        year=temp_date.year,
                        month=temp_date.month,
                        day=(temp_date + MonthEnd()).day,
                    )
                    temp_issue = (temp_issue + 0 * BDay()).date()

                    # Handles case where first occurrence of an FRN is a reopening
                    if temp_month % 3 == 0 or temp_month % 3 == 2:
                        temp_cusip = future_stock[-1].get("CUSIP")
                        while (
                            temp_issue.weekday() != 4
                        ):  # While not a business Friday, decrement by a business day
                            temp_issue = (temp_issue - BDay(1)).date()

                # Consolidate all fields for the newly generated security into a dictionary and append to future stock
                temp_dict = {
                    "Category": temp_cat,
                    "CUSIP": temp_cusip,
                    "Interest Rate": temp_coupon,
                    "Issue Date": temp_issue,
                    "Maturity Date": temp_mat,
                    "Issued": gross_issuance.iat[row, col],
                    "Amt Outstanding": gross_issuance.iat[row, col],
                    "SOMA Par Value": 0.00,
                    "SOMA Percent Outstanding": 0.00,
                    "Amt Ex-SOMA": gross_issuance.iat[row, col],
                }
                future_stock.append(temp_dict)

    # Convert future_stock from array of dictionaries (each dict representing one future CUSIP) into a DataFrame
    future_stock_df = pd.DataFrame(future_stock)

    return future_stock_df


def gen_characteristics(
    os, as_of_date, current_os_bills_row, frn_value, include_bills=True
):
    """
    Computes desired summary statistics and duration supply buckets on the outstanding debt stock. All summary stats
    and supply buckets are computed both inclusive/exclusive of SOMA.
    :param:
        os: DataFrame containing current outstanding debt with joined SOMA holding information
        as_of_date: Date of the outstanding stock
        current_os_bills_row: DataFrame containing current outstanding bills with joined SOMA holding information
        frn_value: Current period front of curve yield, used to calculate SOMA-adjusted WAC
        include_bills: Optional parameter to compute summary stats with or without bills, defaults to True
    :return:
        stats_incl_soma: Dictionary containing summary statistics including SOMA holdings
        buckets_incl_soma: Dictionary containing supply buckets including SOMA holdings
        stats_ex_soma: Dictionary containing summary statistics excluding SOMA holdings
        buckets_ex_soma: Dictionary containing supply buckets excluding SOMA holdings
    """
    current_os = os.copy()

    # Synthetic bills row to outstanding stock for summary stat computation
    bills_row = {
        "Category": "Bills",
        "CUSIP": "SynthBill1",
        "Interest Rate": 0.00,
        "Issue Date": as_of_date.date(),
        "Maturity Date": (as_of_date + dt.timedelta(days=30.5)).date(),
        "Issued": current_os_bills_row["Issued"].sum(),
        "Amt Outstanding": current_os_bills_row["Amt Outstanding"].sum(),
        "SOMA Par Value": current_os_bills_row["SOMA Par Value"].sum(),
        "SOMA Percent Outstanding": current_os_bills_row[
            "SOMA Percent Outstanding"
        ].sum(),
        "Amt Ex-SOMA": current_os_bills_row["Amt Ex-SOMA"].sum(),
    }
    # Only include bills row if specified by input boolean
    if include_bills:
        current_os = current_os.append(bills_row, ignore_index=True)

    # Amt Outstanding
    amt_out_incl_soma = current_os["Amt Outstanding"].sum()
    amt_out_ex_soma = current_os["Amt Ex-SOMA"].sum()

    # WAM
    wt_mat = (
        pd.to_datetime(current_os["Maturity Date"]) - pd.to_datetime(as_of_date)
    ).dt.days / 365.25
    wt_mat_incl_soma = current_os["Amt Outstanding"].multiply(wt_mat, axis=0).sum()
    wt_mat_incl_soma = (
        wt_mat_incl_soma * 12 / amt_out_incl_soma
    )  # *12 to make WAM monthly
    wt_mat_ex_soma = current_os["Amt Ex-SOMA"].multiply(wt_mat, axis=0).sum()
    wt_mat_ex_soma = wt_mat_ex_soma * 12 / amt_out_ex_soma

    # Truncated WAM
    trunc_wt_mat = wt_mat.clip(upper=10.000000)
    trunc_wt_mat_incl_soma = (
        current_os["Amt Outstanding"].multiply(trunc_wt_mat, axis=0).sum()
    )
    trunc_wt_mat_incl_soma = (
        trunc_wt_mat_incl_soma * 12 / amt_out_incl_soma
    )  # *12 to make Trunc WAM monthly
    trunc_wt_mat_ex_soma = (
        current_os["Amt Ex-SOMA"].multiply(trunc_wt_mat, axis=0).sum()
    )
    trunc_wt_mat_ex_soma = trunc_wt_mat_ex_soma * 12 / amt_out_ex_soma

    # WAD (assuming bill and FRN duration = 0)
    is_not_bill_frn = ~(current_os["Category"].isin(["Bills", "FRNs"]))
    num_pers = round(
        (
            pd.to_datetime(current_os["Maturity Date"]) - pd.to_datetime(as_of_date)
        ).dt.days
        / 365.25
        * 2
    )
    wt_dur = (1 / (current_os["Interest Rate"] / 100)) * (
        1 - 1 / (1 + (current_os["Interest Rate"] / 200)).pow(num_pers, axis=0)
    )
    wt_dur = wt_dur.multiply(is_not_bill_frn, axis=0)
    wt_dur_incl_soma = current_os["Amt Outstanding"].multiply(wt_dur, axis=0).sum()
    wt_dur_incl_soma = wt_dur_incl_soma / amt_out_incl_soma
    wt_dur_ex_soma = current_os["Amt Ex-SOMA"].multiply(wt_dur, axis=0).sum()
    wt_dur_ex_soma = wt_dur_ex_soma / amt_out_ex_soma

    # SOMA-adjusted WAD, treating all SOMA holdings as FRNs (duration = 0)
    wt_dur_soma_adj = (wt_dur_ex_soma * amt_out_ex_soma) / amt_out_incl_soma

    # WAC
    wt_coup_incl_soma = (
        current_os["Amt Outstanding"]
        .multiply(current_os["Interest Rate"], axis=0)
        .sum()
    )
    wt_coup_incl_soma = wt_coup_incl_soma / amt_out_incl_soma
    wt_coup_ex_soma = (
        current_os["Amt Ex-SOMA"].multiply(current_os["Interest Rate"], axis=0).sum()
    )
    wt_coup_ex_soma = wt_coup_ex_soma / amt_out_ex_soma

    # SOMA-adjusted WAC, treating SOMA holdings as FRNs (coupon = front of curve in the given period)
    amt_out_soma = current_os["SOMA Par Value"].sum()
    wt_coup_soma = current_os["SOMA Par Value"].multiply(frn_value, axis=0).sum()
    wt_coup_soma = wt_coup_soma / amt_out_soma
    wt_coup_soma_adj = (
        wt_coup_ex_soma * amt_out_ex_soma + wt_coup_soma * amt_out_soma
    ) / amt_out_incl_soma

    # TIPS Share
    tips_out = current_os[current_os["Category"] == "TIPS"]
    tips_share_incl_soma = tips_out["Amt Outstanding"].sum() / amt_out_incl_soma
    tips_share_ex_soma = tips_out["Amt Ex-SOMA"].sum() / amt_out_ex_soma

    # FRNs Share
    frns_out = current_os[current_os["Category"] == "FRNs"]
    frns_share_incl_soma = frns_out["Amt Outstanding"].sum() / amt_out_incl_soma
    frns_share_ex_soma = frns_out["Amt Ex-SOMA"].sum() / amt_out_ex_soma

    # T+1 Mat. (share of debt maturing within the next 1 year)
    mat_lt_curr = current_os[
        current_os["Maturity Date"] <= (as_of_date + dt.timedelta(days=370.25))
    ]
    t_1_mat_incl_soma = mat_lt_curr["Amt Outstanding"].sum()
    t_1_mat_ex_soma = mat_lt_curr["Amt Ex-SOMA"].sum()

    # T+3 Mat.
    mat_lt_curr = current_os[
        current_os["Maturity Date"] <= (as_of_date + dt.timedelta(days=1100.75))
    ]
    t_3_mat_incl_soma = mat_lt_curr["Amt Outstanding"].sum()
    t_3_mat_ex_soma = mat_lt_curr["Amt Ex-SOMA"].sum()

    # T+5 Mat.
    mat_lt_curr = current_os[
        current_os["Maturity Date"] <= (as_of_date + dt.timedelta(days=1831.25))
    ]
    t_5_mat_incl_soma = mat_lt_curr["Amt Outstanding"].sum()
    t_5_mat_ex_soma = mat_lt_curr["Amt Ex-SOMA"].sum()

    # T+10 Mat.
    mat_lt_curr = current_os[
        current_os["Maturity Date"] <= (as_of_date + dt.timedelta(days=3657.50))
    ]
    t_10_mat_incl_soma = mat_lt_curr["Amt Outstanding"].sum()
    t_10_mat_ex_soma = mat_lt_curr["Amt Ex-SOMA"].sum()

    # 2-10 Belly Mat.
    t_2_10_mat_belly = mat_lt_curr[
        mat_lt_curr["Maturity Date"] > (as_of_date + dt.timedelta(days=735.50))
    ]
    t_2_10_mat_incl_soma = t_2_10_mat_belly["Amt Outstanding"].sum()
    t_2_10_mat_ex_soma = t_2_10_mat_belly["Amt Ex-SOMA"].sum()

    stats_incl_soma = {
        "Period": calendar.month_abbr[period_date.month] + " " + str(period_date.year),
        "Amt Outstanding": amt_out_incl_soma,
        "Bill Share": current_os_bills_row["Amt Outstanding"].sum() / amt_out_incl_soma,
        "WAM": wt_mat_incl_soma,
        "Truncated WAM": trunc_wt_mat_incl_soma,
        "WAC": wt_coup_incl_soma,
        "WAD": wt_dur_incl_soma,
        "T+1 Mat": t_1_mat_incl_soma,
        "T+3 Mat": t_3_mat_incl_soma,
        "T+5 Mat": t_5_mat_incl_soma,
        "T+10 Mat": t_10_mat_incl_soma,
        "T 2-10 Mat": t_2_10_mat_incl_soma,
        "TIPS Share": tips_share_incl_soma,
        "SOMA-Adjusted WAC": wt_coup_soma_adj,
        "SOMA-Adjusted WAD": wt_dur_soma_adj,
        "FRNs Amt Outstanding": frns_share_incl_soma,
    }

    stats_ex_soma = {
        "Period": calendar.month_abbr[period_date.month] + " " + str(period_date.year),
        "Amt Outstanding": amt_out_ex_soma,
        "Bill Share": current_os_bills_row["Amt Ex-SOMA"].sum() / amt_out_ex_soma,
        "WAM": wt_mat_ex_soma,
        "Truncated WAM": trunc_wt_mat_ex_soma,
        "WAC": wt_coup_ex_soma,
        "WAD": wt_dur_ex_soma,
        "T+1 Mat": t_1_mat_ex_soma,
        "T+3 Mat": t_3_mat_ex_soma,
        "T+5 Mat": t_5_mat_ex_soma,
        "T+10 Mat": t_10_mat_ex_soma,
        "T 2-10 Mat": t_2_10_mat_ex_soma,
        "TIPS Share": tips_share_ex_soma,
        "FRNs Amt Outstanding": frns_share_ex_soma,
    }

    # Calculation of notinoal and duration supply buckets
    bucket_vals_incl_soma = []
    bucket_vals_ex_soma = []
    interval_range = [0, 2, 5, 7, 10, 20, 30]
    for i in range(1, len(interval_range)):
        # Filter amount outstanding to only securities within the specific maturity bucket
        start_days = 365.25 * interval_range[i - 1]
        end_days = 365.25 * interval_range[i]
        mat_gt = current_os["Maturity Date"] > (
            as_of_date + dt.timedelta(days=start_days + 27)
        )
        mat_lt = current_os["Maturity Date"] <= (
            as_of_date + dt.timedelta(days=end_days + 27)
        )
        in_range = [all(tup) for tup in zip(mat_gt, mat_lt)]

        # Calculated weighted duration of each bucket
        bucket_wt_dur = wt_dur.multiply(in_range, axis=0)
        wt_dur_incl_soma = (
            current_os["Amt Outstanding"].multiply(bucket_wt_dur, axis=0).sum()
        )
        wt_dur_ex_soma = current_os["Amt Ex-SOMA"].multiply(bucket_wt_dur, axis=0).sum()

        # Append bucket information to each array
        bucket_vals_incl_soma.append(
            current_os["Amt Outstanding"].multiply(in_range, axis=0).sum()
        )
        bucket_vals_incl_soma.append(
            wt_dur_incl_soma
            / current_os["Amt Outstanding"].multiply(in_range, axis=0).sum()
        )
        bucket_vals_ex_soma.append(
            current_os["Amt Ex-SOMA"].multiply(in_range, axis=0).sum()
        )
        bucket_vals_ex_soma.append(
            wt_dur_ex_soma / current_os["Amt Ex-SOMA"].multiply(in_range, axis=0).sum()
        )

    buckets_incl_soma = {
        "Period": calendar.month_abbr[period_date.month] + " " + str(period_date.year),
        "0-2Y Supply": bucket_vals_incl_soma[0],
        "0-2Y Duration": bucket_vals_incl_soma[1],
        "2-5Y Supply": bucket_vals_incl_soma[2],
        "2-5Y Duration": bucket_vals_incl_soma[3],
        "5-7Y Supply": bucket_vals_incl_soma[4],
        "5-7Y Duration": bucket_vals_incl_soma[5],
        "7-10Y Supply": bucket_vals_incl_soma[6],
        "7-10Y Duration": bucket_vals_incl_soma[7],
        "10-20Y Supply": bucket_vals_incl_soma[8],
        "10-20Y Duration": bucket_vals_incl_soma[9],
        "20-30Y Supply": bucket_vals_incl_soma[10],
        "20-30Y Duration": bucket_vals_incl_soma[11],
    }

    buckets_ex_soma = {
        "Period": calendar.month_abbr[period_date.month] + " " + str(period_date.year),
        "0-2Y Supply": bucket_vals_ex_soma[0],
        "0-2Y Duration": bucket_vals_ex_soma[1],
        "2-5Y Supply": bucket_vals_ex_soma[2],
        "2-5Y Duration": bucket_vals_ex_soma[3],
        "5-7Y Supply": bucket_vals_ex_soma[4],
        "5-7Y Duration": bucket_vals_ex_soma[5],
        "7-10Y Supply": bucket_vals_ex_soma[6],
        "7-10Y Duration": bucket_vals_ex_soma[7],
        "10-20Y Supply": bucket_vals_ex_soma[8],
        "10-20Y Duration": bucket_vals_ex_soma[9],
        "20-30Y Supply": bucket_vals_ex_soma[10],
        "20-30Y Duration": bucket_vals_ex_soma[11],
    }

    return stats_incl_soma, buckets_incl_soma, stats_ex_soma, buckets_ex_soma


if __name__ == "__main__":
    # Read in external file inputs
    outstanding_stock, mspd_date = read_mspd("input_data/mspd/MSPD_Mar2022.xls")
    soma, soma_date = read_soma("input_data/soma/SOMA_Mar302022.csv")
    gross_iss = read_issuance("input_data/discretionary_assumptions.xlsx")
    funding_need_fy, qe_purchases, soma_runoff_caps = read_disc_assumptions(
        "input_data/discretionary_assumptions.xlsx", gross_iss["Date"]
    )
    inflation_spots, fwd_curve, monthly_budget_dist = read_third_party_inputs(
        "input_data/market_data.xlsx", mspd_date, gross_iss["Date"].iat[-1]
    )

    # If MSPD publication date is not a business day, throw a warning about potentially unsettled issuance
    if not bool(len(pd.bdate_range(mspd_date, mspd_date))):
        warnings.warn(
            "WARNING: MSPD Publication Date is not a business day, so certain EOM issuance that hasn't yet "
            "settled may be missing from the MSPD. Check to make sure this is not a problem, or update the "
            "MSPD file to a business day MSPD instead and restart the program.",
            stacklevel=10,
        )

    # Remove rows from gross issuance table that are before jump-off date, using MSPD date as the jump-off point
    for index, r in gross_iss.iterrows():
        if r["Date"].date() < mspd_date:
            gross_iss.drop(index, inplace=True)
    gross_iss.reset_index(inplace=True, drop=True)

    # Distribute FY funding need assumptions across months based on seasonal distribution history
    funding_need_monthly = gen_monthly_borrowing(
        funding_need_fy, gross_iss["Date"], monthly_budget_dist
    )

    # Calculate monthly inflation adjustments used to scale TIPS from inflation spot rate inputs
    monthly_infl_adj, monthly_infl_deltas = gen_fwd_inflation_rates(inflation_spots)

    # From forward curve inputs, generate a table containing coupon rates for each future security to be issued
    coup_table, frn_rates = gen_monthly_coupons(fwd_curve, gross_iss)

    # Join MSPD and SOMA on CUSIP (bills stored separately in outstanding_bills variable)
    joined_os, outstanding_bills = join_with_soma(outstanding_stock, soma)

    # Generate CUSIPs of future securities, then join with existing stock to get full universe over the model horizon
    fut_stock_cusips = gen_future_debt_stock(gross_iss, coup_table)
    complete_universe = pd.concat(
        [joined_os, fut_stock_cusips],
        axis=0,
        join="outer",
        sort=False,
        ignore_index=True,
    )
    complete_universe = complete_universe.sort_values(
        by=["Category", "Maturity Date"], axis=0
    ).reset_index(drop=True)

    # Initialize outstanding bills / bill share as of model jump-off point
    total_os_bills = outstanding_bills["Amt Outstanding"].sum()
    temp_bills_share = total_os_bills / (
        joined_os["Amt Outstanding"].sum() + total_os_bills
    )

    # Initialize empty arrays to store model outputs
    net_issuance_summary = []
    soma_runoff_summary = []
    final_stats_incl_soma_incl_bills = []
    final_buckets_incl_soma_incl_bills = []
    final_stats_incl_soma_ex_bills = []
    final_buckets_incl_soma_ex_bills = []
    final_stats_ex_soma_incl_bills = []
    final_buckets_ex_soma_incl_bills = []
    final_stats_ex_soma_ex_bills = []
    final_buckets_ex_soma_ex_bills = []

    # Monthly simulation begins here
    for period in range(len(gross_iss["Date"])):
        period_date = gross_iss["Date"][period]

        # Filter full universe down to the currently outstanding stock (+25 to capture all 4 potential settlement dates)
        temp_os = complete_universe.copy()
        temp_os = temp_os[
            (
                temp_os["Issue Date"]
                <= pd.to_datetime(period_date + dt.timedelta(days=25))
            )
            & (temp_os["Maturity Date"] >= period_date)
        ]

        # Adjust TIPS and FRNs based on inflation and forward rates respectively
        temp_os["Amt Outstanding"] = np.where(
            temp_os["Category"] == "TIPS",
            np.maximum(
                temp_os["Issued"],
                temp_os["Amt Outstanding"] * monthly_infl_deltas[period],
            ),
            temp_os["Amt Outstanding"],
        )
        temp_os["SOMA Par Value"] = np.where(
            temp_os["Category"] == "TIPS",
            np.maximum(
                temp_os["SOMA Par Value"],
                temp_os["SOMA Par Value"] * monthly_infl_deltas[period],
            ),
            temp_os["SOMA Par Value"],
        )
        temp_os["Interest Rate"] = np.where(
            temp_os["Category"] == "FRNs", frn_rates[period], temp_os["Interest Rate"]
        )

        # Store a DataFrame of the securities maturing in this period
        temp_t_mat_df = temp_os[
            temp_os["Maturity Date"]
            <= pd.to_datetime(period_date + dt.timedelta(days=25))
        ]
        temp_t_mat = temp_t_mat_df["Amt Outstanding"].sum()

        # Compute SOMA rollovers in two groups, mid-month rollovers and end-of-month rollovers
        temp_mid_sr_date = temp_t_mat_df.groupby("Maturity Date", as_index=False).sum()[
            "Maturity Date"
        ][0]
        temp_mid_soma_rollovers = temp_t_mat_df.groupby(
            "Maturity Date", as_index=False
        ).sum()["SOMA Par Value"][0]
        temp_eom_sr_date = temp_t_mat_df.groupby("Maturity Date", as_index=False).sum()[
            "Maturity Date"
        ][1]
        temp_eom_soma_rollovers = temp_t_mat_df.groupby(
            "Maturity Date", as_index=False
        ).sum()["SOMA Par Value"][1]

        # Total nonbill rollovers this month
        temp_soma_nonbill_rollovers = temp_mid_soma_rollovers + temp_eom_soma_rollovers

        # After calculating rollovers, remove securities maturing in this period from the outstanding stock
        temp_os = temp_os[
            temp_os["Maturity Date"]
            >= pd.to_datetime(period_date + dt.timedelta(days=25))
        ]

        # Read in ssumed SOMA runoff cap from user inputs
        temp_soma_runoff_cap = soma_runoff_caps[
            soma_runoff_caps["Month"] == period_date
        ].iat[0, 1]
        temp_bill_runoff, temp_nonbill_runoff, temp_total_soma_runoff = 0, 0, 0

        # Distribute SOMA runoff cap proportionally between mid-month and EOM rollovers based on their respective sizes
        if temp_soma_runoff_cap != 0:
            # Bill runoff used to "fill in" any remaining runoff that is not alrady met by coupon maturities
            temp_nonbill_runoff = min(temp_soma_nonbill_rollovers, temp_soma_runoff_cap)
            temp_bill_runoff = min(
                temp_soma_runoff_cap - temp_nonbill_runoff,
                outstanding_bills["SOMA Par Value"].sum(),
            )
            temp_total_soma_runoff = temp_nonbill_runoff + temp_bill_runoff

            if temp_soma_nonbill_rollovers != 0:
                temp_mid_soma_rollovers = temp_mid_soma_rollovers - (
                    temp_nonbill_runoff
                    * (temp_mid_soma_rollovers / temp_soma_nonbill_rollovers)
                )
                temp_eom_soma_rollovers = temp_eom_soma_rollovers - (
                    temp_nonbill_runoff
                    * (temp_eom_soma_rollovers / temp_soma_nonbill_rollovers)
                )
                temp_soma_nonbill_rollovers = (
                    temp_soma_nonbill_rollovers - temp_nonbill_runoff
                )

        # SOMA Runoff summary statistics
        temp_period_date = (
            calendar.month_abbr[period_date.month] + " " + str(period_date.year)
        )
        temp_soma_runoff_dict = {
            "Period": temp_period_date,
            "SOMA Coupon Runoff": temp_nonbill_runoff,
            "SOMA Bill Runoff": temp_bill_runoff,
            "SOMA Total Runoff": temp_total_soma_runoff,
        }
        soma_runoff_summary.append(temp_soma_runoff_dict)

        # Calculate issuance add-ons from mid-month SOMA rollovers
        temp_mid_mask = temp_os["Issue Date"] == temp_mid_sr_date
        temp_mid_soma_add_ons = temp_os["Issued"].multiply(temp_mid_mask, axis=0)
        temp_mid_soma_add_ons = temp_mid_soma_add_ons / temp_mid_soma_add_ons.sum()
        temp_mid_soma_add_ons = temp_mid_soma_add_ons * temp_mid_soma_rollovers

        # Calculate issuance add-ons from EOM SOMA rollovers
        temp_eom_mask = temp_os["Issue Date"] == temp_eom_sr_date
        temp_eom_soma_add_ons = temp_os["Issued"].multiply(temp_eom_mask, axis=0)
        temp_eom_soma_add_ons = temp_eom_soma_add_ons / temp_eom_soma_add_ons.sum()
        temp_eom_soma_add_ons = temp_eom_soma_add_ons * temp_eom_soma_rollovers

        # Add mid-month and EOM SOMA add-ons to this period's issuance
        temp_os["Issued"] = (
            temp_os["Issued"] + temp_mid_soma_add_ons + temp_eom_soma_add_ons
        )
        temp_os["Amt Outstanding"] = (
            temp_os["Amt Outstanding"] + temp_mid_soma_add_ons + temp_eom_soma_add_ons
        )
        temp_os["SOMA Par Value"] = (
            temp_os["SOMA Par Value"] + temp_mid_soma_add_ons + temp_eom_soma_add_ons
        )
        temp_os["SOMA Percent Outstanding"] = (
            temp_os["SOMA Par Value"] / temp_os["Amt Outstanding"]
        )
        temp_os["Amt Ex-SOMA"] = temp_os["Amt Outstanding"] - temp_os["SOMA Par Value"]

        # Compute net bill issuance, then update total outstanding bills and bills share
        temp_monthly_issuance_sum = gross_iss.iloc[period, 1:-1].sum()
        outstanding_bills["Amt Outstanding"] = outstanding_bills[
            "Amt Outstanding"
        ].sum() + (
            funding_need_monthly.iat[period, 0]
            - temp_monthly_issuance_sum
            - temp_soma_nonbill_rollovers
            + temp_t_mat
        )
        outstanding_bills["Issued"] = outstanding_bills["Amt Outstanding"]
        temp_bills_share = outstanding_bills["Amt Outstanding"].sum() / (
            temp_os["Amt Outstanding"].sum()
            + outstanding_bills["Amt Outstanding"].sum()
        )

        # Assumed QE purchases from user input
        temp_qe_total_purchases = qe_purchases[
            qe_purchases["Month"] == period_date
        ].iat[0, 1]
        temp_qe_nonbill_purchases = 0.00
        temp_qe_bill_purchases = 0.00

        # Adjustment factor to set QE bill purchases (as a % of bill share)
        qe_bill_purchase_adj = 0.0

        # Apply QE purchases across the currently outstanding outstanding debt stock
        if temp_qe_total_purchases != 0:
            temp_qe_bill_purchases = temp_qe_total_purchases * (
                temp_bills_share * qe_bill_purchase_adj
            )

            # Limit QE purchases of new securities to keep SOMA holdings of any individual security below 70%
            temp_qe_eligible_mask = temp_os["SOMA Percent Outstanding"] < 0.699
            temp_dist_for_qe = temp_os["Amt Outstanding"].multiply(
                temp_qe_eligible_mask, axis=0
            )
            temp_dist_for_qe = temp_dist_for_qe / temp_dist_for_qe.sum()
            temp_qe_nonbill_purchases = (
                temp_qe_total_purchases - temp_qe_bill_purchases
            ) * temp_dist_for_qe
            temp_os["SOMA Par Value"] = (
                temp_os["SOMA Par Value"] + temp_qe_nonbill_purchases
            )

        # Update SOMA nobill holdings post-QE
        temp_os["SOMA Percent Outstanding"] = (
            temp_os["SOMA Par Value"] / temp_os["Amt Outstanding"]
        )
        temp_os["Amt Ex-SOMA"] = temp_os["Amt Outstanding"] - temp_os["SOMA Par Value"]

        # Update SOMA bill holdings post-QE
        outstanding_bills["SOMA Par Value"] = (
            outstanding_bills["SOMA Par Value"].sum()
            + temp_qe_bill_purchases
            - temp_bill_runoff
        )
        outstanding_bills["SOMA Percent Outstanding"] = (
            outstanding_bills["SOMA Par Value"] / outstanding_bills["Amt Outstanding"]
        )
        outstanding_bills["Amt Ex-SOMA"] = (
            outstanding_bills["Amt Outstanding"] - outstanding_bills["SOMA Par Value"]
        )

        # Update full universe DataFrame (over the entire model horizon) with the changes from this period
        for index, r in temp_os.iterrows():
            row_val = temp_os.index.get_loc(index)
            complete_universe.iat[index, 2] = temp_os.iat[row_val, 2]
            complete_universe.iat[index, 5] = temp_os.iat[row_val, 5]
            complete_universe.iat[index, 6] = temp_os.iat[row_val, 6]
            complete_universe.iat[index, 7] = temp_os.iat[row_val, 7]
            complete_universe.iat[index, 8] = temp_os.iat[row_val, 8]
            complete_universe.iat[index, 9] = temp_os.iat[row_val, 9]

        # Compute final summary stats and supply buckets (incl/excl bills, incl/excl SOMA)
        (
            temp_stats_incl_soma_incl_bills,
            temp_buckets_incl_soma_incl_bills,
            temp_stats_ex_soma_incl_bills,
            temp_buckets_ex_soma_incl_bills,
        ) = gen_characteristics(
            temp_os, period_date, outstanding_bills, frn_rates[period]
        )
        (
            temp_stats_incl_soma_ex_bills,
            temp_buckets_incl_soma_ex_bills,
            temp_stats_ex_soma_ex_bills,
            temp_buckets_ex_soma_ex_bills,
        ) = gen_characteristics(
            temp_os,
            period_date,
            outstanding_bills,
            frn_rates[period],
            include_bills=False,
        )

        # Append statistics for the current period to the result arrays
        final_stats_incl_soma_incl_bills.append(temp_stats_incl_soma_incl_bills)
        final_buckets_incl_soma_incl_bills.append(temp_buckets_incl_soma_incl_bills)
        final_stats_incl_soma_ex_bills.append(temp_stats_incl_soma_ex_bills)
        final_buckets_incl_soma_ex_bills.append(temp_buckets_incl_soma_ex_bills)
        final_stats_ex_soma_incl_bills.append(temp_stats_ex_soma_incl_bills)
        final_buckets_ex_soma_incl_bills.append(temp_buckets_ex_soma_incl_bills)
        final_stats_ex_soma_ex_bills.append(temp_stats_ex_soma_ex_bills)
        final_buckets_ex_soma_ex_bills.append(temp_buckets_ex_soma_ex_bills)

        # Net Issuance Summary
        temp_net_iss_row = {
            "Period": calendar.month_abbr[period_date.month]
            + " "
            + str(period_date.year),
            "Net Borrowing Need": funding_need_monthly.iat[period, 0],
            "Net Coupon Issuance": temp_monthly_issuance_sum
            - temp_t_mat
            + temp_soma_nonbill_rollovers,
            "Net Bill Issuance": (
                funding_need_monthly.iat[period, 0]
                - temp_monthly_issuance_sum
                - temp_soma_nonbill_rollovers
                + temp_t_mat
            ),
            "Total Outstanding Bills": outstanding_bills["Amt Outstanding"].sum(),
        }
        net_issuance_summary.append(temp_net_iss_row)

    # Convert final result arrays to DataFrames
    final_net_issuance_summary_df = pd.DataFrame(net_issuance_summary).round(1)
    final_soma_runoff_summary_df = pd.DataFrame(soma_runoff_summary).round(5)

    final_stats_incl_soma_incl_bills_df = pd.DataFrame(
        final_stats_incl_soma_incl_bills
    ).round(5)
    final_buckets_incl_soma_incl_bills_df = pd.DataFrame(
        final_buckets_incl_soma_incl_bills
    ).round(5)
    final_stats_incl_soma_ex_bills_df = pd.DataFrame(
        final_stats_incl_soma_ex_bills
    ).round(5)
    final_stats_incl_soma_ex_bills_df.drop(
        "Bill Share", axis=1, inplace=True
    )  # Drop bill share from nonbill stats
    final_buckets_incl_soma_ex_bills_df = pd.DataFrame(
        final_buckets_incl_soma_ex_bills
    ).round(5)

    final_stats_ex_soma_incl_bills_df = pd.DataFrame(
        final_stats_ex_soma_incl_bills
    ).round(5)
    final_buckets_ex_soma_incl_bills_df = pd.DataFrame(
        final_buckets_ex_soma_incl_bills
    ).round(5)
    final_stats_ex_soma_ex_bills_df = pd.DataFrame(final_stats_ex_soma_ex_bills).round(
        5
    )
    final_stats_ex_soma_ex_bills_df.drop(
        "Bill Share", axis=1, inplace=True
    )  # Drop bill share from nonbill stats
    final_buckets_ex_soma_ex_bills_df = pd.DataFrame(
        final_buckets_ex_soma_ex_bills
    ).round(5)

    # Add SOMA amt outstanding to summary statistics DataFrames
    final_stats_incl_soma_incl_bills_df["SOMA Amt Outstanding"] = (
        final_stats_incl_soma_incl_bills_df["Amt Outstanding"]
        - final_stats_ex_soma_incl_bills_df["Amt Outstanding"]
    )
    final_stats_incl_soma_ex_bills_df["SOMA Amt Outstanding"] = (
        final_stats_incl_soma_ex_bills_df["Amt Outstanding"]
        - final_stats_ex_soma_ex_bills_df["Amt Outstanding"]
    )

    # Plot bill share
    plt.plot(
        gross_iss["Date"],
        final_stats_incl_soma_incl_bills_df["Bill Share"],
        label="Bill Share Incl SOMA",
    )
    x1, x2, y1, y2 = plt.axis()
    plt.axis((x1, x2, 0.0, 0.3))
    plt.legend()
    plt.grid(axis="y")
    plt.title("Bill Share")
    plt.show()

    # Write outputs to the specified Excel file
    output_file_path = "output_files/Output_2022_03_31.xlsx"
    with pd.ExcelWriter(output_file_path, engine="openpyxl", mode="a") as writer:
        final_net_issuance_summary_df.to_excel(
            writer, sheet_name="Net Issuance Summary"
        )

        final_soma_runoff_summary_df.to_excel(writer, sheet_name="SOMA Runoff Summary")

        final_stats_incl_soma_incl_bills_df.to_excel(writer, sheet_name="Summary Stats")
        final_stats_incl_soma_ex_bills_df.to_excel(
            writer, sheet_name="Nonbill Summary Stats"
        )
        final_buckets_incl_soma_incl_bills_df.to_excel(
            writer, sheet_name="Supply Buckets"
        )
        final_buckets_incl_soma_ex_bills_df.to_excel(
            writer, sheet_name="Nonbill Supply Buckets"
        )

        # Outputs ex-SOMA not written to final spreadsheet; can be enabled by uncommenting the lines below
        # final_stats_ex_soma_incl_bills_df.to_excel(writer, sheet_name='Summary Stats Ex SOMA')
        # final_stats_ex_soma_ex_bills_df.to_excel(writer, sheet_name='Nonbill Summary Stats Ex SOMA')
        # final_buckets_ex_soma_incl_bills_df.to_excel(writer, sheet_name='Supply Buckets w Bills Ex SOMA')
        # final_buckets_ex_soma_ex_bills_df.to_excel(writer, sheet_name='Nonbill Supply Buckets Ex SOMA')
