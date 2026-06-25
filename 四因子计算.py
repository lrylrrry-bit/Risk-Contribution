"""Four-factor weekly return and correlation calculation from Excel data."""

from typing import Any, Optional, Tuple

import numpy as np
import pandas as pd

from 数据读取 import read_duration_data, read_market_price, read_risk_free_yield


FACTOR_COLUMNS = ["Level", "Slope", "Credit", "Equity"]
DEFAULT_CODES = {
    "level": "CBA00602.CB",
    "slope_602": "CBA00602.CB",
    "slope_722": "CBA00722.CB",
    "credit_aaa": "CBA04202.CB",
    "credit_policy": "CBA02502.CB",
    "equity": "000300.SH",
}


def calculate_factor_returns(
    w: int,
    end_date: Optional[Any] = None,
    market_file: Optional[Any] = None,
    duration_file: Optional[Any] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Calculate weekly factor returns and factor correlation matrix."""
    validate_w(w)
    anchor = latest_friday(end_date)

    level_df = read_market_price(DEFAULT_CODES["level"], market_file=market_file)
    level_weekly = level_df.resample("W-FRI").last()
    level = np.log(level_weekly["close"] / level_weekly["close"].shift(1)).rename("Level")

    slope_602_df = read_market_price(DEFAULT_CODES["slope_602"], market_file=market_file)
    slope_722_df = read_market_price(DEFAULT_CODES["slope_722"], market_file=market_file)
    duration_df = read_duration_data(duration_file)
    slope = calculate_slope_factor(slope_602_df, slope_722_df, duration_df)

    credit_aaa_df = read_market_price(DEFAULT_CODES["credit_aaa"], market_file=market_file)
    credit_policy_df = read_market_price(DEFAULT_CODES["credit_policy"], market_file=market_file)
    credit_aaa_weekly = credit_aaa_df.resample("W-FRI").last()
    credit_policy_weekly = credit_policy_df.resample("W-FRI").last()
    credit = (
        np.log(credit_aaa_weekly["close"] / credit_aaa_weekly["close"].shift(1))
        - np.log(credit_policy_weekly["close"] / credit_policy_weekly["close"].shift(1))
    ).rename("Credit")

    equity_df = read_market_price(DEFAULT_CODES["equity"], market_file=market_file)
    risk_free_df = read_risk_free_yield(market_file=market_file)
    equity_weekly = equity_df.resample("W-FRI").last()
    risk_free_weekly = risk_free_df.resample("W-FRI").last()
    equity_joined = pd.concat([equity_weekly, risk_free_weekly], axis=1).dropna()
    equity = (
        np.log(equity_joined["close"] / equity_joined["close"].shift(1))
        - equity_joined["yield_10y"] / 52 / 100
    ).rename("Equity")

    factor_returns = pd.concat([level, slope, credit, equity], axis=1).dropna()
    factor_returns = trim_weekly_window(factor_returns, w, anchor, "factor returns")
    factor_returns = factor_returns[FACTOR_COLUMNS]
    return factor_returns, factor_returns.corr()


def calculate_slope_factor(
    slope_602_df: pd.DataFrame,
    slope_722_df: pd.DataFrame,
    duration_df: pd.DataFrame,
) -> pd.Series:
    """Calculate the Slope factor using the original duration-neutral formula."""
    cba_602_weekly = slope_602_df.resample("W-FRI").last()["close"]
    cba_722_weekly = slope_722_df.resample("W-FRI").last()["close"]
    df_full = pd.DataFrame(
        {
            "CBA00602_close": cba_602_weekly,
            "CBA00722_close": cba_722_weekly,
        }
    ).dropna()
    duration_weekly = duration_df.resample("W-FRI").last()
    df_full = df_full.join(duration_weekly, how="inner").dropna()
    if len(df_full) < 2:
        raise ValueError("Not enough data to calculate Slope factor.")

    df_full["组合价值"] = 1.0
    for i in range(1, len(df_full)):
        prev_date = df_full.index[i - 1]
        curr_date = df_full.index[i]

        prev_dur_602 = df_full.loc[prev_date, "CBA00602_duration"]
        prev_dur_722 = df_full.loc[prev_date, "CBA00722_duration"]
        dur_diff = prev_dur_602 - prev_dur_722
        if abs(dur_diff) > 1e-10:
            weight_602 = -prev_dur_722 / dur_diff
            weight_722 = prev_dur_602 / dur_diff
        else:
            weight_602 = 0.0
            weight_722 = 0.0

        prev_portfolio_value = df_full.loc[prev_date, "组合价值"]
        position_602 = (
            weight_602 * prev_portfolio_value / df_full.loc[prev_date, "CBA00602_close"]
        )
        position_722 = (
            weight_722 * prev_portfolio_value / df_full.loc[prev_date, "CBA00722_close"]
        )
        curr_portfolio_value = (
            position_602 * df_full.loc[curr_date, "CBA00602_close"]
            + position_722 * df_full.loc[curr_date, "CBA00722_close"]
        )
        df_full.loc[curr_date, "组合价值"] = curr_portfolio_value

    return np.log(df_full["组合价值"] / df_full["组合价值"].shift(1)).rename("Slope")


def latest_friday(date_value: Optional[Any]) -> pd.Timestamp:
    """Return the latest Friday on or before date_value."""
    if date_value is None:
        date_value = pd.Timestamp.today().normalize()
    ts = pd.Timestamp(date_value).normalize()
    return ts - pd.Timedelta(days=(ts.weekday() - 4) % 7)


def trim_weekly_window(
    df: pd.DataFrame,
    w: int,
    anchor: Any,
    label: str,
) -> pd.DataFrame:
    """Keep the latest W rows on or before the anchor date."""
    anchor_ts = pd.Timestamp(anchor).normalize()
    data = df.copy()
    data.index = pd.to_datetime(data.index)
    data = data.sort_index()
    data = data[data.index <= anchor_ts]
    data = data.tail(w)
    if len(data) < w:
        raise ValueError(
            "Not enough {} before {}: need {}, got {}.".format(
                label, anchor_ts.date(), w, len(data)
            )
        )
    return data


def validate_w(w: int) -> None:
    """Validate the weekly window length."""
    if not isinstance(w, int):
        raise TypeError("W must be an integer number of weeks.")
    if w <= 0:
        raise ValueError("W must be positive.")
