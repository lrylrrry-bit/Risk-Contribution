"""Product return regression and variance decomposition from Excel data."""

from typing import Any, Dict, Optional

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats

from 四因子计算 import FACTOR_COLUMNS, latest_friday, trim_weekly_window, validate_w
from 数据读取 import read_product_returns


LOWER_FACTOR_COLUMNS = ["level", "slope", "credit", "equity"]


def run_product_attribution(
    w: int,
    factor_returns: pd.DataFrame,
    product_file: Optional[Any] = None,
    end_date: Optional[Any] = None,
    output_excel: Optional[str] = None,
    output_chart: Optional[str] = None,
) -> Dict[str, Any]:
    """Run product return regression and variance contribution analysis."""
    validate_w(w)
    if w <= len(FACTOR_COLUMNS) + 1:
        raise ValueError("W must be greater than 5 for four-factor regression.")

    product_returns = read_product_returns(product_file)
    factor_data = normalize_factor_returns(factor_returns)
    anchor = latest_friday(end_date) if end_date is not None else factor_data.index.max()

    regression_data = pd.concat(
        [product_returns.rename("期间收益率"), factor_data[FACTOR_COLUMNS]], axis=1
    ).dropna()
    regression_data = trim_weekly_window(regression_data, w, anchor, "regression data")

    y = regression_data["期间收益率"]
    x = regression_data[FACTOR_COLUMNS]
    model = sm.OLS(y, sm.add_constant(x)).fit()

    n = len(y)
    k = len(x.columns)
    multiple_r = np.sqrt(model.rsquared)
    standard_error = np.sqrt(model.mse_resid)

    ssr = model.ess
    sse = model.ssr
    sst = ssr + sse
    msr = ssr / k
    mse = sse / (n - k - 1)
    f_statistic = msr / mse
    f_p_value = 1 - stats.f.cdf(f_statistic, k, n - k - 1)

    regression_stats = pd.DataFrame(
        {
            "统计量": [
                "Multiple R",
                "R Square",
                "Adjusted R Square",
                "标准误差",
                "观测值",
            ],
            "数值": [
                multiple_r,
                model.rsquared,
                model.rsquared_adj,
                standard_error,
                n,
            ],
        }
    )
    anova_table = pd.DataFrame(
        {
            "项目": ["回归分析", "残差", "总计"],
            "df": [k, n - k - 1, n - 1],
            "SS": [ssr, sse, sst],
            "MS": [msr, mse, np.nan],
            "F": [f_statistic, np.nan, np.nan],
            "Significance F": [f_p_value, np.nan, np.nan],
        }
    )

    conf_int = model.conf_int()
    coef_table = pd.DataFrame(
        {
            "Coefficients": model.params,
            "标准误差": model.bse,
            "t Stat": model.tvalues,
            "P-value": model.pvalues,
            "Lower 95%": conf_int[0],
            "Upper 95%": conf_int[1],
            "下限 95.0%": conf_int[0],
            "上限 95.0%": conf_int[1],
        }
    )
    coef_table.index = ["Intercept"] + LOWER_FACTOR_COLUMNS

    cov_data = regression_data[["期间收益率"] + FACTOR_COLUMNS].copy()
    cov_data.columns = ["期间收益率"] + LOWER_FACTOR_COLUMNS
    cov_matrix = cov_data.cov()

    beta = model.params[1:].copy()
    beta.index = LOWER_FACTOR_COLUMNS
    factor_cov = cov_matrix.loc[LOWER_FACTOR_COLUMNS, LOWER_FACTOR_COLUMNS]
    portfolio_total_var = cov_matrix.loc["期间收益率", "期间收益率"]
    explained_var = portfolio_total_var * model.rsquared
    unexplained_var = portfolio_total_var * (1 - model.rsquared)

    factor_contributions = {}
    for factor_i in LOWER_FACTOR_COLUMNS:
        contribution = 0.0
        for factor_j in LOWER_FACTOR_COLUMNS:
            contribution += beta[factor_j] * factor_cov.loc[factor_i, factor_j]
        factor_contributions[factor_i] = contribution * beta[factor_i]

    total_contribution = sum(factor_contributions.values())
    factor_contrib_df = pd.DataFrame(
        {
            "因子": list(factor_contributions.keys()),
            "方差贡献": list(factor_contributions.values()),
            "方差占比(%)": [
                safe_percentage(value, total_contribution)
                for value in factor_contributions.values()
            ],
        }
    )
    variance_decomp = pd.DataFrame(
        {
            "因子": [
                "组合总方差",
                "R Square",
                "可被解释的组合方差",
                "未被解释的组合方差",
                "因子贡献合计",
            ],
            "数值": [
                portfolio_total_var,
                model.rsquared,
                explained_var,
                unexplained_var,
                total_contribution,
            ],
        }
    )

    fig = create_factor_contribution_chart(
        factor_contrib_df,
        total_contribution=total_contribution,
        r_squared=model.rsquared,
    )

    result = {
        "product_returns": product_returns,
        "regression_data": regression_data,
        "regression_stats": regression_stats,
        "anova_table": anova_table,
        "coef_table": coef_table,
        "cov_matrix": cov_matrix,
        "variance_decomp": variance_decomp,
        "factor_contrib_df": factor_contrib_df,
        "factor_contrib_fig": fig,
        "regression_model": model,
    }

    if output_excel:
        export_attribution_to_excel(output_excel, factor_returns, result)
    if output_chart:
        fig.savefig(output_chart, dpi=200, bbox_inches="tight")

    return result


def normalize_factor_returns(factor_returns: pd.DataFrame) -> pd.DataFrame:
    """Validate and normalize the factor return DataFrame."""
    if not isinstance(factor_returns, pd.DataFrame):
        raise TypeError("factor_returns must be a pandas DataFrame.")
    missing = [col for col in FACTOR_COLUMNS if col not in factor_returns.columns]
    if missing:
        raise ValueError("factor_returns missing columns: {}".format(", ".join(missing)))
    data = factor_returns[FACTOR_COLUMNS].copy()
    data.index = pd.to_datetime(data.index)
    return data.sort_index()


def create_factor_contribution_chart(
    factor_contrib_df: pd.DataFrame,
    total_contribution: float,
    r_squared: float,
):
    """Create a horizontal bar chart for factor variance contribution."""
    import matplotlib.pyplot as plt

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(10, 6))
    factors = factor_contrib_df["因子"].tolist()
    contributions = factor_contrib_df["方差占比(%)"].tolist()
    bars = ax.barh(factors, contributions, color="#1f77b4", alpha=0.8, height=0.5)

    ax.set_xlabel("方差贡献占比 (%)", fontsize=12)
    ax.set_title("因子方差贡献分析", fontsize=14, fontweight="bold", pad=15)
    for bar in bars:
        width = bar.get_width()
        label_x = width + 0.5 if width >= 0 else width - 0.5
        label_ha = "left" if width >= 0 else "right"
        ax.text(
            label_x,
            bar.get_y() + bar.get_height() / 2,
            "{:.2f}%".format(width),
            va="center",
            ha=label_ha,
            fontsize=11,
            fontweight="bold",
        )

    stats_text = "总解释方差: {:.2e}\nR²: {:.2%}".format(total_contribution, r_squared)
    ax.text(
        0.98,
        0.02,
        stats_text,
        transform=ax.transAxes,
        fontsize=10,
        verticalalignment="bottom",
        horizontalalignment="right",
        bbox=dict(boxstyle="round", facecolor="#f0f0f0", alpha=0.8),
    )
    ax.grid(axis="x", alpha=0.3, linestyle="--")
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return fig


def export_attribution_to_excel(
    output_excel: str,
    factor_returns: pd.DataFrame,
    result: Dict[str, Any],
) -> None:
    """Export factor data and attribution output tables to Excel."""
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        factor_returns.to_excel(writer, sheet_name="四因子收益率", index=True)
        factor_returns.corr().to_excel(writer, sheet_name="因子相关性矩阵", index=True)
        result["product_returns"].to_excel(writer, sheet_name="产品周频收益率", index=True)
        result["regression_stats"].to_excel(writer, sheet_name="回归统计", index=False)
        result["anova_table"].to_excel(writer, sheet_name="方差分析", index=False)
        result["coef_table"].to_excel(writer, sheet_name="回归系数", index=True)
        result["cov_matrix"].to_excel(writer, sheet_name="协方差矩阵", index=True)
        result["variance_decomp"].to_excel(writer, sheet_name="方差分解", index=False)
        result["factor_contrib_df"].to_excel(writer, sheet_name="因子贡献分析", index=False)


def safe_percentage(value: float, total: float) -> float:
    """Calculate percentage safely when the denominator is near zero."""
    if abs(total) <= 1e-15:
        return 0.0
    return value / total * 100
