"""Main entrypoint for the Excel-based risk attribution workflow."""

import argparse
from typing import Any, Dict, Optional

import pandas as pd

from 产品归因 import run_product_attribution
from 四因子计算 import calculate_factor_returns
from 数据读取 import DEFAULT_DURATION_FILE, DEFAULT_MARKET_FILE, DEFAULT_PRODUCT_FILE


def print_outputs(
    factor_returns: pd.DataFrame,
    factor_correlation: pd.DataFrame,
    attribution: Dict[str, Any],
) -> None:
    """Print the main output tables from both stages."""
    print("\n四因子收益率序列")
    print("=" * 80)
    print(factor_returns)
    print("\n因子相关性矩阵")
    print("=" * 80)
    print(factor_correlation)
    print("\n回归统计表")
    print("=" * 80)
    print(attribution["regression_stats"])
    print("\n方差分析表")
    print("=" * 80)
    print(attribution["anova_table"])
    print("\n回归系数表")
    print("=" * 80)
    print(attribution["coef_table"])
    print("\n协方差矩阵")
    print("=" * 80)
    print(attribution["cov_matrix"])
    print("\n方差分解表")
    print("=" * 80)
    print(attribution["variance_decomp"])
    print("\n因子方差贡献分析表")
    print("=" * 80)
    print(attribution["factor_contrib_df"])


def build_arg_parser() -> argparse.ArgumentParser:
    """Build command-line arguments for the main script."""
    parser = argparse.ArgumentParser(description="Excel-based risk attribution script.")
    parser.add_argument("--w", type=int, required=True, help="Window length in weeks.")
    parser.add_argument("--end-date", default=None, help="Window end date, e.g. 2026-06-19.")
    parser.add_argument("--output-excel", default=None, help="Excel output path.")
    parser.add_argument("--output-chart", default=None, help="Chart image output path.")
    parser.add_argument("--market-file", default=str(DEFAULT_MARKET_FILE), help="Market data Excel path.")
    parser.add_argument("--duration-file", default=str(DEFAULT_DURATION_FILE), help="Duration Excel path.")
    parser.add_argument("--product-file", default=str(DEFAULT_PRODUCT_FILE), help="Product value Excel path.")
    return parser


def main(argv: Optional[list] = None) -> Dict[str, Any]:
    """Run factor calculation first, then product attribution."""
    args = build_arg_parser().parse_args(argv)
    factor_returns, factor_correlation = calculate_factor_returns(
        w=args.w,
        end_date=args.end_date,
        market_file=args.market_file,
        duration_file=args.duration_file,
    )
    attribution = run_product_attribution(
        w=args.w,
        factor_returns=factor_returns,
        product_file=args.product_file,
        end_date=args.end_date,
        output_excel=args.output_excel,
        output_chart=args.output_chart,
    )
    print_outputs(factor_returns, factor_correlation, attribution)
    if args.output_excel:
        print("\nExcel output saved to: {}".format(args.output_excel))
    if args.output_chart:
        print("Chart output saved to: {}".format(args.output_chart))
    return {
        "factor_returns": factor_returns,
        "factor_correlation": factor_correlation,
        "attribution": attribution,
    }


if __name__ == "__main__":
    main()
