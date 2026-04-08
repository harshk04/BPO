from __future__ import annotations

from dataclasses import dataclass, asdict
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Dict, Any

TWOPLACES = Decimal('0.01')
ONE = Decimal('1')
HUNDRED = Decimal('100')


def q2(value: Decimal) -> Decimal:
    return value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def to_decimal(text: str) -> Decimal:
    try:
        return Decimal(text.strip())
    except (InvalidOperation, AttributeError):
        raise ValueError(f"Invalid numeric input: {text!r}")


def normalize_percent(value: Decimal) -> Decimal:
    """
    Accepts either 14 or 0.14 and returns 0.14.
    """
    if abs(value) > ONE:
        return value / HUNDRED
    return value


@dataclass
class LoanOutputs:
    reduced_value: Decimal
    purchase_value_reduction_amount: Decimal
    remark: str
    down_payment_value: Decimal
    loan_amount: Decimal
    annual_principal: Decimal
    monthly_principal: Decimal
    principal_after_monthly_reduction: Decimal
    remark_2: str
    per_annum_interest: Decimal
    total_interest_for_loan_period: Decimal
    total_interest_after_reduction: Decimal
    insurance_amount_per_month: Decimal



def calculate(
    purchase_value: Decimal,
    reduction_percent: Decimal,
    down_payment_percent: Decimal,
    loan_period_years: Decimal,
    monthly_principal_reduction_percent: Decimal,
    interest_rate_percent: Decimal,
    total_interest_reduction_percent: Decimal,
    insurance_rate_percent: Decimal,
) -> LoanOutputs:
    if loan_period_years <= 0:
        raise ValueError("Loan period must be greater than 0.")

    # Percent inputs can be typed as 14 or 0.14.
    reduction_percent = normalize_percent(reduction_percent)
    down_payment_percent = normalize_percent(down_payment_percent)
    monthly_principal_reduction_percent = normalize_percent(monthly_principal_reduction_percent)
    interest_rate_percent = normalize_percent(interest_rate_percent)
    total_interest_reduction_percent = normalize_percent(total_interest_reduction_percent)
    insurance_rate_percent = normalize_percent(insurance_rate_percent)

    # These formulas intentionally follow the sheet step-by-step with ROUND at each stage.
    reduced_value = q2(purchase_value * (ONE - reduction_percent))

    # Formula reference:
    # Purchase Value.1 = ROUND(Purchase Value - Reduced Value, 2)
    # (Excel-style sequential rounding can differ by 0.01 from direct multiplication.)
    purchase_value_reduction_amount = q2(purchase_value - reduced_value)
    remark = f"{purchase_value_reduction_amount:.2f} AND {(down_payment_percent * HUNDRED):.2f}%"

    # Formula reference:
    # Loan Amount = ROUND(Purchase Value.1 × Down Payment %, 2)
    # Down Payment Value = ROUND(Purchase Value.1 × (1 − Down Payment %), 2)
    loan_amount = q2(purchase_value_reduction_amount * down_payment_percent)
    down_payment_value = q2(purchase_value_reduction_amount * (ONE - down_payment_percent))

    annual_principal = q2(loan_amount / loan_period_years)
    monthly_principal = q2(annual_principal / Decimal('12'))
    principal_after_monthly_reduction = q2(
        monthly_principal * (ONE - monthly_principal_reduction_percent)
    )

    remark_2 = f"{loan_amount:.2f} AND {principal_after_monthly_reduction:.2f}"

    per_annum_interest = q2(loan_amount * interest_rate_percent)
    total_interest_for_loan_period = q2(per_annum_interest * loan_period_years)
    total_interest_after_reduction = q2(
        total_interest_for_loan_period * (ONE - total_interest_reduction_percent)
    )
    insurance_amount_per_month = q2((loan_amount * insurance_rate_percent) / Decimal('12'))

    return LoanOutputs(
        reduced_value=reduced_value,
        purchase_value_reduction_amount=purchase_value_reduction_amount,
        remark=remark,
        down_payment_value=q2(down_payment_value),
        loan_amount=loan_amount,
        annual_principal=annual_principal,
        monthly_principal=monthly_principal,
        principal_after_monthly_reduction=principal_after_monthly_reduction,
        remark_2=remark_2,
        per_annum_interest=per_annum_interest,
        total_interest_for_loan_period=total_interest_for_loan_period,
        total_interest_after_reduction=total_interest_after_reduction,
        insurance_amount_per_month=insurance_amount_per_month,
    )



def print_outputs(result: LoanOutputs) -> None:
    print("\nOutputs")
    print("-" * 60)
    print(f"Reduced value: {result.reduced_value:.2f}")
    print(f"Purchase Value: {result.purchase_value_reduction_amount:.2f}")
    print(f"Remark: {result.remark}")
    print(f"Down payment value: {result.down_payment_value:.2f}")
    print(f"Loan amount: {result.loan_amount:.2f}")
    print(f"Loan amount / Loan period = Annual Principal: {result.annual_principal:.2f}")
    print(f"Annual principal / 12 months = Monthly Principal: {result.monthly_principal:.2f}")
    print(
        "Monthly Principal - Reduction percentage = Principal: "
        f"{result.principal_after_monthly_reduction:.2f}"
    )
    print(f"Remark 2: {result.remark_2}")
    print(f"Per annum interest: {result.per_annum_interest:.2f}")
    print(f"Total Interest for loan period: {result.total_interest_for_loan_period:.2f}")
    print(f"Total Interest after reduction: {result.total_interest_after_reduction:.2f}")
    print(f"Insurance Amount Per Month: {result.insurance_amount_per_month:.2f}")
    print("-" * 60)



def prompt_decimal(label: str) -> Decimal:
    while True:
        raw = input(f"{label}: ").strip()
        try:
            return to_decimal(raw)
        except ValueError as exc:
            print(exc)



def main() -> None:
    print("Excel formula based loan calculator")
    print("Enter percentages either like 14 or 0.14")

    while True:
        try:
            purchase_value = prompt_decimal("Purchase Value")
            reduction_percent = prompt_decimal("Reduction % value")
            down_payment_percent = prompt_decimal("Down payment In %")
            loan_period_years = prompt_decimal("Loan period In Year")
            monthly_principal_reduction_percent = prompt_decimal("Monthly Principal Reduction%")
            interest_rate_percent = prompt_decimal("Int rate %")
            total_interest_reduction_percent = prompt_decimal("Total Interest Reduction")
            insurance_rate_percent = prompt_decimal("Insurance Rate")

            result = calculate(
                purchase_value=purchase_value,
                reduction_percent=reduction_percent,
                down_payment_percent=down_payment_percent,
                loan_period_years=loan_period_years,
                monthly_principal_reduction_percent=monthly_principal_reduction_percent,
                interest_rate_percent=interest_rate_percent,
                total_interest_reduction_percent=total_interest_reduction_percent,
                insurance_rate_percent=insurance_rate_percent,
            )
            print_outputs(result)
        except KeyboardInterrupt:
            print("\nStopped.")
            return
        except Exception as exc:
            print(f"Error: {exc}")

        again = input("Do you want to calculate again? (y/n): ").strip().lower()
        if again not in {"y", "yes"}:
            print("Done.")
            break


if __name__ == "__main__":
    main()
