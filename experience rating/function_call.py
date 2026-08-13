"""Deterministic experience-rating functions for the Reinsurance Rating custom GPT.

This module is intentionally self-contained and uses only the Python standard library.
It does not read Excel files. The GPT must extract, normalize, validate, and document
workbook inputs first, then pass explicit values to these functions and retain exact
workbook/prompt provenance in its JSON call register.

Conventions
-----------
* Rates, trends, shares, and loss ratios are decimals (5% is 0.05).
* Date intervals are half-open: [start, end).
* Triangles are rectangular lists of rows; future/unobserved cells are ``None``.
* Loss development factors are age-to-age factors. CDFs include any tail factor.
* Functions raise ``ValueError`` rather than silently inventing missing assumptions.
* Results are JSON-serializable dictionaries/lists/scalars.

This library supports experience-rating calculations and diagnostics. It deliberately
does not run catastrophe models or full exposure-curve pricing.
Actuarial judgment remains necessary for data basis, selections, and method choice.

Public API quick reference (required inputs -> principal output)
----------------------------------------------------------------
Rate level, on-leveling, and trend:
* ``compound_change_factors``: decimal changes, initial index -> cumulative indices.
* ``rate_level_at_date``: dated decimal changes, as-of date, base index -> rate index.
* ``time_weighted_average_rate_level``: date interval, dated changes, optional daily
  weights -> written/earning-date average rate index.
* ``parallelogram_average_rate_level``: experience interval, dated changes, policy
  term, optional writing weights -> earned average rate index.
* ``parallelogram_on_level_factor``: preceding inputs plus target date -> historical
  average, target index, and OLF.
* ``extension_of_exposures_on_level_factor``: matched historical and rerated policy
  premiums -> totals and OLF.
* ``rate_book_on_level_factor``: premium by rate book, historical indices, target
  index -> on-leveled total and OLF.
* ``annual_rate_change_on_level_factors``: ascending years, annual decimal changes,
  target year, averaging convention -> approximate rate levels and OLF by year.
* ``trend_factor``: from/to dates, annual decimal trend, convention -> scalar factor.
* ``apply_factors``: values and scalar/vector multipliers -> adjusted values.
* ``on_level_and_trend_premium``: earned premium, OLFs, exposure-trend factors ->
  nominal, on-leveled, and trended/on-leveled premium vectors.
* ``trend_losses``: losses, corresponding from-dates, target date, annual trends ->
  trend factors and trended loss vector.
* ``piecewise_trend_factor``: from/to dates and a dated annual-trend schedule ->
  one factor that compounds the applicable rate over each schedule segment.
* ``trend_losses_piecewise``: claim losses, dates, keyed trend schedules, and claim
  schedule keys -> claim-level factors and trended losses.

Triangle preparation, development, and diagnostics:
* ``validate_triangle``: rectangular triangle and cumulative flag -> validation report.
* ``incremental_to_cumulative``: incremental triangle -> cumulative triangle.
* ``cumulative_to_incremental``: cumulative triangle -> incremental triangle.
* ``link_ratio_triangle``: cumulative triangle -> observed age-to-age ratios.
* ``development_factor_diagnostics``: cumulative triangle -> counts, averages,
  medians, and coefficients of variation by development period.
* ``select_age_to_age_factors``: cumulative triangle, selection method and optional
  windows/exclusions/explicit selections -> factors and observations used.
* ``cumulative_development_factors``: age-to-age factors and tail -> CDFs by age.
* ``reported_percentages``: CDFs -> reciprocal reported/paid percentages.
* ``latest_diagonal``: cumulative triangle -> latest age and value by origin row.
* ``fill_lower_triangle_chain_ladder``: cumulative triangle and selected factors ->
  fully projected cumulative triangle (excluding a separate tail beyond its columns).
* ``chain_ladder_ultimate``: cumulative triangle and factor-selection controls ->
  selections, CDFs, completed triangle, ultimate and IBNR by origin and total.
* ``expected_loss_ultimate``: exposure basis and a priori loss ratio -> expected
  ultimate by origin and total.
* ``bornhuetter_ferguson_ultimate``: actual loss, origin CDFs, exposure basis, a
  priori loss ratio -> BF ultimate and IBNR by origin and total.
* ``cape_cod_ultimate``: actual loss, origin CDFs, exposure basis, optional weights ->
  fitted ELR, ultimate and IBNR by origin and total.
* ``benktander_ultimate``: actual loss, origin CDFs, initial expected ultimates,
  iteration count -> iterative ultimate history and final totals.
* ``loglinear_tail_factor``: observed age-to-age factors and extrapolation controls ->
  regression details, projected factors, and deterministic tail.
* ``mack_chain_ladder``: cumulative triangle and tail -> volume-weighted Chain Ladder
  ultimates plus Mack process/parameter standard-error diagnostics (no tail error).
* ``compare_paid_incurred_ultimates``: paid/incurred triangles, explicit blend weight,
  optional factor selections/tails -> both projections and blended ultimates.

Claim, layer, proportional, frequency, and severity:
* ``cap_losses``: losses and cap -> individually capped losses.
* ``apply_per_claim_layer``: claim losses, attachment, limit, share, optional ALAE and
  treatment -> claim-level and total ceded layer loss.
* ``aggregate_occurrence_losses``: claim losses and event IDs -> event totals.
* ``apply_occurrence_layer``: claim losses, event IDs, layer terms -> event-layer losses.
* ``quota_share_ceded``: values, ceded share, optional commission -> ceded/retained
  values and optional commission.
* ``limited_average_severity_empirical``: claims, attachment, limit -> empirical gross
  and layer mean severities and layer total.
* ``empirical_layer_share``: claims, attachment, limit -> gross/layer totals and share.
* ``frequency_severity_summary``: claims, exposure, zero-loss convention -> count,
  frequency, severity, and pure premium.
* ``claim_level_xol_summary``: claim IDs, reported losses, trend/development factors,
  grouping fields, and optional layer terms -> synchronized claim-level and grouped
  nominal/trended reported/ultimate ground-up and layer losses.

Ratios, summaries, credibility, and orchestration:
* ``loss_ratio``: loss and positive earned premium -> scalar loss ratio.
* ``loss_ratio_series``: loss and earned-premium vectors -> period and aggregate LRs.
* ``weighted_average``: values and nonnegative weights -> normalized weighted mean.
* ``trimmed_mean``: values and tail-trim count -> trimmed mean.
* ``credibility_weight``: expected count and full-credibility count -> square-root Z.
* ``credibility_blend``: experience, complement, and Z -> credibility-weighted result.
* ``experience_rating_table``: origin labels, premium/loss vectors, selected adjustment
  factors, methods and method assumptions -> standard rows and method results.
* ``method_suitability_screen``: explicit stability/data-availability indicators,
  expected count, and immature share -> transparent ranked considerations for review.
* ``run_self_tests``: no inputs -> pass/fail report covering every registered function.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from math import exp, isfinite, log, prod, sqrt
from statistics import median
from typing import Any, Iterable, Mapping, Sequence


Number = int | float
Triangle = Sequence[Sequence[Number | None]]


def _as_date(value: date | datetime | str) -> date:
    """Convert an ISO date/datetime/date object to ``date``."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value[:10])
        except ValueError as exc:
            raise ValueError(f"Expected ISO date, got {value!r}") from exc
    raise ValueError(f"Unsupported date value: {value!r}")


def _float(value: Number, name: str) -> float:
    result = float(value)
    if not isfinite(result):
        raise ValueError(f"{name} must be finite")
    return result


def _numbers(values: Sequence[Number], name: str, *, nonnegative: bool = False) -> list[float]:
    result = [_float(v, name) for v in values]
    if nonnegative and any(v < 0 for v in result):
        raise ValueError(f"{name} must be nonnegative")
    return result


def _broadcast(value: Number | Sequence[Number], length: int, name: str) -> list[float]:
    if isinstance(value, (int, float)):
        return [_float(value, name)] * length
    result = _numbers(value, name)
    if len(result) != length:
        raise ValueError(f"{name} length {len(result)} does not match required length {length}")
    return result


def _safe_div(numerator: Number, denominator: Number, name: str) -> float:
    den = _float(denominator, name)
    if den == 0:
        raise ValueError(f"{name} denominator cannot be zero")
    return _float(numerator, name) / den


def _coerce_triangle(triangle: Triangle, *, require_cumulative: bool = False) -> list[list[float | None]]:
    rows = [list(row) for row in triangle]
    if not rows or not rows[0]:
        raise ValueError("triangle cannot be empty")
    width = len(rows[0])
    if any(len(row) != width for row in rows):
        raise ValueError("triangle must be rectangular")
    result: list[list[float | None]] = []
    for r, row in enumerate(rows):
        seen_none = False
        converted: list[float | None] = []
        previous: float | None = None
        for c, value in enumerate(row):
            if value is None:
                seen_none = True
                converted.append(None)
                continue
            if seen_none:
                raise ValueError(f"triangle row {r} has an observed value after a missing cell at column {c}")
            number = _float(value, f"triangle[{r}][{c}]")
            if number < 0:
                raise ValueError("triangle values must be nonnegative; handle recoveries explicitly before calling")
            if require_cumulative and previous is not None and number < previous:
                raise ValueError(
                    f"cumulative triangle decreases at row {r}, column {c}; investigate recoveries or basis changes"
                )
            previous = number
            converted.append(number)
        result.append(converted)
    return result


def _last_observed_index(row: Sequence[float | None]) -> int:
    for index in range(len(row) - 1, -1, -1):
        if row[index] is not None:
            return index
    raise ValueError("triangle row has no observed values")


def compound_change_factors(changes: Sequence[Number], initial_level: Number = 1.0) -> list[float]:
    """Return cumulative multiplicative levels, including the initial level.

    Inputs: decimal changes in chronological order and an initial index.
    Output: ``[initial, after_change_1, ..., after_change_n]``.
    """
    level = _float(initial_level, "initial_level")
    if level <= 0:
        raise ValueError("initial_level must be positive")
    levels = [level]
    for change in _numbers(changes, "changes"):
        if change <= -1:
            raise ValueError("each change must be greater than -100%")
        level *= 1.0 + change
        levels.append(level)
    return levels


def rate_level_at_date(
    rate_changes: Sequence[Sequence[Any]],
    as_of: date | datetime | str,
    base_level: Number = 1.0,
) -> float:
    """Return the cumulative rate index effective at ``as_of``.

    ``rate_changes`` contains ``(effective_date, decimal_change)`` pairs. Changes on
    ``as_of`` are included. Duplicate dates are compounded in input order.
    """
    target = _as_date(as_of)
    parsed = sorted([(_as_date(item[0]), _float(item[1], "rate_change")) for item in rate_changes], key=lambda x: x[0])
    level = _float(base_level, "base_level")
    if level <= 0:
        raise ValueError("base_level must be positive")
    for effective, change in parsed:
        if change <= -1:
            raise ValueError("rate changes must be greater than -100%")
        if effective <= target:
            level *= 1.0 + change
    return level


def time_weighted_average_rate_level(
    period_start: date | datetime | str,
    period_end: date | datetime | str,
    rate_changes: Sequence[Sequence[Any]],
    daily_weights: Mapping[date | str, Number] | None = None,
    base_level: Number = 1.0,
) -> float:
    """Average a rate index over calendar days in ``[period_start, period_end)``.

    Use for written-rate or earning-date averages when the appropriate daily weights
    are known. With no weights, every day receives equal weight.
    """
    start, end = _as_date(period_start), _as_date(period_end)
    if end <= start:
        raise ValueError("period_end must be after period_start")
    weights = {_as_date(k): _float(v, "daily_weight") for k, v in (daily_weights or {}).items()}
    numerator = denominator = 0.0
    current = start
    while current < end:
        weight = weights.get(current, 0.0 if daily_weights is not None else 1.0)
        if weight < 0:
            raise ValueError("daily_weights must be nonnegative")
        numerator += rate_level_at_date(rate_changes, current, base_level) * weight
        denominator += weight
        current += timedelta(days=1)
    if denominator == 0:
        raise ValueError("daily_weights sum to zero")
    return numerator / denominator


def parallelogram_average_rate_level(
    experience_start: date | datetime | str,
    experience_end: date | datetime | str,
    rate_changes: Sequence[Sequence[Any]],
    policy_term_days: int = 365,
    writing_weights: Mapping[date | str, Number] | None = None,
    base_level: Number = 1.0,
) -> float:
    """Compute an earned average rate level using a discrete exact parallelogram.

    The calculation assumes the rate level is fixed at policy inception and premium
    earns uniformly over a fixed policy term. With no ``writing_weights``, policies
    are assumed written uniformly. Custom daily writing counts/premium weights provide
    the generalized non-uniform version. This is consistent with the extension-of-
    exposures interpretation under the stated assumptions.
    """
    start, end = _as_date(experience_start), _as_date(experience_end)
    if end <= start:
        raise ValueError("experience_end must be after experience_start")
    if not isinstance(policy_term_days, int) or policy_term_days <= 0:
        raise ValueError("policy_term_days must be a positive integer")
    weights = {_as_date(k): _float(v, "writing_weight") for k, v in (writing_weights or {}).items()}
    numerator = denominator = 0.0
    inception = start - timedelta(days=policy_term_days - 1)
    last_inception = end - timedelta(days=1)
    while inception <= last_inception:
        policy_end = inception + timedelta(days=policy_term_days)
        overlap_start = max(start, inception)
        overlap_end = min(end, policy_end)
        overlap_days = max(0, (overlap_end - overlap_start).days)
        write_weight = weights.get(inception, 0.0 if writing_weights is not None else 1.0)
        if write_weight < 0:
            raise ValueError("writing_weights must be nonnegative")
        earned_weight = write_weight * overlap_days / policy_term_days
        if earned_weight:
            numerator += rate_level_at_date(rate_changes, inception, base_level) * earned_weight
            denominator += earned_weight
        inception += timedelta(days=1)
    if denominator == 0:
        raise ValueError("no earned weight falls in the experience period")
    return numerator / denominator


def parallelogram_on_level_factor(
    experience_start: date | datetime | str,
    experience_end: date | datetime | str,
    target_date: date | datetime | str,
    rate_changes: Sequence[Sequence[Any]],
    policy_term_days: int = 365,
    writing_weights: Mapping[date | str, Number] | None = None,
    base_level: Number = 1.0,
) -> dict[str, float]:
    """Return average/current rate levels and the parallelogram on-level factor."""
    average = parallelogram_average_rate_level(
        experience_start, experience_end, rate_changes, policy_term_days, writing_weights, base_level
    )
    target = rate_level_at_date(rate_changes, target_date, base_level)
    return {"average_rate_level": average, "target_rate_level": target, "on_level_factor": target / average}


def extension_of_exposures_on_level_factor(
    historical_premiums: Sequence[Number], rerated_current_premiums: Sequence[Number]
) -> dict[str, float]:
    """Compute the ideal extension-of-exposures OLF from policy-level rerating.

    Historical and rerated amounts must cover the identical policy/exposure records.
    Output includes both totals and ``rerated_total / historical_total``.
    """
    historical = _numbers(historical_premiums, "historical_premiums", nonnegative=True)
    rerated = _numbers(rerated_current_premiums, "rerated_current_premiums", nonnegative=True)
    if len(historical) != len(rerated):
        raise ValueError("historical and rerated lists must have the same length")
    old_total, new_total = sum(historical), sum(rerated)
    return {"historical_total": old_total, "rerated_total": new_total, "on_level_factor": _safe_div(new_total, old_total, "premium")}


def rate_book_on_level_factor(
    premium_by_rate_book: Sequence[Number], rate_book_levels: Sequence[Number], target_rate_level: Number
) -> dict[str, float]:
    """On-level premium grouped by applicable historical rate-book index.

    Each premium group is multiplied by ``target_rate_level / rate_book_level``.
    This requires premium to be assigned to the rate book that priced it.
    """
    premium = _numbers(premium_by_rate_book, "premium_by_rate_book", nonnegative=True)
    levels = _numbers(rate_book_levels, "rate_book_levels")
    if len(premium) != len(levels):
        raise ValueError("premium and rate_book_levels must have the same length")
    if any(level <= 0 for level in levels):
        raise ValueError("rate_book_levels must be positive")
    target = _float(target_rate_level, "target_rate_level")
    if target <= 0:
        raise ValueError("target_rate_level must be positive")
    adjusted = [p * target / level for p, level in zip(premium, levels)]
    total = sum(premium)
    return {"historical_total": total, "on_level_total": sum(adjusted), "on_level_factor": _safe_div(sum(adjusted), total, "premium")}


def annual_rate_change_on_level_factors(
    years: Sequence[int],
    annual_rate_changes: Sequence[Number],
    target_year: int | None = None,
    average_method: str = "uniform",
    base_level: Number = 1.0,
) -> dict[str, Any]:
    """Approximate OLFs when only one aggregate rate change per year is available.

    ``uniform`` uses the average of beginning- and end-of-year cumulative levels;
    ``year_end`` uses the end-of-year level. This is an approximation, not a substitute
    for effective dates and policy earning information. Confirm the rate-change basis.
    """
    if len(years) != len(annual_rate_changes) or not years:
        raise ValueError("years and annual_rate_changes must have equal nonzero length")
    if len(set(years)) != len(years) or list(years) != sorted(years):
        raise ValueError("years must be unique and ascending")
    changes = _numbers(annual_rate_changes, "annual_rate_changes")
    levels = compound_change_factors(changes, base_level)
    end_levels = levels[1:]
    if average_method == "uniform":
        average_levels = [(levels[i] + levels[i + 1]) / 2.0 for i in range(len(years))]
    elif average_method == "year_end":
        average_levels = end_levels[:]
    else:
        raise ValueError("average_method must be 'uniform' or 'year_end'")
    target = target_year if target_year is not None else years[-1]
    if target not in years:
        raise ValueError("target_year must be one of years")
    target_level = end_levels[list(years).index(target)]
    rows = [
        {
            "year": year,
            "beginning_level": levels[i],
            "end_level": end_levels[i],
            "average_level": average_levels[i],
            "on_level_factor": target_level / average_levels[i],
        }
        for i, year in enumerate(years)
    ]
    return {"target_year": target, "target_rate_level": target_level, "average_method": average_method, "rows": rows}


def trend_factor(
    from_date: date | datetime | str,
    to_date: date | datetime | str,
    annual_trend: Number,
    method: str = "compound",
    day_count: Number = 365.25,
) -> float:
    """Trend a quantity between dates using compound or simple annual trend."""
    start, end = _as_date(from_date), _as_date(to_date)
    years = (end - start).days / _float(day_count, "day_count")
    trend = _float(annual_trend, "annual_trend")
    if method == "compound":
        if trend <= -1:
            raise ValueError("compound annual_trend must exceed -100%")
        return (1.0 + trend) ** years
    if method == "simple":
        result = 1.0 + trend * years
        if result < 0:
            raise ValueError("simple trend factor cannot be negative")
        return result
    raise ValueError("method must be 'compound' or 'simple'")


def apply_factors(values: Sequence[Number], *factors: Number | Sequence[Number]) -> list[float]:
    """Multiply each value by one or more scalar/vector factors."""
    result = _numbers(values, "values")
    for index, factor in enumerate(factors):
        vector = _broadcast(factor, len(result), f"factor_{index + 1}")
        result = [value * multiplier for value, multiplier in zip(result, vector)]
    return result


def on_level_and_trend_premium(
    earned_premium: Sequence[Number],
    on_level_factors: Number | Sequence[Number],
    exposure_trend_factors: Number | Sequence[Number] = 1.0,
) -> dict[str, list[float]]:
    """Return nominal, on-leveled, and on-leveled/trended earned premium."""
    premium = _numbers(earned_premium, "earned_premium", nonnegative=True)
    olf = _broadcast(on_level_factors, len(premium), "on_level_factors")
    trend = _broadcast(exposure_trend_factors, len(premium), "exposure_trend_factors")
    on_level = apply_factors(premium, olf)
    return {"nominal_earned_premium": premium, "on_level_earned_premium": on_level, "trended_on_level_earned_premium": apply_factors(on_level, trend)}


def trend_losses(
    losses: Sequence[Number],
    from_dates: Sequence[date | datetime | str],
    target_date: date | datetime | str,
    annual_loss_trend: Number | Sequence[Number],
    method: str = "compound",
) -> dict[str, list[float]]:
    """Trend aggregate or individual losses to a common target date."""
    loss = _numbers(losses, "losses", nonnegative=True)
    if len(from_dates) != len(loss):
        raise ValueError("from_dates must match losses")
    trends = _broadcast(annual_loss_trend, len(loss), "annual_loss_trend")
    factors = [trend_factor(d, target_date, t, method) for d, t in zip(from_dates, trends)]
    return {"trend_factors": factors, "trended_losses": apply_factors(loss, factors)}


def _piecewise_trend_factor_impl(
    from_date: date | datetime | str,
    to_date: date | datetime | str,
    annual_trend_schedule: Sequence[Sequence[Any]],
    method: str,
    day_count: Number,
) -> float:
    start, end = _as_date(from_date), _as_date(to_date)
    if end < start:
        raise ValueError("to_date must be on or after from_date")
    if end == start:
        return 1.0
    if method not in {"compound", "simple"}:
        raise ValueError("method must be 'compound' or 'simple'")
    denominator = _float(day_count, "day_count")
    if denominator <= 0:
        raise ValueError("day_count must be positive")
    parsed = sorted(
        [(_as_date(item[0]), _float(item[1], "annual_trend")) for item in annual_trend_schedule],
        key=lambda item: item[0],
    )
    if not parsed:
        raise ValueError("annual_trend_schedule cannot be empty")
    if len({effective for effective, _ in parsed}) != len(parsed):
        raise ValueError("annual_trend_schedule effective dates must be unique")
    if parsed[0][0] > start:
        raise ValueError("annual_trend_schedule must begin on or before from_date")
    if method == "compound" and any(rate <= -1 for _, rate in parsed):
        raise ValueError("compound annual trends must exceed -100%")

    boundaries = [effective for effective, _ in parsed if start < effective < end]
    segment_start = start
    factor = 1.0
    for segment_end in boundaries + [end]:
        applicable_rate = next(rate for effective, rate in reversed(parsed) if effective <= segment_start)
        years = (segment_end - segment_start).days / denominator
        segment_factor = (
            (1.0 + applicable_rate) ** years
            if method == "compound"
            else 1.0 + applicable_rate * years
        )
        if segment_factor < 0:
            raise ValueError("simple trend segment produced a negative factor")
        factor *= segment_factor
        segment_start = segment_end
    return factor


def piecewise_trend_factor(
    from_date: date | datetime | str,
    to_date: date | datetime | str,
    annual_trend_schedule: Sequence[Sequence[Any]],
    method: str = "compound",
    day_count: Number = 365.25,
) -> float:
    """Trend between dates using dated annual trend selections.

    ``annual_trend_schedule`` contains ``(effective_date, annual_decimal_trend)``
    pairs. Each selection applies from its effective date until the next selection;
    the last selection continues through ``to_date``. The schedule must begin on or
    before ``from_date``. Segment factors are multiplied, including under ``simple``.
    """
    return _piecewise_trend_factor_impl(from_date, to_date, annual_trend_schedule, method, day_count)


def trend_losses_piecewise(
    losses: Sequence[Number],
    from_dates: Sequence[date | datetime | str],
    target_date: date | datetime | str,
    annual_trend_schedules: Mapping[Any, Sequence[Sequence[Any]]],
    schedule_keys: Sequence[Any] | None = None,
    method: str = "compound",
    day_count: Number = 365.25,
) -> dict[str, Any]:
    """Trend individual claims under one or more dated trend schedules.

    ``annual_trend_schedules`` maps a LoB/segment key to dated annual selections.
    ``schedule_keys`` assigns one schedule to each claim. If omitted, exactly one
    schedule must be supplied. Output order matches the input claim order.
    """
    values = _numbers(losses, "losses", nonnegative=True)
    if len(from_dates) != len(values):
        raise ValueError("from_dates must match losses")
    if not annual_trend_schedules:
        raise ValueError("annual_trend_schedules cannot be empty")
    if schedule_keys is None:
        if len(annual_trend_schedules) != 1:
            raise ValueError("schedule_keys are required when multiple trend schedules are supplied")
        only_key = next(iter(annual_trend_schedules))
        keys = [only_key] * len(values)
    else:
        keys = list(schedule_keys)
        if len(keys) != len(values):
            raise ValueError("schedule_keys must match losses")
    missing = sorted({str(key) for key in keys if key not in annual_trend_schedules})
    if missing:
        raise ValueError(f"no annual trend schedule for keys {missing}")
    factors = [
        _piecewise_trend_factor_impl(from_value, target_date, annual_trend_schedules[key], method, day_count)
        for from_value, key in zip(from_dates, keys)
    ]
    return {
        "target_date": _as_date(target_date).isoformat(),
        "method": method,
        "schedule_keys": [str(key) for key in keys],
        "trend_factors": factors,
        "trended_losses": apply_factors(values, factors),
    }


def validate_triangle(triangle: Triangle, cumulative: bool = True) -> dict[str, Any]:
    """Validate shape/contiguity and report cumulative decreases without modifying data."""
    errors: list[str] = []
    warnings: list[str] = []
    try:
        rows = _coerce_triangle(triangle, require_cumulative=False)
    except ValueError as exc:
        return {"valid": False, "errors": [str(exc)], "warnings": [], "rows": len(triangle), "columns": 0}
    observed = 0
    for r, row in enumerate(rows):
        previous = None
        for c, value in enumerate(row):
            if value is None:
                continue
            observed += 1
            if cumulative and previous is not None and value < previous:
                warnings.append(f"row {r} cumulative value decreases at column {c}")
            previous = value
    return {"valid": not errors, "errors": errors, "warnings": warnings, "rows": len(rows), "columns": len(rows[0]), "observed_cells": observed}


def incremental_to_cumulative(incremental_triangle: Triangle) -> list[list[float | None]]:
    """Convert an incremental triangle to cumulative form."""
    rows = _coerce_triangle(incremental_triangle)
    result: list[list[float | None]] = []
    for row in rows:
        running = 0.0
        converted: list[float | None] = []
        for value in row:
            if value is None:
                converted.append(None)
            else:
                running += value
                converted.append(running)
        result.append(converted)
    return result


def cumulative_to_incremental(cumulative_triangle: Triangle) -> list[list[float | None]]:
    """Convert a cumulative triangle to incremental form."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    result: list[list[float | None]] = []
    for row in rows:
        previous = 0.0
        converted: list[float | None] = []
        for value in row:
            if value is None:
                converted.append(None)
            else:
                converted.append(value - previous)
                previous = value
        result.append(converted)
    return result


def link_ratio_triangle(cumulative_triangle: Triangle) -> list[list[float | None]]:
    """Return observed age-to-age link ratios for a cumulative triangle."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    output: list[list[float | None]] = []
    for row in rows:
        ratios: list[float | None] = []
        for current, nxt in zip(row[:-1], row[1:]):
            if current is None or nxt is None or current == 0:
                ratios.append(None)
            else:
                ratios.append(nxt / current)
        output.append(ratios)
    return output


def development_factor_diagnostics(cumulative_triangle: Triangle) -> dict[str, Any]:
    """Summarize link-ratio count, volume/simple/median selection, and coefficient of variation."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    diagnostics = []
    for column in range(len(rows[0]) - 1):
        pairs = [(row[column], row[column + 1]) for row in rows if row[column] not in (None, 0) and row[column + 1] is not None]
        ratios = [nxt / current for current, nxt in pairs]  # type: ignore[operator]
        if not ratios:
            diagnostics.append({"development_index": column, "count": 0, "volume": None, "simple": None, "median": None, "cv": None})
            continue
        mean = sum(ratios) / len(ratios)
        variance = sum((x - mean) ** 2 for x in ratios) / (len(ratios) - 1) if len(ratios) > 1 else 0.0
        volume = sum(nxt for _, nxt in pairs) / sum(current for current, _ in pairs)  # type: ignore[misc]
        diagnostics.append({"development_index": column, "count": len(ratios), "volume": volume, "simple": mean, "median": median(ratios), "cv": sqrt(variance) / mean if mean else None})
    return {"development_periods": diagnostics}


def select_age_to_age_factors(
    cumulative_triangle: Triangle,
    method: str = "volume",
    latest_n: int | None = None,
    exclude_high_low: bool = False,
    selected_factors: Sequence[Number] | None = None,
) -> dict[str, Any]:
    """Select deterministic age-to-age factors.

    Methods: ``volume``, ``simple``, ``geometric``, ``median``, or ``trimmed``.
    ``latest_n`` uses the most recent observed ratios. ``exclude_high_low`` removes
    one minimum and maximum ratio when at least three remain. Explicit selections
    override calculated selections but observed diagnostics are still returned.
    """
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    width = len(rows[0])
    if latest_n is not None and latest_n <= 0:
        raise ValueError("latest_n must be positive")
    explicit: list[float] | None = None
    if selected_factors is not None:
        explicit = _numbers(selected_factors, "selected_factors")
        if len(explicit) != width - 1 or any(factor <= 0 for factor in explicit):
            raise ValueError("selected_factors must have columns-1 positive entries")
    selected: list[float] = []
    observations: list[dict[str, Any]] = []
    for column in range(width - 1):
        pairs = [(row[column], row[column + 1]) for row in rows if row[column] not in (None, 0) and row[column + 1] is not None]
        if latest_n is not None:
            pairs = pairs[-latest_n:]
        ratios = [nxt / current for current, nxt in pairs]  # type: ignore[operator]
        if not ratios and explicit is None:
            raise ValueError(f"no usable link ratios for development column {column}")
        filtered = ratios[:]
        if (exclude_high_low or method == "trimmed") and len(filtered) >= 3:
            filtered.remove(min(filtered))
            filtered.remove(max(filtered))
        if explicit is not None:
            factor = explicit[column]
        elif method == "volume":
            kept_pairs = pairs
            if exclude_high_low and len(ratios) >= 3:
                low, high = min(ratios), max(ratios)
                removed_low = removed_high = False
                kept_pairs = []
                for pair, ratio in zip(pairs, ratios):
                    if ratio == low and not removed_low:
                        removed_low = True
                    elif ratio == high and not removed_high:
                        removed_high = True
                    else:
                        kept_pairs.append(pair)
            factor = sum(nxt for _, nxt in kept_pairs) / sum(current for current, _ in kept_pairs)  # type: ignore[misc]
        elif method in {"simple", "trimmed"}:
            factor = sum(filtered) / len(filtered)
        elif method == "geometric":
            if any(x <= 0 for x in filtered):
                raise ValueError("geometric link ratio average requires positive ratios")
            factor = prod(filtered) ** (1.0 / len(filtered))
        elif method == "median":
            factor = median(filtered)
        else:
            raise ValueError("unknown factor selection method")
        if factor <= 0:
            raise ValueError("selected development factors must be positive")
        selected.append(factor)
        observations.append({"development_index": column, "ratios": ratios, "ratios_used": filtered, "selected_factor": factor})
    return {"method": "explicit" if explicit is not None else method, "selected_factors": selected, "observations": observations}


def cumulative_development_factors(age_to_age_factors: Sequence[Number], tail_factor: Number = 1.0) -> list[float]:
    """Convert age-to-age factors to CDFs for every development age."""
    factors = _numbers(age_to_age_factors, "age_to_age_factors")
    tail = _float(tail_factor, "tail_factor")
    if any(f <= 0 for f in factors) or tail <= 0:
        raise ValueError("development and tail factors must be positive")
    cdfs = [1.0] * (len(factors) + 1)
    running = tail
    cdfs[-1] = tail
    for index in range(len(factors) - 1, -1, -1):
        running *= factors[index]
        cdfs[index] = running
    return cdfs


def reported_percentages(cumulative_factors: Sequence[Number]) -> list[float]:
    """Return percent reported/paid as the reciprocal of each CDF."""
    cdfs = _numbers(cumulative_factors, "cumulative_factors")
    if any(cdf <= 0 for cdf in cdfs):
        raise ValueError("CDFs must be positive")
    return [1.0 / cdf for cdf in cdfs]


def latest_diagonal(cumulative_triangle: Triangle) -> list[dict[str, float | int]]:
    """Return each origin row's latest observed development index and value."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    result = []
    for origin, row in enumerate(rows):
        index = _last_observed_index(row)
        result.append({"origin_index": origin, "development_index": index, "latest": float(row[index])})  # type: ignore[arg-type]
    return result


def fill_lower_triangle_chain_ladder(
    cumulative_triangle: Triangle, age_to_age_factors: Sequence[Number]
) -> list[list[float]]:
    """Fill future cumulative cells using selected age-to-age factors."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    factors = _numbers(age_to_age_factors, "age_to_age_factors")
    if len(factors) != len(rows[0]) - 1 or any(f <= 0 for f in factors):
        raise ValueError("age_to_age_factors must have columns-1 positive entries")
    completed: list[list[float]] = []
    for row in rows:
        last = _last_observed_index(row)
        filled = [0.0 if value is None else value for value in row]
        for column in range(last, len(factors)):
            filled[column + 1] = filled[column] * factors[column]
        completed.append(filled)
    return completed


def chain_ladder_ultimate(
    cumulative_triangle: Triangle,
    method: str = "volume",
    latest_n: int | None = None,
    exclude_high_low: bool = False,
    selected_factors: Sequence[Number] | None = None,
    tail_factor: Number = 1.0,
) -> dict[str, Any]:
    """Project cumulative loss/count triangle to ultimate with Chain Ladder."""
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    selection = select_age_to_age_factors(rows, method, latest_n, exclude_high_low, selected_factors)
    factors = selection["selected_factors"]
    cdfs = cumulative_development_factors(factors, tail_factor)
    completed = fill_lower_triangle_chain_ladder(rows, factors)
    origins = []
    for origin, row in enumerate(rows):
        last = _last_observed_index(row)
        latest = float(row[last])  # type: ignore[arg-type]
        ultimate = latest * cdfs[last]
        origins.append({"origin_index": origin, "development_index": last, "latest": latest, "cdf": cdfs[last], "percent_reported": 1.0 / cdfs[last], "ultimate": ultimate, "ibnr": ultimate - latest})
    return {"selection": selection, "cdfs": cdfs, "completed_triangle": completed, "origins": origins, "total_latest": sum(x["latest"] for x in origins), "total_ultimate": sum(x["ultimate"] for x in origins), "total_ibnr": sum(x["ibnr"] for x in origins)}


def expected_loss_ultimate(exposure_basis: Sequence[Number], expected_loss_ratio: Number | Sequence[Number]) -> dict[str, Any]:
    """Expected-loss method: ultimate equals exposure basis times a priori loss ratio."""
    exposure = _numbers(exposure_basis, "exposure_basis", nonnegative=True)
    ratios = _broadcast(expected_loss_ratio, len(exposure), "expected_loss_ratio")
    if any(r < 0 for r in ratios):
        raise ValueError("expected_loss_ratio must be nonnegative")
    ultimate = apply_factors(exposure, ratios)
    return {"exposure_basis": exposure, "expected_loss_ratio": ratios, "ultimate": ultimate, "total_ultimate": sum(ultimate)}


def bornhuetter_ferguson_ultimate(
    reported_or_paid_loss: Sequence[Number],
    cumulative_factors: Sequence[Number],
    exposure_basis: Sequence[Number],
    apriori_loss_ratio: Number | Sequence[Number],
) -> dict[str, Any]:
    """Reported/paid Bornhuetter-Ferguson ultimate by origin period."""
    actual = _numbers(reported_or_paid_loss, "reported_or_paid_loss", nonnegative=True)
    cdfs = _numbers(cumulative_factors, "cumulative_factors")
    exposure = _numbers(exposure_basis, "exposure_basis", nonnegative=True)
    if not (len(actual) == len(cdfs) == len(exposure)):
        raise ValueError("loss, CDF, and exposure lists must have equal length")
    expected = expected_loss_ultimate(exposure, apriori_loss_ratio)["ultimate"]
    pct = reported_percentages(cdfs)
    ibnr = [e * (1.0 - p) for e, p in zip(expected, pct)]
    ultimate = [a + u for a, u in zip(actual, ibnr)]
    rows = [{"origin_index": i, "actual": actual[i], "cdf": cdfs[i], "percent_reported": pct[i], "expected_ultimate": expected[i], "expected_unreported_or_unpaid": ibnr[i], "ultimate": ultimate[i]} for i in range(len(actual))]
    return {"origins": rows, "ultimate": ultimate, "total_ultimate": sum(ultimate), "total_ibnr": sum(ibnr)}


def cape_cod_ultimate(
    reported_or_paid_loss: Sequence[Number],
    cumulative_factors: Sequence[Number],
    exposure_basis: Sequence[Number],
    year_weights: Sequence[Number] | None = None,
) -> dict[str, Any]:
    """Cape Cod method with optional deterministic calendar/origin weights.

    The fitted ELR is weighted reported loss divided by weighted used-up exposure,
    where used-up exposure is exposure times percent reported/paid.
    """
    actual = _numbers(reported_or_paid_loss, "reported_or_paid_loss", nonnegative=True)
    cdfs = _numbers(cumulative_factors, "cumulative_factors")
    exposure = _numbers(exposure_basis, "exposure_basis", nonnegative=True)
    if not (len(actual) == len(cdfs) == len(exposure)):
        raise ValueError("loss, CDF, and exposure lists must have equal length")
    weights = [1.0] * len(actual) if year_weights is None else _numbers(year_weights, "year_weights", nonnegative=True)
    if len(weights) != len(actual) or sum(weights) == 0:
        raise ValueError("year_weights must match data and contain positive total weight")
    pct = reported_percentages(cdfs)
    used_up = [e * p for e, p in zip(exposure, pct)]
    fitted_elr = _safe_div(sum(a * w for a, w in zip(actual, weights)), sum(u * w for u, w in zip(used_up, weights)), "Cape Cod used-up exposure")
    expected = [e * fitted_elr for e in exposure]
    ibnr = [e * (1.0 - p) for e, p in zip(expected, pct)]
    ultimate = [a + u for a, u in zip(actual, ibnr)]
    rows = [{"origin_index": i, "actual": actual[i], "percent_reported": pct[i], "used_up_exposure": used_up[i], "expected_ultimate": expected[i], "ultimate": ultimate[i]} for i in range(len(actual))]
    return {"fitted_expected_loss_ratio": fitted_elr, "origins": rows, "ultimate": ultimate, "total_ultimate": sum(ultimate), "total_ibnr": sum(ibnr)}


def benktander_ultimate(
    reported_or_paid_loss: Sequence[Number],
    cumulative_factors: Sequence[Number],
    initial_expected_ultimate: Sequence[Number],
    iterations: int = 2,
) -> dict[str, Any]:
    """Iterative BF/Benktander ultimate.

    Iteration 1 is BF using the supplied prior. Iteration 2 is the usual "BF of BF"
    Benktander result. Further iterations approach the development-method result.
    """
    actual = _numbers(reported_or_paid_loss, "reported_or_paid_loss", nonnegative=True)
    cdfs = _numbers(cumulative_factors, "cumulative_factors")
    prior = _numbers(initial_expected_ultimate, "initial_expected_ultimate", nonnegative=True)
    if not (len(actual) == len(cdfs) == len(prior)):
        raise ValueError("actual, CDF, and prior lists must have equal length")
    if not isinstance(iterations, int) or iterations < 1:
        raise ValueError("iterations must be a positive integer")
    pct = reported_percentages(cdfs)
    history = []
    current = prior[:]
    for iteration in range(1, iterations + 1):
        current = [a + (1.0 - p) * expected for a, p, expected in zip(actual, pct, current)]
        history.append({"iteration": iteration, "ultimate": current[:]})
    return {"iterations": iterations, "history": history, "ultimate": current, "total_ultimate": sum(current), "total_ibnr": sum(u - a for u, a in zip(current, actual))}


def loglinear_tail_factor(
    age_to_age_factors: Sequence[Number],
    future_periods: int = 5,
    latest_points: int | None = None,
    enforce_decay: bool = True,
) -> dict[str, Any]:
    """Extrapolate a deterministic tail via log(factor - 1) linear regression.

    Only observed factors greater than 1 are usable. Tail-factor uncertainty is not
    estimated. The user must confirm development-period spacing and appropriateness.
    """
    factors = _numbers(age_to_age_factors, "age_to_age_factors")
    points = [(i, log(f - 1.0)) for i, f in enumerate(factors) if f > 1.0]
    if latest_points is not None:
        if latest_points < 2:
            raise ValueError("latest_points must be at least 2")
        points = points[-latest_points:]
    if len(points) < 2 or future_periods < 1:
        raise ValueError("at least two factors greater than 1 and one future period are required")
    xbar = sum(x for x, _ in points) / len(points)
    ybar = sum(y for _, y in points) / len(points)
    denominator = sum((x - xbar) ** 2 for x, _ in points)
    if denominator == 0:
        raise ValueError("tail regression requires distinct development indices")
    slope = sum((x - xbar) * (y - ybar) for x, y in points) / denominator
    intercept = ybar - slope * xbar
    if enforce_decay and slope >= 0:
        raise ValueError("fitted tail does not decay; select a tail explicitly or revise points")
    start_index = len(factors)
    projected = [1.0 + exp(intercept + slope * (start_index + step)) for step in range(future_periods)]
    return {"intercept": intercept, "slope": slope, "projected_age_to_age_factors": projected, "tail_factor": prod(projected), "points_used": [{"index": x, "log_factor_minus_one": y} for x, y in points]}


def mack_chain_ladder(cumulative_triangle: Triangle, tail_factor: Number = 1.0) -> dict[str, Any]:
    """Mack distribution-free Chain Ladder point estimates and standard errors.

    Volume-weighted factors are used. The final process variance is extrapolated using
    the standard Mack minimum rule when necessary. Tail uncertainty is not included.
    Results are diagnostics and are not a substitute for residual review.
    """
    rows = _coerce_triangle(cumulative_triangle, require_cumulative=False)
    if len(rows) < 3 or len(rows[0]) < 3:
        raise ValueError("Mack Chain Ladder requires at least a 3x3 triangle")
    selection = select_age_to_age_factors(rows, method="volume")
    factors = selection["selected_factors"]
    width = len(rows[0])
    sigma2: list[float | None] = []
    denominators: list[float] = []
    for column in range(width - 1):
        pairs = [(row[column], row[column + 1]) for row in rows if row[column] not in (None, 0) and row[column + 1] is not None]
        denominators.append(sum(current for current, _ in pairs))  # type: ignore[misc]
        if len(pairs) >= 2:
            value = sum(current * (nxt / current - factors[column]) ** 2 for current, nxt in pairs) / (len(pairs) - 1)  # type: ignore[operator]
            sigma2.append(value)
        else:
            sigma2.append(None)
    for column in range(len(sigma2)):
        if sigma2[column] is not None:
            continue
        previous = [x for x in sigma2[:column] if x is not None]
        if len(previous) >= 2 and previous[-2] > 0:
            candidate = previous[-1] ** 2 / previous[-2]
            sigma2[column] = min(candidate, previous[-1], previous[-2])
        elif previous:
            sigma2[column] = previous[-1]
        else:
            sigma2[column] = 0.0
    completed = fill_lower_triangle_chain_ladder(rows, factors)
    cdfs = cumulative_development_factors(factors, tail_factor)
    origin_results = []
    for origin, row in enumerate(rows):
        last = _last_observed_index(row)
        latest = float(row[last])  # type: ignore[arg-type]
        ultimate = latest * cdfs[last]
        variance_sum = 0.0
        for column in range(last, width - 1):
            projected = completed[origin][column]
            if projected <= 0 or denominators[column] <= 0:
                continue
            variance_sum += float(sigma2[column]) / (factors[column] ** 2) * (1.0 / projected + 1.0 / denominators[column])
        msep = ultimate**2 * variance_sum
        origin_results.append({"origin_index": origin, "development_index": last, "latest": latest, "ultimate": ultimate, "ibnr": ultimate - latest, "mack_msep": msep, "mack_standard_error": sqrt(max(0.0, msep))})
    total_msep = sum(row["mack_msep"] for row in origin_results)
    for i in range(len(origin_results)):
        for j in range(i + 1, len(origin_results)):
            last_common = max(origin_results[i]["development_index"], origin_results[j]["development_index"])
            covariance_term = 0.0
            for column in range(last_common, width - 1):
                if denominators[column] > 0:
                    covariance_term += float(sigma2[column]) / (factors[column] ** 2 * denominators[column])
            total_msep += 2.0 * origin_results[i]["ultimate"] * origin_results[j]["ultimate"] * covariance_term
    return {"selected_factors": factors, "sigma_squared": [float(x) for x in sigma2], "tail_factor": float(tail_factor), "tail_uncertainty_included": False, "origins": origin_results, "total_ultimate": sum(x["ultimate"] for x in origin_results), "total_ibnr": sum(x["ibnr"] for x in origin_results), "total_mack_msep": total_msep, "total_mack_standard_error": sqrt(max(0.0, total_msep))}


def compare_paid_incurred_ultimates(
    paid_cumulative_triangle: Triangle,
    incurred_cumulative_triangle: Triangle,
    paid_weight: Number = 0.5,
    paid_selected_factors: Sequence[Number] | None = None,
    incurred_selected_factors: Sequence[Number] | None = None,
    tail_factor_paid: Number = 1.0,
    tail_factor_incurred: Number = 1.0,
) -> dict[str, Any]:
    """Compare separate paid and incurred Chain Ladder projections and a chosen blend.

    This is not Munich Chain Ladder. ``paid_weight`` is an explicit actuarial selection.
    """
    weight = _float(paid_weight, "paid_weight")
    if not 0 <= weight <= 1:
        raise ValueError("paid_weight must be between 0 and 1")
    paid = chain_ladder_ultimate(paid_cumulative_triangle, selected_factors=paid_selected_factors, tail_factor=tail_factor_paid)
    incurred = chain_ladder_ultimate(incurred_cumulative_triangle, selected_factors=incurred_selected_factors, tail_factor=tail_factor_incurred)
    if len(paid["origins"]) != len(incurred["origins"]):
        raise ValueError("paid and incurred triangles must have the same origin count")
    rows = []
    for p, inc in zip(paid["origins"], incurred["origins"]):
        blended = weight * p["ultimate"] + (1.0 - weight) * inc["ultimate"]
        rows.append({"origin_index": p["origin_index"], "paid_ultimate": p["ultimate"], "incurred_ultimate": inc["ultimate"], "paid_to_incurred_ultimate": _safe_div(p["ultimate"], inc["ultimate"], "incurred ultimate"), "selected_blended_ultimate": blended})
    return {"paid_weight": weight, "incurred_weight": 1.0 - weight, "paid": paid, "incurred": incurred, "origins": rows, "total_selected_ultimate": sum(x["selected_blended_ultimate"] for x in rows)}


def cap_losses(losses: Sequence[Number], cap: Number) -> list[float]:
    """Cap each nonnegative loss at a per-claim or per-occurrence cap."""
    values = _numbers(losses, "losses", nonnegative=True)
    limit = _float(cap, "cap")
    if limit < 0:
        raise ValueError("cap must be nonnegative")
    return [min(value, limit) for value in values]


def _apply_layer_value(loss: float, attachment: float, limit: float, share: float) -> float:
    return min(max(loss - attachment, 0.0), limit) * share


def apply_per_claim_layer(
    gross_losses: Sequence[Number],
    attachment: Number,
    limit: Number,
    share: Number = 1.0,
    allocated_expenses: Sequence[Number] | None = None,
    alae_treatment: str = "excluded",
) -> dict[str, Any]:
    """Apply a per-claim layer and optional ALAE treatment.

    ``excluded`` excludes ALAE. ``pro_rata`` cedes ALAE in proportion to ceded
    indemnity divided by gross indemnity. ``included`` assumes ``gross_losses`` already
    includes ALAE and therefore requires ``allocated_expenses=None``. Contract wording
    must be confirmed before use.
    """
    losses = _numbers(gross_losses, "gross_losses", nonnegative=True)
    attach, layer_limit, participation = _float(attachment, "attachment"), _float(limit, "limit"), _float(share, "share")
    if attach < 0 or layer_limit < 0 or not 0 <= participation <= 1:
        raise ValueError("attachment/limit must be nonnegative and share must be in [0,1]")
    if alae_treatment not in {"excluded", "pro_rata", "included"}:
        raise ValueError("alae_treatment must be excluded, pro_rata, or included")
    if alae_treatment == "included" and allocated_expenses is not None:
        raise ValueError("do not provide allocated_expenses when gross_losses already includes ALAE")
    expenses = [0.0] * len(losses) if allocated_expenses is None else _numbers(allocated_expenses, "allocated_expenses", nonnegative=True)
    if len(expenses) != len(losses):
        raise ValueError("allocated_expenses must match gross_losses")
    rows = []
    for index, (loss, expense_value) in enumerate(zip(losses, expenses)):
        indemnity = _apply_layer_value(loss, attach, layer_limit, participation)
        ceded_expense = expense_value * indemnity / loss if alae_treatment == "pro_rata" and loss > 0 else 0.0
        total = indemnity + ceded_expense
        rows.append({"claim_index": index, "gross_loss": loss, "gross_allocated_expense": expense_value, "ceded_indemnity": indemnity, "ceded_allocated_expense": ceded_expense, "ceded_total": total})
    return {"attachment": attach, "limit": layer_limit, "share": participation, "alae_treatment": alae_treatment, "claims": rows, "total_ceded_indemnity": sum(x["ceded_indemnity"] for x in rows), "total_ceded_allocated_expense": sum(x["ceded_allocated_expense"] for x in rows), "total_ceded": sum(x["ceded_total"] for x in rows)}


def claim_level_xol_summary(
    claim_ids: Sequence[Any],
    reported_losses: Sequence[Number],
    group_fields: Mapping[str, Sequence[Any]],
    trend_factors: Number | Sequence[Number] = 1.0,
    development_factors: Number | Sequence[Number] = 1.0,
    attachment: Number | None = None,
    limit: Number | None = None,
    share: Number = 1.0,
    layer_development_order: str = "layer_then_develop",
) -> dict[str, Any]:
    """Build claim-level and grouped XoL reported/ultimate loss summaries.

    ``reported_losses`` is the explicitly selected claim loss basis (for example,
    gross incurred including ALAE). ``trend_factors`` and ``development_factors``
    may be scalar or claim-level vectors. ``group_fields`` maps requested dimension
    names (such as state, line_of_business, accident_year, or cat_non_cat) to one
    label per claim.

    When layer terms are supplied, reported layer losses are calculated separately
    from nominal and trended claim amounts. ``layer_then_develop`` then applies each
    claim's development factor to those layer losses, matching a common experience-
    rating workflow. ``develop_then_layer`` first develops ground-up claim amounts
    and then applies the layer. The order is material and must be an explicit
    actuarial selection. Run one call per layer when multiple layers are requested.
    """
    losses = _numbers(reported_losses, "reported_losses", nonnegative=True)
    n = len(losses)
    ids = [str(value) for value in claim_ids]
    if len(ids) != n:
        raise ValueError("claim_ids must match reported_losses")
    if any(not value.strip() for value in ids) or len(set(ids)) != len(ids):
        raise ValueError("claim_ids must be nonblank and unique")
    trend = _broadcast(trend_factors, n, "trend_factors")
    development = _broadcast(development_factors, n, "development_factors")
    if any(factor <= 0 for factor in trend):
        raise ValueError("trend_factors must be positive")
    if any(factor <= 0 for factor in development):
        raise ValueError("development_factors must be positive")
    if layer_development_order not in {"layer_then_develop", "develop_then_layer"}:
        raise ValueError("layer_development_order must be layer_then_develop or develop_then_layer")

    field_names = list(group_fields)
    if any(not isinstance(name, str) or not name.strip() for name in field_names) or len(set(field_names)) != len(field_names):
        raise ValueError("group field names must be nonblank unique strings")
    normalized_groups: dict[str, list[Any]] = {}
    for name, values in group_fields.items():
        labels = list(values)
        if len(labels) != n:
            raise ValueError(f"group field {name!r} must match reported_losses")
        normalized: list[Any] = []
        for label in labels:
            if label is None or (isinstance(label, str) and not label.strip()):
                raise ValueError(f"group field {name!r} contains a blank label")
            if isinstance(label, datetime):
                normalized.append(label.isoformat())
            elif isinstance(label, date):
                normalized.append(label.isoformat())
            elif isinstance(label, (str, int, float, bool)):
                normalized.append(label)
            else:
                normalized.append(str(label))
        normalized_groups[str(name)] = normalized

    layer_applied = attachment is not None or limit is not None
    if layer_applied and (attachment is None or limit is None):
        raise ValueError("attachment and limit must both be supplied for layer results")
    attach = _float(attachment, "attachment") if attachment is not None else 0.0
    layer_limit = _float(limit, "limit") if limit is not None else 0.0
    participation = _float(share, "share")
    if attach < 0 or layer_limit < 0 or not 0 <= participation <= 1:
        raise ValueError("attachment/limit must be nonnegative and share must be in [0,1]")

    claim_rows: list[dict[str, Any]] = []
    for index, (claim_id, reported, trend_factor_value, development_factor_value) in enumerate(
        zip(ids, losses, trend, development)
    ):
        trended_reported = reported * trend_factor_value
        nominal_ultimate = reported * development_factor_value
        trended_ultimate = trended_reported * development_factor_value
        row: dict[str, Any] = {
            "claim_index": index,
            "claim_id": claim_id,
            "groups": {name: normalized_groups[name][index] for name in field_names},
            "trend_factor": trend_factor_value,
            "development_factor": development_factor_value,
            "nominal_reported_ground_up": reported,
            "trended_reported_ground_up": trended_reported,
            "nominal_ultimate_ground_up": nominal_ultimate,
            "trended_ultimate_ground_up": trended_ultimate,
        }
        if layer_applied:
            nominal_reported_layer = _apply_layer_value(reported, attach, layer_limit, participation)
            trended_reported_layer = _apply_layer_value(trended_reported, attach, layer_limit, participation)
            if layer_development_order == "layer_then_develop":
                nominal_ultimate_layer = nominal_reported_layer * development_factor_value
                trended_ultimate_layer = trended_reported_layer * development_factor_value
            else:
                nominal_ultimate_layer = _apply_layer_value(nominal_ultimate, attach, layer_limit, participation)
                trended_ultimate_layer = _apply_layer_value(trended_ultimate, attach, layer_limit, participation)
            row.update(
                {
                    "nominal_reported_layer": nominal_reported_layer,
                    "trended_reported_layer": trended_reported_layer,
                    "nominal_ultimate_layer": nominal_ultimate_layer,
                    "trended_ultimate_layer": trended_ultimate_layer,
                }
            )
        claim_rows.append(row)

    measures = [
        "nominal_reported_ground_up",
        "trended_reported_ground_up",
        "nominal_ultimate_ground_up",
        "trended_ultimate_ground_up",
    ]
    if layer_applied:
        measures.extend(
            [
                "nominal_reported_layer",
                "trended_reported_layer",
                "nominal_ultimate_layer",
                "trended_ultimate_layer",
            ]
        )

    def summarize(indices: Sequence[int]) -> dict[str, Any]:
        summary: dict[str, Any] = {"claim_count": len(indices)}
        for measure in measures:
            summary[measure] = sum(float(claim_rows[index][measure]) for index in indices)
        if layer_applied:
            summary["claims_with_nominal_reported_layer_loss"] = sum(
                float(claim_rows[index]["nominal_reported_layer"]) > 0 for index in indices
            )
            summary["claims_with_trended_reported_layer_loss"] = sum(
                float(claim_rows[index]["trended_reported_layer"]) > 0 for index in indices
            )
        return summary

    grouped_indices: dict[tuple[Any, ...], list[int]] = {}
    for index in range(n):
        key = tuple(normalized_groups[name][index] for name in field_names)
        grouped_indices.setdefault(key, []).append(index)
    group_rows = []
    for key in sorted(grouped_indices, key=lambda item: tuple(str(value) for value in item)):
        group_rows.append(
            {
                "group": {name: value for name, value in zip(field_names, key)},
                **summarize(grouped_indices[key]),
            }
        )

    return {
        "loss_basis": "reported_losses supplied by caller",
        "group_fields": field_names,
        "layer_applied": layer_applied,
        "layer_terms": (
            {"attachment": attach, "limit": layer_limit, "share": participation}
            if layer_applied
            else None
        ),
        "layer_development_order": layer_development_order if layer_applied else None,
        "claims": claim_rows,
        "groups": group_rows,
        "totals": summarize(list(range(n))),
    }


def aggregate_occurrence_losses(losses: Sequence[Number], event_ids: Sequence[Any]) -> dict[str, Any]:
    """Aggregate claim losses by explicit occurrence/event identifier."""
    values = _numbers(losses, "losses", nonnegative=True)
    if len(values) != len(event_ids):
        raise ValueError("event_ids must match losses")
    totals: dict[str, float] = {}
    for loss, event_id in zip(values, event_ids):
        key = str(event_id)
        if not key or key.lower() == "none":
            raise ValueError("every loss requires a nonblank event_id")
        totals[key] = totals.get(key, 0.0) + loss
    return {"events": [{"event_id": key, "gross_loss": value} for key, value in totals.items()], "event_count": len(totals), "total_gross_loss": sum(values)}


def apply_occurrence_layer(
    losses: Sequence[Number], event_ids: Sequence[Any], attachment: Number, limit: Number, share: Number = 1.0
) -> dict[str, Any]:
    """Aggregate claims by occurrence, then apply an occurrence layer."""
    aggregated = aggregate_occurrence_losses(losses, event_ids)
    event_losses = [row["gross_loss"] for row in aggregated["events"]]
    layer = apply_per_claim_layer(event_losses, attachment, limit, share)
    rows = []
    for event, result in zip(aggregated["events"], layer["claims"]):
        rows.append({"event_id": event["event_id"], "gross_loss": event["gross_loss"], "ceded_loss": result["ceded_total"]})
    return {"events": rows, "event_count": len(rows), "total_gross_loss": aggregated["total_gross_loss"], "total_ceded_loss": sum(x["ceded_loss"] for x in rows)}


def quota_share_ceded(values: Sequence[Number], ceded_share: Number, ceding_commission: Number | None = None) -> dict[str, Any]:
    """Apply a quota-share percentage; optionally calculate commission on ceded values."""
    gross = _numbers(values, "values")
    share = _float(ceded_share, "ceded_share")
    if not 0 <= share <= 1:
        raise ValueError("ceded_share must be in [0,1]")
    ceded = apply_factors(gross, share)
    output: dict[str, Any] = {"gross": gross, "ceded_share": share, "ceded": ceded, "retained": [g - c for g, c in zip(gross, ceded)], "total_ceded": sum(ceded)}
    if ceding_commission is not None:
        commission = _float(ceding_commission, "ceding_commission")
        if not 0 <= commission <= 1:
            raise ValueError("ceding_commission must be in [0,1]")
        output["ceding_commission_rate"] = commission
        output["ceding_commission"] = apply_factors(ceded, commission)
    return output


def limited_average_severity_empirical(losses: Sequence[Number], attachment: Number, limit: Number) -> dict[str, float]:
    """Empirical mean ground-up severity and mean severity in a specified layer."""
    values = _numbers(losses, "losses", nonnegative=True)
    if not values:
        raise ValueError("losses cannot be empty")
    layer = apply_per_claim_layer(values, attachment, limit)["claims"]
    layer_values = [row["ceded_indemnity"] for row in layer]
    return {"claim_count": float(len(values)), "mean_ground_up_severity": sum(values) / len(values), "mean_layer_severity": sum(layer_values) / len(values), "total_layer_loss": sum(layer_values)}


def empirical_layer_share(losses: Sequence[Number], attachment: Number, limit: Number) -> dict[str, float]:
    """Empirical share of gross loss dollars falling in a layer."""
    values = _numbers(losses, "losses", nonnegative=True)
    gross = sum(values)
    layer = apply_per_claim_layer(values, attachment, limit)["total_ceded_indemnity"]
    return {"gross_loss": gross, "layer_loss": layer, "layer_share": _safe_div(layer, gross, "gross_loss")}


def frequency_severity_summary(losses: Sequence[Number], exposure: Number, include_zero_losses: bool = False) -> dict[str, float]:
    """Compute claim frequency, average severity, and pure premium/loss cost."""
    values = _numbers(losses, "losses", nonnegative=True)
    exp_basis = _float(exposure, "exposure")
    if exp_basis <= 0:
        raise ValueError("exposure must be positive")
    selected = values if include_zero_losses else [value for value in values if value > 0]
    count = len(selected)
    total = sum(selected)
    return {"claim_count": float(count), "total_loss": total, "frequency": count / exp_basis, "average_severity": total / count if count else 0.0, "pure_premium": total / exp_basis}


def loss_ratio(loss: Number, earned_premium: Number) -> float:
    """Return loss divided by earned premium."""
    premium = _float(earned_premium, "earned_premium")
    if premium <= 0:
        raise ValueError("earned_premium must be positive")
    return _float(loss, "loss") / premium


def loss_ratio_series(losses: Sequence[Number], earned_premiums: Sequence[Number]) -> dict[str, Any]:
    """Return period loss ratios and the aggregate ratio of sums."""
    loss = _numbers(losses, "losses")
    premium = _numbers(earned_premiums, "earned_premiums", nonnegative=True)
    if len(loss) != len(premium):
        raise ValueError("losses and earned_premiums must have equal length")
    ratios = [loss_ratio(l, p) for l, p in zip(loss, premium)]
    return {"loss_ratios": ratios, "aggregate_loss_ratio": loss_ratio(sum(loss), sum(premium))}


def weighted_average(values: Sequence[Number], weights: Sequence[Number]) -> float:
    """Return a normalized nonnegative weighted average."""
    vals = _numbers(values, "values")
    wts = _numbers(weights, "weights", nonnegative=True)
    if len(vals) != len(wts) or not vals or sum(wts) == 0:
        raise ValueError("values/weights must match and weights must have positive total")
    return sum(v * w for v, w in zip(vals, wts)) / sum(wts)


def trimmed_mean(values: Sequence[Number], trim_each_tail: int = 1) -> float:
    """Return a mean after removing a fixed count from each tail."""
    vals = sorted(_numbers(values, "values"))
    if trim_each_tail < 0 or 2 * trim_each_tail >= len(vals):
        raise ValueError("trim_each_tail removes too many observations")
    kept = vals[trim_each_tail : len(vals) - trim_each_tail] if trim_each_tail else vals
    return sum(kept) / len(kept)


def credibility_weight(expected_claim_count: Number, full_credibility_claim_count: Number) -> float:
    """Limited-fluctuation square-root credibility weight, capped at 1."""
    count = _float(expected_claim_count, "expected_claim_count")
    full = _float(full_credibility_claim_count, "full_credibility_claim_count")
    if count < 0 or full <= 0:
        raise ValueError("expected count must be nonnegative and full-credibility count positive")
    return min(1.0, sqrt(count / full))


def credibility_blend(experience_estimate: Number, complement_estimate: Number, credibility: Number) -> float:
    """Blend experience and complement estimates using an explicit credibility weight."""
    z = _float(credibility, "credibility")
    if not 0 <= z <= 1:
        raise ValueError("credibility must be in [0,1]")
    return z * _float(experience_estimate, "experience_estimate") + (1.0 - z) * _float(complement_estimate, "complement_estimate")


def experience_rating_table(
    origin_periods: Sequence[Any],
    earned_premium: Sequence[Number],
    reported_or_paid_loss: Sequence[Number],
    on_level_factors: Number | Sequence[Number] = 1.0,
    exposure_trend_factors: Number | Sequence[Number] = 1.0,
    loss_trend_factors: Number | Sequence[Number] = 1.0,
    cumulative_factors: Number | Sequence[Number] = 1.0,
    methods: Sequence[str] = ("chain_ladder", "cape_cod"),
    apriori_loss_ratio: Number | Sequence[Number] | None = None,
    benktander_iterations: int = 2,
    cape_cod_weights: Sequence[Number] | None = None,
) -> dict[str, Any]:
    """Build a standard experience-rating table from already-selected factors.

    Supported method names: ``chain_ladder``, ``expected_loss``, ``bf``, ``cape_cod``,
    and ``benktander``. The function does not select assumptions or methods.
    """
    n = len(origin_periods)
    premium = _numbers(earned_premium, "earned_premium", nonnegative=True)
    actual = _numbers(reported_or_paid_loss, "reported_or_paid_loss", nonnegative=True)
    if len(premium) != n or len(actual) != n:
        raise ValueError("origin_periods, premium, and loss lists must match")
    olf = _broadcast(on_level_factors, n, "on_level_factors")
    exposure_trend = _broadcast(exposure_trend_factors, n, "exposure_trend_factors")
    loss_trend = _broadcast(loss_trend_factors, n, "loss_trend_factors")
    cdf = _broadcast(cumulative_factors, n, "cumulative_factors")
    adjusted_premium = apply_factors(premium, olf, exposure_trend)
    trended_loss = apply_factors(actual, loss_trend)
    rows = []
    for i, origin in enumerate(origin_periods):
        rows.append({"origin_period": origin, "nominal_earned_premium": premium[i], "on_level_factor": olf[i], "exposure_trend_factor": exposure_trend[i], "trended_on_level_earned_premium": adjusted_premium[i], "reported_or_paid_loss": actual[i], "loss_trend_factor": loss_trend[i], "trended_loss": trended_loss[i], "cdf": cdf[i], "percent_reported_or_paid": 1.0 / cdf[i], "nominal_loss_ratio": loss_ratio(actual[i], premium[i]), "trended_on_level_reported_loss_ratio": loss_ratio(trended_loss[i], adjusted_premium[i])})
    allowed = {"chain_ladder", "expected_loss", "bf", "cape_cod", "benktander"}
    requested = list(dict.fromkeys(methods))
    if any(method not in allowed for method in requested):
        raise ValueError(f"unsupported method in {requested}")
    outputs: dict[str, Any] = {}
    if "chain_ladder" in requested:
        ultimate = apply_factors(trended_loss, cdf)
        outputs["chain_ladder"] = {"ultimate": ultimate, "ultimate_loss_ratios": [loss_ratio(u, p) for u, p in zip(ultimate, adjusted_premium)], "total_ultimate": sum(ultimate)}
    if any(method in requested for method in {"expected_loss", "bf", "benktander"}) and apriori_loss_ratio is None:
        raise ValueError("apriori_loss_ratio is required for expected_loss, bf, or benktander")
    expected = expected_loss_ultimate(adjusted_premium, apriori_loss_ratio)["ultimate"] if apriori_loss_ratio is not None else None
    if "expected_loss" in requested:
        outputs["expected_loss"] = {"ultimate": expected, "ultimate_loss_ratios": [loss_ratio(u, p) for u, p in zip(expected, adjusted_premium)], "total_ultimate": sum(expected)}
    if "bf" in requested:
        outputs["bf"] = bornhuetter_ferguson_ultimate(trended_loss, cdf, adjusted_premium, apriori_loss_ratio)
        outputs["bf"]["ultimate_loss_ratios"] = [loss_ratio(u, p) for u, p in zip(outputs["bf"]["ultimate"], adjusted_premium)]
    if "cape_cod" in requested:
        outputs["cape_cod"] = cape_cod_ultimate(trended_loss, cdf, adjusted_premium, cape_cod_weights)
        outputs["cape_cod"]["ultimate_loss_ratios"] = [loss_ratio(u, p) for u, p in zip(outputs["cape_cod"]["ultimate"], adjusted_premium)]
    if "benktander" in requested:
        outputs["benktander"] = benktander_ultimate(trended_loss, cdf, expected, benktander_iterations)
        outputs["benktander"]["ultimate_loss_ratios"] = [loss_ratio(u, p) for u, p in zip(outputs["benktander"]["ultimate"], adjusted_premium)]
    return {"rows": rows, "method_results": outputs, "aggregate_nominal_loss_ratio": loss_ratio(sum(actual), sum(premium)), "aggregate_trended_on_level_reported_loss_ratio": loss_ratio(sum(trended_loss), sum(adjusted_premium))}


def method_suitability_screen(
    *,
    stable_development_pattern: bool,
    stable_case_reserve_adequacy: bool,
    stable_settlement_speed: bool,
    stable_business_mix: bool,
    credible_apriori_available: bool,
    individual_claims_available: bool,
    paid_triangle_available: bool,
    incurred_triangle_available: bool,
    granularity_matches_request: bool,
    expected_claim_count: Number,
    immature_share: Number,
) -> dict[str, Any]:
    """Create a deterministic method-screening matrix for actuarial review.

    This is a transparent screen, not an automated actuarial selection. Scores compare
    relative suitability only within this call and must be accompanied by rationale.
    """
    count = _float(expected_claim_count, "expected_claim_count")
    immature = _float(immature_share, "immature_share")
    if count < 0 or not 0 <= immature <= 1:
        raise ValueError("expected_claim_count must be nonnegative and immature_share in [0,1]")
    scores = {name: 0 for name in ["chain_ladder", "bornhuetter_ferguson", "cape_cod", "benktander", "expected_loss", "frequency_severity_layered"]}
    notes: dict[str, list[str]] = {name: [] for name in scores}
    if stable_development_pattern:
        scores["chain_ladder"] += 3
        scores["benktander"] += 1
        notes["chain_ladder"].append("development pattern appears stable")
    else:
        scores["chain_ladder"] -= 3
        notes["chain_ladder"].append("development pattern instability weakens extrapolation")
    for stable, label in [(stable_case_reserve_adequacy, "case reserve adequacy"), (stable_settlement_speed, "settlement speed")]:
        if not stable:
            scores["chain_ladder"] -= 2
            scores["cape_cod"] -= 1
            notes["chain_ladder"].append(f"changing {label} requires adjustment or reduced reliance")
    if credible_apriori_available:
        for name in ["bornhuetter_ferguson", "benktander", "expected_loss"]:
            scores[name] += 3
            notes[name].append("credible a priori estimate is available")
    else:
        for name in ["bornhuetter_ferguson", "benktander", "expected_loss"]:
            scores[name] -= 3
            notes[name].append("credible a priori estimate is unavailable")
    if stable_business_mix:
        scores["cape_cod"] += 2
        notes["cape_cod"].append("stable exposure mix supports a fitted ELR")
    else:
        scores["cape_cod"] -= 2
        scores["expected_loss"] -= 1
        notes["cape_cod"].append("business-mix change can distort a single fitted ELR")
    if immature >= 0.5:
        scores["bornhuetter_ferguson"] += 2
        scores["expected_loss"] += 1
        scores["chain_ladder"] -= 1
        notes["bornhuetter_ferguson"].append("a large immature share favors a stable prior-based method")
    if count >= 50:
        scores["chain_ladder"] += 1
        scores["cape_cod"] += 1
    elif count < 10:
        scores["chain_ladder"] -= 2
        scores["frequency_severity_layered"] -= 2
    if individual_claims_available:
        scores["frequency_severity_layered"] += 3
        notes["frequency_severity_layered"].append("claim-level data supports trending and relayering")
    else:
        scores["frequency_severity_layered"] -= 4
    if paid_triangle_available and incurred_triangle_available:
        notes["chain_ladder"].append("compare paid and incurred projections; do not blend without an explicit weight")
    elif not (paid_triangle_available or incurred_triangle_available):
        scores["chain_ladder"] -= 4
        notes["chain_ladder"].append("no loss triangle is available")
    if not granularity_matches_request:
        for name in scores:
            scores[name] -= 3
            notes[name].append("requested granularity is not supported by the available data")
    ranking = sorted([{"method": name, "score": score, "considerations": notes[name]} for name, score in scores.items()], key=lambda item: (-item["score"], item["method"]))
    return {"ranking": ranking, "screen_only": True, "required_actuarial_review": True}


PUBLIC_FUNCTIONS = [
    "compound_change_factors",
    "rate_level_at_date",
    "time_weighted_average_rate_level",
    "parallelogram_average_rate_level",
    "parallelogram_on_level_factor",
    "extension_of_exposures_on_level_factor",
    "rate_book_on_level_factor",
    "annual_rate_change_on_level_factors",
    "trend_factor",
    "apply_factors",
    "on_level_and_trend_premium",
    "trend_losses",
    "piecewise_trend_factor",
    "trend_losses_piecewise",
    "validate_triangle",
    "incremental_to_cumulative",
    "cumulative_to_incremental",
    "link_ratio_triangle",
    "development_factor_diagnostics",
    "select_age_to_age_factors",
    "cumulative_development_factors",
    "reported_percentages",
    "latest_diagonal",
    "fill_lower_triangle_chain_ladder",
    "chain_ladder_ultimate",
    "expected_loss_ultimate",
    "bornhuetter_ferguson_ultimate",
    "cape_cod_ultimate",
    "benktander_ultimate",
    "loglinear_tail_factor",
    "mack_chain_ladder",
    "compare_paid_incurred_ultimates",
    "cap_losses",
    "apply_per_claim_layer",
    "aggregate_occurrence_losses",
    "apply_occurrence_layer",
    "quota_share_ceded",
    "limited_average_severity_empirical",
    "empirical_layer_share",
    "frequency_severity_summary",
    "claim_level_xol_summary",
    "loss_ratio",
    "loss_ratio_series",
    "weighted_average",
    "trimmed_mean",
    "credibility_weight",
    "credibility_blend",
    "experience_rating_table",
    "method_suitability_screen",
]


def run_self_tests() -> dict[str, Any]:
    """Run deterministic smoke/invariant tests for every public callable function."""
    passed: list[str] = []

    def check(name: str, condition: bool) -> None:
        if not condition:
            raise AssertionError(f"self-test failed: {name}")
        passed.append(name)

    changes = [("2024-01-16", 0.10)]
    check("compound_change_factors", abs(compound_change_factors([0.10, -0.05])[-1] - 1.045) < 1e-12)
    check("rate_level_at_date", abs(rate_level_at_date(changes, "2024-01-20") - 1.1) < 1e-12)
    avg = time_weighted_average_rate_level("2024-01-01", "2024-02-01", changes)
    weighted_avg = time_weighted_average_rate_level(
        "2024-01-01", "2024-02-01", changes, daily_weights={"2024-01-20": 1.0}
    )
    check("time_weighted_average_rate_level", 1.0 < avg < 1.1 and abs(weighted_avg - 1.1) < 1e-12)
    para = parallelogram_average_rate_level("2024-01-01", "2025-01-01", [("2024-07-01", 0.10)], policy_term_days=365)
    weighted_para = parallelogram_average_rate_level(
        "2024-01-01", "2025-01-01", [("2024-07-01", 0.10)], policy_term_days=365,
        writing_weights={"2024-07-01": 1.0},
    )
    check("parallelogram_average_rate_level", 1.0 < para < 1.1 and abs(weighted_para - 1.1) < 1e-12)
    check("parallelogram_on_level_factor", parallelogram_on_level_factor("2024-01-01", "2025-01-01", "2025-01-01", [("2024-07-01", 0.10)])["on_level_factor"] > 1.0)
    check("extension_of_exposures_on_level_factor", abs(extension_of_exposures_on_level_factor([100, 100], [110, 110])["on_level_factor"] - 1.1) < 1e-12)
    check("rate_book_on_level_factor", abs(rate_book_on_level_factor([100], [1.0], 1.1)["on_level_factor"] - 1.1) < 1e-12)
    check("annual_rate_change_on_level_factors", annual_rate_change_on_level_factors([2023, 2024], [0.1, 0.1])["rows"][0]["on_level_factor"] > 1.0)
    check("trend_factor", abs(trend_factor("2023-01-01", "2024-01-01", 0.10) - 1.10) < 0.001)
    check("apply_factors", apply_factors([10, 20], [2, 3]) == [20.0, 60.0])
    check("on_level_and_trend_premium", on_level_and_trend_premium([100], 1.1, 1.2)["trended_on_level_earned_premium"] == [132.0])
    check("trend_losses", trend_losses([100], ["2023-01-01"], "2024-01-01", 0.10)["trended_losses"][0] > 109.9)
    trend_schedule = [("2023-01-01", 0.10)]
    check("piecewise_trend_factor", abs(piecewise_trend_factor("2023-01-01", "2024-01-01", trend_schedule, day_count=365) - 1.10) < 1e-12)
    piecewise = trend_losses_piecewise(
        [100, 200],
        ["2023-01-01", "2023-01-01"],
        "2024-01-01",
        {"Commercial": trend_schedule, "Personal": [("2023-01-01", 0.20)]},
        ["Commercial", "Personal"],
        day_count=365,
    )
    check("trend_losses_piecewise", piecewise["trended_losses"] == [110.00000000000001, 240.0])

    incremental = [[100, 50, 30], [110, 55, None], [120, None, None]]
    cumulative = [[100, 150, 180], [110, 165, None], [120, None, None]]
    check("validate_triangle", validate_triangle(cumulative)["valid"])
    check("incremental_to_cumulative", incremental_to_cumulative(incremental) == [[100.0, 150.0, 180.0], [110.0, 165.0, None], [120.0, None, None]])
    check("cumulative_to_incremental", cumulative_to_incremental(cumulative) == [[100.0, 50.0, 30.0], [110.0, 55.0, None], [120.0, None, None]])
    check("link_ratio_triangle", abs(link_ratio_triangle(cumulative)[0][0] - 1.5) < 1e-12)  # type: ignore[operator]
    check("development_factor_diagnostics", development_factor_diagnostics(cumulative)["development_periods"][0]["count"] == 2)
    selection = select_age_to_age_factors(cumulative)
    sparse = [[100, None, None], [110, 165, None], [120, None, None]]
    explicit_selection = select_age_to_age_factors(sparse, selected_factors=[1.5, 1.2])
    check("select_age_to_age_factors", selection["selected_factors"] == [1.5, 1.2] and explicit_selection["selected_factors"] == [1.5, 1.2])
    cdfs = cumulative_development_factors([1.5, 1.2])
    check("cumulative_development_factors", all(abs(actual - expected) < 1e-12 for actual, expected in zip(cdfs, [1.8, 1.2, 1.0])))
    check("reported_percentages", reported_percentages([2.0]) == [0.5])
    check("latest_diagonal", latest_diagonal(cumulative)[1]["development_index"] == 1)
    completed = fill_lower_triangle_chain_ladder(cumulative, [1.5, 1.2])
    check("fill_lower_triangle_chain_ladder", completed[2][-1] == 216.0)
    chain = chain_ladder_ultimate(cumulative)
    check("chain_ladder_ultimate", [round(x["ultimate"], 6) for x in chain["origins"]] == [180.0, 198.0, 216.0])
    check("expected_loss_ultimate", expected_loss_ultimate([100, 200], 0.6)["ultimate"] == [60.0, 120.0])
    bf = bornhuetter_ferguson_ultimate([50, 40], [1.25, 2.0], [100, 100], 0.6)
    check("bornhuetter_ferguson_ultimate", bf["ultimate"] == [62.0, 70.0])
    cc = cape_cod_ultimate([50, 40], [1.25, 2.0], [100, 100])
    check("cape_cod_ultimate", cc["fitted_expected_loss_ratio"] > 0)
    benk = benktander_ultimate([50], [2.0], [60], iterations=2)
    check("benktander_ultimate", benk["ultimate"] == [90.0])
    check("loglinear_tail_factor", loglinear_tail_factor([1.5, 1.25, 1.12, 1.06], future_periods=3)["tail_factor"] > 1.0)
    mack = mack_chain_ladder(cumulative)
    check("mack_chain_ladder", mack["total_mack_standard_error"] >= 0)
    paid_incurred = compare_paid_incurred_ultimates(cumulative, cumulative)
    check("compare_paid_incurred_ultimates", paid_incurred["total_selected_ultimate"] == chain["total_ultimate"])
    check("cap_losses", cap_losses([50, 150], 100) == [50.0, 100.0])
    layer = apply_per_claim_layer([50, 150, 300], 100, 100)
    check("apply_per_claim_layer", [x["ceded_total"] for x in layer["claims"]] == [0.0, 50.0, 100.0])
    xol = claim_level_xol_summary(
        ["A", "B", "C"],
        [800000, 2500000, 6000000],
        {"state": ["CA", "CA", "NY"], "line_of_business": ["Commercial"] * 3, "accident_year": [2022, 2022, 2023]},
        trend_factors=[2.0, 1.0, 1.0],
        development_factors=[1.5, 1.2, 1.1],
        attachment=1000000,
        limit=4000000,
        layer_development_order="layer_then_develop",
    )
    check(
        "claim_level_xol_summary",
        [row["nominal_reported_layer"] for row in xol["claims"]] == [0.0, 1500000.0, 4000000.0]
        and xol["totals"]["trended_reported_layer"] == 6100000.0
        and xol["totals"]["trended_ultimate_layer"] == 7100000.0
        and len(xol["groups"]) == 2,
    )
    occurrence = aggregate_occurrence_losses([40, 70, 30], ["A", "A", "B"])
    check("aggregate_occurrence_losses", occurrence["events"][0]["gross_loss"] == 110.0)
    check("apply_occurrence_layer", apply_occurrence_layer([40, 70, 30], ["A", "A", "B"], 100, 50)["total_ceded_loss"] == 10.0)
    check("quota_share_ceded", quota_share_ceded([100], 0.4)["ceded"] == [40.0])
    check("limited_average_severity_empirical", limited_average_severity_empirical([50, 150], 100, 100)["mean_layer_severity"] == 25.0)
    check("empirical_layer_share", abs(empirical_layer_share([50, 150], 100, 100)["layer_share"] - 0.25) < 1e-12)
    check("frequency_severity_summary", frequency_severity_summary([10, 20], 10)["pure_premium"] == 3.0)
    check("loss_ratio", loss_ratio(50, 100) == 0.5)
    check("loss_ratio_series", loss_ratio_series([50, 50], [100, 200])["aggregate_loss_ratio"] == 1 / 3)
    check("weighted_average", weighted_average([1, 3], [1, 1]) == 2.0)
    check("trimmed_mean", trimmed_mean([1, 2, 100]) == 2.0)
    check("credibility_weight", credibility_weight(25, 100) == 0.5)
    check("credibility_blend", credibility_blend(100, 50, 0.4) == 70.0)
    table = experience_rating_table([2023, 2024], [100, 120], [40, 30], cumulative_factors=[1.2, 2.0], methods=["chain_ladder", "bf", "cape_cod", "benktander"], apriori_loss_ratio=0.6)
    check("experience_rating_table", set(table["method_results"]) == {"chain_ladder", "bf", "cape_cod", "benktander"})
    screen = method_suitability_screen(stable_development_pattern=True, stable_case_reserve_adequacy=True, stable_settlement_speed=True, stable_business_mix=True, credible_apriori_available=True, individual_claims_available=True, paid_triangle_available=True, incurred_triangle_available=True, granularity_matches_request=True, expected_claim_count=100, immature_share=0.6)
    check("method_suitability_screen", len(screen["ranking"]) == 6)

    missing = sorted(set(PUBLIC_FUNCTIONS) - set(passed))
    unexpected = sorted(set(passed) - set(PUBLIC_FUNCTIONS))
    if missing or unexpected:
        raise AssertionError(f"self-test registry mismatch; missing={missing}, unexpected={unexpected}")
    return {"status": "passed", "public_functions_tested": len(passed), "functions": passed}


if __name__ == "__main__":
    import json

    print(json.dumps(run_self_tests(), indent=2))
