def _parse_scenario_sheet(ws) -> pd.DataFrame:
    """
    Parse ONE scenario sheet (e.g., 'Base Case', 'Breaker 1', etc.) from a
    Combined_ViolationCTG_Comparison.xlsx-style workbook.

    Supports both layouts:
      - Old: B=Contingency, C=Resulting Issue, D=Value, E=%.
      - New: B=Contingency, C=Resulting Issue, D=Limit, E=Value, F=%.

    Returns long-form rows with:
      CaseType, CTGLabel, ResultingIssue, LimViolLimit, LimViolValue, LimViolPct
    """
    records = []

    def _norm(x) -> str:
        return str(x).strip().lower()

    def _header_map_for_row(row_idx: int) -> dict:
        mapping = {}
        for col_idx in range(2, 9):  # B..H
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is None:
                continue
            t = _norm(v)

            if t in ("contingency events", "contingency", "ctglabel"):
                mapping["ctg"] = col_idx
            elif t in ("resulting issue", "limviolid", "issue"):
                mapping["issue"] = col_idx
            elif t in ("limit", "limviollimit"):
                mapping["limit"] = col_idx
            elif "value" in t:  # "Contingency Value (MVA)" or similar
                mapping["value"] = col_idx
            elif "percent" in t or t in ("limviolpct",):
                mapping["pct"] = col_idx

        # Fallbacks
        mapping.setdefault("ctg", 2)
        mapping.setdefault("issue", 3)
        mapping.setdefault("pct", 5)

        # Value fallback depends on pct position
        if "value" not in mapping:
            if mapping["pct"] == 5:
                mapping["value"] = 4  # old layout
            else:
                mapping["value"] = mapping["pct"] - 1

        return mapping

    max_row = ws.max_row or 0
    r = 1
    while r <= max_row:
        cell_b = ws.cell(row=r, column=2).value
        case_type = _canonical_case_type(cell_b)
        if not case_type:
            r += 1
            continue

        header_row = r + 1
        if header_row > max_row:
            break

        colmap = _header_map_for_row(header_row)
        r = header_row + 1

        current_issue = None
        current_limit = None

        while r <= max_row:
            next_title = ws.cell(row=r, column=2).value
            if _canonical_case_type(next_title):
                break

            ctg_val = ws.cell(row=r, column=colmap["ctg"]).value
            issue_val = ws.cell(row=r, column=colmap["issue"]).value
            limit_val = ws.cell(row=r, column=colmap.get("limit", 0)).value if "limit" in colmap else None
            value_val = ws.cell(row=r, column=colmap["value"]).value if "value" in colmap else None
            pct_val = ws.cell(row=r, column=colmap["pct"]).value

            # Skip empty
            if ctg_val is None and issue_val is None and pct_val is None and value_val is None and limit_val is None:
                r += 1
                continue

            # Carry-forward for collapsed rows
            if issue_val not in (None, ""):
                current_issue = str(issue_val).strip()
            if limit_val not in (None, ""):
                current_limit = limit_val

            records.append(
                {
                    "CaseType": case_type,
                    "CTGLabel": "" if ctg_val is None else str(ctg_val),
                    "ResultingIssue": current_issue or "",
                    "LimViolLimit": current_limit,
                    "LimViolValue": value_val,
                    "LimViolPct": _to_float(pct_val),
                }
            )

            r += 1

    df = pd.DataFrame.from_records(records)
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "CaseType",
                "CTGLabel",
                "ResultingIssue",
                "LimViolLimit",
                "LimViolValue",
                "LimViolPct",
            ]
        )
    return df

def build_case_type_comparison(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    case_type: str,
    max_rows: Optional[int] = None,
    log_func=None,
) -> pd.DataFrame:
    """
    Build a comparison DataFrame for a single case type.

    Output columns:
      CaseType, Contingency, ResultingIssue, Limit, LeftPct, RightPct, DeltaPct
    """
    if log_func:
        log_func(f"  Building case-type comparison: {case_type}")

    base_df = _load_sheet_df(workbook_path, base_sheet, log_func=log_func)
    new_df = _load_sheet_df(workbook_path, new_sheet, log_func=log_func)

    base_df = base_df[base_df["CaseType"] == case_type].copy()
    new_df = new_df[new_df["CaseType"] == case_type].copy()

    base_df["CTGLabel"] = base_df["CTGLabel"].astype(str)
    new_df["CTGLabel"] = new_df["CTGLabel"].astype(str)
    base_df["ResultingIssue"] = base_df["ResultingIssue"].astype(str)
    new_df["ResultingIssue"] = new_df["ResultingIssue"].astype(str)

    base_df = base_df.rename(
        columns={
            "CTGLabel": "Contingency",
            "LimViolPct": "LeftPct",
            "LimViolLimit": "LeftLimit",
        }
    )
    new_df = new_df.rename(
        columns={
            "CTGLabel": "Contingency",
            "LimViolPct": "RightPct",
            "LimViolLimit": "RightLimit",
        }
    )

    df = pd.merge(
        base_df[["CaseType", "Contingency", "ResultingIssue", "LeftPct", "LeftLimit"]],
        new_df[["CaseType", "Contingency", "ResultingIssue", "RightPct", "RightLimit"]],
        on=["CaseType", "Contingency", "ResultingIssue"],
        how="outer",
    )

    # Prefer RightLimit when available, else LeftLimit
    df["Limit"] = df["RightLimit"].combine_first(df["LeftLimit"])
    df["DeltaPct"] = df["RightPct"] - df["LeftPct"]

    df = df[
        ["CaseType", "Contingency", "ResultingIssue", "Limit", "LeftPct", "RightPct", "DeltaPct"]
    ]

    if max_rows is not None and max_rows > 0:
        df = df.head(max_rows)

    return df

def build_pair_comparison_df(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func=None,
) -> pd.DataFrame:
    """
    Output columns:
      CaseType, Contingency, ResultingIssue, Limit, LeftPct, RightPct, DeltaDisplay
    """
    records = []

    for display_name, canonical_case in CANONICAL_CASE_TYPES.items():
        df_case = build_case_type_comparison(
            workbook_path,
            base_sheet=left_sheet,
            new_sheet=right_sheet,
            case_type=canonical_case,
            max_rows=None,
            log_func=log_func,
        )

        if df_case.empty:
            continue

        for _, row in df_case.iterrows():
            cont = str(row.get("Contingency", "") or "")
            issue = str(row.get("ResultingIssue", "") or "")
            limit = row.get("Limit", None)

            left_pct = row.get("LeftPct", math.nan)
            right_pct = row.get("RightPct", math.nan)
            delta_pct = row.get("DeltaPct", math.nan)

            vals = []
            if not _is_nan(left_pct):
                vals.append(left_pct)
            if not _is_nan(right_pct):
                vals.append(right_pct)
            if not vals:
                continue
            if max(vals) < threshold:
                continue

            if _is_nan(left_pct) and not _is_nan(right_pct):
                delta_text = "Only in right"
            elif not _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = "Only in left"
            elif _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = ""
            else:
                delta_text = f"{float(delta_pct):.2f}" if not _is_nan(delta_pct) else ""

            records.append(
                {
                    "CaseType": canonical_case,
                    "Contingency": cont,
                    "ResultingIssue": issue,
                    "Limit": limit,
                    "LeftPct": float(left_pct) if not _is_nan(left_pct) else None,
                    "RightPct": float(right_pct) if not _is_nan(right_pct) else None,
                    "DeltaDisplay": delta_text,
                }
            )

    df_all = pd.DataFrame.from_records(records)

    if not df_all.empty:
        sort_vals = df_all[["LeftPct", "RightPct"]].max(axis=1)
        df_all["_SortKey"] = sort_vals
        df_all = df_all.sort_values(
            by=["CaseType", "_SortKey"], ascending=[True, False], na_position="last"
        ).drop(columns=["_SortKey"])

    return df_all