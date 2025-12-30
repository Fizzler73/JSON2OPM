from typing import Any, Dict, Set, Tuple


def _type_name(value: Any) -> str:
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return "array"
    if value is None:
        return "null"
    return type(value).__name__


def diff_schemas(
    source: Dict[str, Any],
    target: Dict[str, Any],
    path: str = ""
) -> Tuple[Set[str], Set[str], Set[str]]:
    """
    Compare JSON schemas structurally.

    Returns:
        missing_in_source: keys present in target but missing in source
        extra_in_source: keys present in source but missing in target
        type_mismatches: keys with same name/path but different types
    """
    missing_in_source: Set[str] = set()
    extra_in_source: Set[str] = set()
    type_mismatches: Set[str] = set()

    source_keys = set(source.keys())
    target_keys = set(target.keys())

    for key in target_keys - source_keys:
        missing_in_source.add(f"{path}/{key}")

    for key in source_keys - target_keys:
        extra_in_source.add(f"{path}/{key}")

    for key in source_keys & target_keys:
        src_val = source[key]
        tgt_val = target[key]

        src_type = _type_name(src_val)
        tgt_type = _type_name(tgt_val)

        current_path = f"{path}/{key}"

        if src_type != tgt_type:
            type_mismatches.add(
                f"{current_path} : {src_type} → {tgt_type}"
            )
            continue

        if isinstance(src_val, dict) and isinstance(tgt_val, dict):
            m, e, t = diff_schemas(src_val, tgt_val, current_path)
            missing_in_source |= m
            extra_in_source |= e
            type_mismatches |= t

    return missing_in_source, extra_in_source, type_mismatches
