from core.validate import safe_eval


def test_safe_eval_formula() -> None:
    values = {
        "carry_over_cases": 10,
        "initiated_cases": 5,
        "received_from_other_organs": 2,
        "separated_cases": 1,
        "returned_for_fixing_gaps": 1,
        "merged_cases": 1,
        "sent_to_prosecutor_256": 3,
        "sent_to_prosecutor_487": 1,
        "terminated_cases": 2,
        "suspended_cases": 1,
    }
    expr = (
        "carry_over_cases + initiated_cases + received_from_other_organs + separated_cases + "
        "returned_for_fixing_gaps + merged_cases - (sent_to_prosecutor_256 + sent_to_prosecutor_487 + "
        "terminated_cases + suspended_cases)"
    )
    assert safe_eval(expr, values) == 13
