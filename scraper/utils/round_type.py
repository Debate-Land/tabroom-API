from enum import Enum


class RoundType(Enum):
    """Round types."""

    PRELIM = 'Prelim'
    ELIM = 'Elim'


def get_round_type(name: str) -> RoundType:
    """Identifies the type of round given its name.

    Args:
        name (str): The round name.

    Raises:
        Exception: Thrown round type is not in scope of the event (breakout Novice/JV rounds).

    Returns:
        RoundType: The type of round identified.
    """

    name = name.lower()

    if name[0] == "n" and not name.startswith("not"):
        raise Exception(
            f"Classified {name} as a novice elimination breakout round.")

    if (name[0] == 'r' and name[1] != 'u') or 'round' in name or (any(c.isdigit() for c in name) and 'x' not in name and 'elim' not in name):
        return "Prelim"

    else:
        return "Elim"
