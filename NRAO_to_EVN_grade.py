#!/usr/bin/env python3
"""Convert a 0-10 grade (10 = highest priority) to the 0.5-2.5 priority scale
(0.5 = highest priority, 2.5 = lowest priority)."""

import sys


def grade_to_priority(grade):
    """Map grade in [0, 10] to priority in [0.5, 2.5]."""
    if not 0 <= grade <= 10:
        raise ValueError("grade must be between 0 and 10")
    return 2.5 - grade / 5.0


def main():
    if len(sys.argv) > 1:
        grade = float(sys.argv[1])
    else:
        grade = float(input("Enter grade (0-10): "))
    print(f"{grade_to_priority(grade):.2f}")


if __name__ == "__main__":
    main()
