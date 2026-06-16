"""Pure Excel domain rules shared by tools, operations, and readers."""

from excel_tools.domain.merged_cells import build_range_read_result, fill_merged_values, merged_cells_to_map
from excel_tools.domain.ranges import format_bounds, intersection, parse_range, rects_intersect

__all__ = [
    "build_range_read_result",
    "fill_merged_values",
    "format_bounds",
    "intersection",
    "merged_cells_to_map",
    "parse_range",
    "rects_intersect",
]
