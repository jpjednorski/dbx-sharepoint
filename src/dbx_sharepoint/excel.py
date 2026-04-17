from __future__ import annotations

import io
from typing import Optional, Union

import openpyxl
import pandas as pd
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

_VALID_ORIENTATIONS = ("rows", "columns")


def dataframe_from_excel_bytes(
    data: bytes,
    sheet_name: Optional[Union[str, int]] = None,
) -> pd.DataFrame:
    """Read Excel bytes into a pandas DataFrame.

    Args:
        data: Raw .xlsx file bytes.
        sheet_name: Sheet to read. Defaults to first sheet.

    Returns:
        DataFrame with the sheet data.
    """
    return pd.read_excel(
        io.BytesIO(data),
        engine="openpyxl",
        sheet_name=sheet_name if sheet_name is not None else 0,
    )


def dataframe_to_excel_bytes(
    df: pd.DataFrame,
    sheet_name: str = "Sheet1",
) -> bytes:
    """Write a DataFrame to .xlsx bytes.

    Args:
        df: The DataFrame to write.
        sheet_name: Name of the sheet in the output workbook.

    Returns:
        Raw .xlsx file bytes.
    """
    buf = io.BytesIO()
    df.to_excel(buf, engine="openpyxl", sheet_name=sheet_name, index=False)
    return buf.getvalue()


def _cell_to_row_col(cell_ref: str) -> tuple:
    """Convert a cell reference like 'B3' to (row, col) 1-indexed tuple."""
    try:
        col_letter, row = coordinate_from_string(cell_ref)
    except (ValueError, TypeError) as exc:
        raise ValueError(f"Invalid cell reference: '{cell_ref}'") from exc
    return row, column_index_from_string(col_letter)


class Template:
    """An Excel template that can be populated with data and saved.

    Args:
        data: Raw .xlsx file bytes of the template.
    """

    def __init__(self, data: bytes):
        self._workbook = openpyxl.load_workbook(io.BytesIO(data))

    def fill_range(
        self,
        sheet: Optional[str] = None,
        start_cell: Optional[str] = None,
        end_cell: Optional[str] = None,
        named_range: Optional[str] = None,
        data: Optional[pd.DataFrame] = None,
        orientation: str = "rows",
        allow_expand: bool = False,
    ) -> None:
        """Fill a range in the template with DataFrame data.

        Specify either (sheet + start_cell) or named_range, not both.

        Args:
            sheet: Sheet name (required when using start_cell).
            start_cell: Top-left cell to begin writing (e.g., "B3").
            end_cell: Optional bottom-right boundary. Raises if data exceeds it.
            named_range: Name of a defined range in the workbook.
            data: DataFrame to write.
            orientation: "rows" (default) writes each df row as an Excel row.
                "columns" transposes — each df row becomes an Excel column.
            allow_expand: If True, allow writing beyond a named range's bounds.
        """
        if data is None:
            raise ValueError("data is required")
        if orientation not in _VALID_ORIENTATIONS:
            raise ValueError(
                f"orientation must be one of {_VALID_ORIENTATIONS}, got '{orientation}'"
            )

        if named_range is not None:
            self._fill_named_range(named_range, data, orientation, allow_expand)
        elif sheet is not None and start_cell is not None:
            self._fill_cell_range(sheet, start_cell, end_cell, data, orientation)
        else:
            raise ValueError("Provide either (sheet + start_cell) or named_range")

    def _fill_cell_range(
        self,
        sheet: str,
        start_cell: str,
        end_cell: Optional[str],
        data: pd.DataFrame,
        orientation: str,
    ) -> None:
        ws = self._workbook[sheet]
        start_row, start_col = _cell_to_row_col(start_cell)
        end_coords = _cell_to_row_col(end_cell) if end_cell is not None else None
        self._write_block(
            ws,
            start_row,
            start_col,
            data,
            orientation,
            end_coords=end_coords,
            range_label=f"{start_cell}:{end_cell}" if end_cell else None,
            allow_expand=False,
        )

    def _fill_named_range(
        self,
        range_name: str,
        data: pd.DataFrame,
        orientation: str,
        allow_expand: bool,
    ) -> None:
        defn = self._workbook.defined_names.get(range_name)
        if defn is None:
            raise ValueError(f"Named range '{range_name}' not found in workbook")

        dest_sheet, coord_range = next(iter(defn.destinations))
        ws = self._workbook[dest_sheet]
        parts = coord_range.replace("$", "").split(":")
        start_ref = parts[0]
        end_ref = parts[1] if len(parts) > 1 else None

        start_row, start_col = _cell_to_row_col(start_ref)
        end_coords = _cell_to_row_col(end_ref) if end_ref is not None else None
        self._write_block(
            ws,
            start_row,
            start_col,
            data,
            orientation,
            end_coords=end_coords,
            range_label=f"named range '{range_name}'",
            allow_expand=allow_expand,
        )

    @staticmethod
    def _write_block(
        ws,
        start_row: int,
        start_col: int,
        data: pd.DataFrame,
        orientation: str,
        end_coords: Optional[tuple] = None,
        range_label: Optional[str] = None,
        allow_expand: bool = False,
    ) -> None:
        values = data.values
        if orientation == "columns":
            values = values.T
        num_rows, num_cols = values.shape

        if end_coords is not None and not allow_expand:
            end_row, end_col = end_coords
            max_rows = end_row - start_row + 1
            max_cols = end_col - start_col + 1
            if num_rows > max_rows or num_cols > max_cols:
                msg = (
                    f"Data ({num_rows} rows x {num_cols} cols) exceeds "
                    f"{range_label or 'range'} ({max_rows} rows x {max_cols} cols)"
                )
                if range_label and range_label.startswith("named range"):
                    msg += ". Set allow_expand=True to write beyond the range."
                raise ValueError(msg)

        for r in range(num_rows):
            for c in range(num_cols):
                ws.cell(
                    row=start_row + r,
                    column=start_col + c,
                    value=values[r][c],
                )

    def set_value(self, sheet: str, cell: str, value: object) -> None:
        """Set a single cell value in the template.

        Args:
            sheet: Sheet name.
            cell: Cell reference (e.g., "A1").
            value: Value to write.
        """
        ws = self._workbook[sheet]
        ws[cell] = value

    def to_bytes(self) -> bytes:
        """Serialize the modified template to .xlsx bytes."""
        buf = io.BytesIO()
        self._workbook.save(buf)
        return buf.getvalue()
