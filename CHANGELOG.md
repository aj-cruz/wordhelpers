# 0.1.4 (1 February, 2026)
- ```inject_table()``` now accepts type ```WordTableModel``` in addition to ```dict``` so that we don't have to re-calculate the pydantic model if ```WordTableModel``` is supplied

# 0.1.3 (31, January, 2026)
- Added class methods to ```WordTableModel``` for building tables directly in the pydantic model
    - add_row(width: int, text: list[str] = [], merge_cols: list[int] = [], background_color: str | None = None, style: str | None = None, alignment: AlignmentEnum | None = None)
    - add_text_to_row(row_index: int, text: list[str], style: str | None = None, alignment: AlignmentEnum | None = None)
    - add_text_to_cell(row_index: int, col_index: int, text: str, style: str | None = None, alignment: AlignmentEnum | None = None)
    - style_row(row_index: int, text_style: str)
    - style_cell(row_index: int, col_index: int, text_style: str)
    - color_row(row_index: int, background_color: str)
    - color_cell(row_index: int, col_index: int, background_color: str)
    - align_row(row_index: int, alignment: AlignmentEnum)
    - align_cell(row_index: int, col_index: int, alignment: AlignmentEnum)
    - add_table_to_cell(row_index: int, col_index: int, table: WordTableModel)
    - delete_row(row_index: int)
    - model_dump()
    - pretty_print()
    - write()

# 0.1.2 (30 January, 2026)
- Attempting to fix Pypi.org version resolution

# 0.0.1 (28 January, 2026)
- Initial release
- Sort of, this is a complete division & refactor of an older msoffice project