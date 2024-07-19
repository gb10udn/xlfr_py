import pylightxl as xl  # type: ignore


def parse(
        path: str, sheet_name: str='Sheet1', *,
        start_row_idx: int | None=None, end_row_idx: int | None=None, start_col_idx: int | None=None, end_col_idx: int | None=None,
    ) -> list[list[str]]:
    """
    
    """
    db = xl.readxl(path)

    results = []
    for row_idx, row in enumerate(db.ws(ws=sheet_name).rows):
        row_conditions = [
            (start_row_idx is None) or ((start_row_idx is not None) and (row_idx >= start_row_idx)),
            (end_row_idx is None) or ((end_row_idx is not None) and (row_idx <= end_row_idx)),
        ]
        if all(row_conditions):
            result = []
            for col_idx, data in enumerate(row):
                col_conditions = [
                    (start_col_idx is None) or ((start_col_idx is not None) and (col_idx >= start_col_idx)),
                    (end_col_idx is None) or ((end_col_idx is not None) and (col_idx <= end_col_idx)),
                ]
                if all(col_conditions):
                    result.append(str(data))
            
            results.append(result)

    return results

    
# if __name__ == '__main__':
#     result = parse('test.xlsm', start_row_idx=2, start_col_idx=0, end_row_idx=3)
#     print(result)