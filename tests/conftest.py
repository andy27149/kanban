import openpyxl
import pytest


@pytest.fixture
def uploads_dir(tmp_path):
    d = tmp_path / "uploads"
    d.mkdir()
    return d


@pytest.fixture
def make_workbook(uploads_dir):
    def _make(filename, sheet_data):
        """sheet_data: {sheet_name: [[row1_values], [row2_values], ...]}"""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sheet_name, rows in sheet_data.items():
            ws = wb.create_sheet(sheet_name)
            for row in rows:
                ws.append(row)
        path = uploads_dir / filename
        wb.save(path)
        return path
    return _make
