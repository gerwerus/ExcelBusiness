import pytest
from excelProcessing import ExcelProccessing
import os

files_data = [
    # ("Аспирантура", os.walk("../plans/Аспирантура/")),
    # ("ДОТ", os.walk("../plans/ДОТ/")),
    # ("ЗО", os.walk("../plans/ЗО/")),
    ("ИБ", os.walk("../plans/ИБ/")),
    # ("ИИВТ", os.walk("../plans/ИИВТ/")),
    # ("ИТ", os.walk("../plans/ИТ/")),
]

@pytest.fixture
def getInstituteName():
    return 'безопасности'
@pytest.mark.usefixtures("getInstituteName")
class TestExcelProccessing:
    validInstituteName = ["информатики и вычислительной техники", "центр подготовки научных кадров", "безопасности", "телекоммуникаций"]
    @pytest.mark.parametrize("name, walk", files_data)
    def test_institute_name(self, name, walk, getInstituteName: str):
        print(name)
        for top, dir, files in walk:
            print(top)
            assert all(ExcelProccessing(os.path.join(top, file)).getInstituteName().lower() in self.validInstituteName for file in files)
            print("OK")
    