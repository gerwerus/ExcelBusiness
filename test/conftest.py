import pytest

from excelProcessing import ExcelProccessing
from getFilesData import getFilesData


files_data = getFilesData()


@pytest.fixture(params=files_data)
def instance(request):
    excel = ExcelProccessing(request.param)
    yield excel
    excel.book.close()


@pytest.fixture()
def valid_institute_name():
    validInstituteName = [
        "информатики и вычислительной техники",
        "центр подготовки научных кадров",
        "безопасности",
        "телекоммуникаций",
        "мобильной радиосвязи и мультимедиа",
    ]
    return validInstituteName


@pytest.fixture()
def valid_view_name():
    validViewName = [
        "учебная практика",
        "производственная практика",
        "преддипломная практика",
        "научно-исследовательская работа",
    ]
    return validViewName
