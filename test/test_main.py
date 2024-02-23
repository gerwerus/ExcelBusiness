import pytest
from excelProcessing import ExcelProccessing
import re
from getFilesData import getFilesData

files_data = getFilesData()
@pytest.fixture(params=files_data)
def instance(request):
    return ExcelProccessing(request.param)

def test_excel_instance(instance):
    validInstituteName = ["информатики и вычислительной техники", "центр подготовки научных кадров", "безопасности", "телекоммуникаций", "мобильной радиосвязи и мультимедиа"]
    assert instance.getInstituteName().lower() in validInstituteName        
    assert bool(re.search(r'\d\d.\d\d.\d\d', instance.getCode()))
    assert bool(re.search(r'\w+', instance.getProfile()))
    assert bool(re.search(r'\d{4}', instance.getDateStart()))


    