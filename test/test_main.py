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
    validViewName = ["учебная практика", "производственная практика", "преддипломная практика", "научно-исследовательская работа"]
    assert instance.getInstituteName().lower() in validInstituteName        
    assert bool(re.search(r'\d+.\d+.\d+', instance.getCode()))
    assert bool(re.search(r'\w+', instance.getProfile()))
    assert bool(re.search(r'\d{4}', instance.getDateStart()))
    view = instance.getView()
    for practiceSem, practiceView, practiceTyp, laborTime, practiceComp in view:
        assert bool(re.search(r'\d{1,2}', str(practiceSem)))
        assert practiceView.lower() in validViewName
        assert bool(re.search(r'\w+', practiceTyp))
        assert bool(re.fullmatch(r'\d{2,4}/\d{1,3}', laborTime))
        for comp in practiceComp:
            assert  bool(re.search(r'[А-Я]+К-\d+', comp))
    dates = instance.getPracticeDates()
    for date in dates:
        for type_ in dates[date]:
            if (dates[date][type_]):
                for qdate in dates[date][type_]:
                    assert qdate.year() == int(date)
                    assert 1 <= qdate.month() <= 12
                    assert 1 <= qdate.day() <= 31

            
        



    