import re

from conftest import (
    instance,
    valid_institute_name as validInstituteName,
    valid_view_name as validViewName,
)


def test_excel_instance(instance, validInstituteName, validViewName):
    assert instance.getInstituteName().lower() in validInstituteName
    assert bool(re.search(r"\d+.\d+.\d+", instance.getCode()))
    assert bool(re.search(r"\w+", instance.getProfile()))
    assert bool(re.search(r"\d{4}", instance.getDateStart()))
    view = instance.getView()
    for practiceSem, practiceView, practiceTyp, laborTime, practiceComp in view:
        assert bool(re.search(r"\d{1,2}", str(practiceSem)))
        assert practiceView.lower() in validViewName
        assert bool(re.search(r"\w+", practiceTyp))
        assert bool(re.fullmatch(r"\d{2,4}/\d{1,3}", laborTime))
        for comp in practiceComp:
            assert bool(re.search(r"[А-Я]+К-\d+", comp))
    dates = instance.getPracticeDates()
    for date in dates:
        for type_ in dates[date]:
            if dates[date][type_]:
                for qdate in dates[date][type_]:
                    assert qdate.year() == int(date) or qdate.year() == int(date) + 1
                    assert 1 <= qdate.month() <= 12
                    assert 1 <= qdate.day() <= 31
