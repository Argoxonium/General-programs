from replace_URLs import edit_url, change_url

def test_change_url():
    assert edit_url("https://docs.anl.gov/main/groups/intranet/@shared/@lms/documents/procedure/lms-proc-52.pdf","https://my.anl.gov/esb/view/") == "https://my.anl.gov/esb/view/lms-proc-52"

def test_chage_urls():
    test_info:dict[str:str] = {'test words':"https://docs.anl.gov/main/groups/intranet/@shared/@lms/documents/procedure/lms-proc-52.pdf"}
    new_url:str = "https://my.anl.gov/esb/view/"
    actual_results = change_url(test_info,new_url)
    expected_results: dict[str:str] = {'test words':"https://my.anl.gov/esb/view/lms-proc-52"}
    assert actual_results == expected_results