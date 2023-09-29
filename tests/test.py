from pbi.pbi import PowerBI
from pbi.report import PbiReport
from tests.config import config

test_pbi = PowerBI(config)

test_pbi.download(
    'test_report',
    test_pbi.download_folder
)

initial_report = PbiReport(
    test_pbi.download_folder,
    'test_report'
)

report = initial_report.copy(
    new_name='report',
    new_folder=test_pbi.upload_folder
)

home_page = report.get_page('Home')
page1 = report.get_page('Page 1')

header = home_page.get_visual_group('Header')
added_header = page1.add_visuals(header)

buttons = home_page.get_visuals('Button')
added_buttons = page1.add_visuals(buttons)

future_button = home_page.get_visuals('Future Link')[0]
future_button.hide()

wip_button = home_page.get_visuals('WIP Button')[0]
wip_button.remove_link(
    update_style=False,
    hide=False
)

title = home_page.get_visuals('Awesome title')[0]
title['config']['singleVisual']['objects']['general'][0]['properties']['paragraphs'][0]['textRuns'][0]['textStyle']['fontSize'] = '20pt'

report.update_multiselect()
report.add_search()
report.update_keep_layer_order()
report.disable_headers(
    types_to_filter=['shape', 'textbox']
)
report.save()

test_pbi.upload(
    test_pbi.upload_folder
)
