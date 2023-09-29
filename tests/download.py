from pbi.pbi import PowerBI
from tests.config import config

test_pbi = PowerBI(config)

report = test_pbi.download(
    'test_report',
    test_pbi.download_folder
)
