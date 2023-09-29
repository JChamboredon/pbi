from pbi.pbi import PowerBI
from tests.config import config

test_pbi = PowerBI(config)

test_pbi.upload(
    test_pbi.upload_folder
)
