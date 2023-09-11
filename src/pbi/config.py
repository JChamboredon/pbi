import copy
import json

from src.pbi.object import _PbiObject


class PbiConfig(dict):
    """
    The Power BI Config class.
    """
    def __init__(self, strg):
        """
        Creates a PbiConfig object with all appropriate json strings converted as dictionaries or similar objects
        :param strg: list of arguments
        """
        super().__init__(json.loads(strg))

    def export(self):
        """
        Converts appropriate dictionaries to json strings
        :return: a string
        """
        config = copy.deepcopy(self)
        return json.dumps(config)


class _PbiConfigObject(_PbiObject):
    """
    A class for all object having a Power BI config.
    """
    def format_config(self):
        """
        Formats the config to a Power BI config Object
        :return: None
        """
        self['config'] = PbiConfig(self['config'])

    def export_config(self):
        """
        Exports the config to a list of json strings
        :return: a list
        """
        return self['config'].export()
