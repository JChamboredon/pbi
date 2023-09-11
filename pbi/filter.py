import copy
import json

from pbi.object import _PbiObject


class PbiFilter(dict, _PbiObject):
    name_prefix = 'Filter'

    def __init__(self, *arg, **kw):
        """
        Creates a PbiFilter object with all appropriate json strings converted as dictionaries or similar objects
        :param arg: list of arguments
        :param kw: dictionary of key words
        """
        super().__init__(*arg, **kw)
        
    def copy(self):
        """
        Returns a new Power BI Filter object
        :return: A Power BI  Filter
        """
        return PbiFilter(copy.deepcopy(self))

    def update_name(self, new_name):
        """
        Updates the name of the current Power BI Filter
        :param new_name: a string
        :return: None
        """
        self['name'] = new_name

    def update_value(self, value):
        """
        Updates the value of the filter
        :param value: a string
        :return: None
        """
        corrected_value = value.replace("'", "''")
        try:
            self['filter']['Where'][0]['Condition']['In']['Values'][0][0]['Literal']['Value'] = f"'{corrected_value}'"
        except KeyError:
            try:
                self['filter']['Where'][0]['Condition']['Comparison']['Right']['Literal']['Value'] = f"'{corrected_value}'"
            except KeyError:
                print(f'Unable to update filter to {value}.')

    def export(self):
        """
        Converts appropriately
        :return: a json string
        """
        return self


class PbiFilters(list):

    @classmethod
    def get_from_strg(cls, strg):
        """
        Creates a list of PbiFilter objects with all appropriate json strings converted as dictionaries or similar objects
        :param strg: a json string that represents a list of Power BI filters
        :return: A Power BI Filters Object
        """
        res = PbiFilters(json.loads(strg))
        for i, filter in enumerate(res):
            res[i] = PbiFilter(res[i])
        return res

    def get_filters(self, filter_name):
        """
        Returns the filter with the name given
        :param filter_name: a string
        :return: a list of Power Bi Filter Object
        """
        return [
            filter for filter in self
            if 'Column' not in filter['expression']
            or filter['expression']['Column']['Property'] == filter_name
        ]

    def export(self):
        """
        Converts appropriate dictionaries to json strings
        :return: a json string
        """
        lst = self.copy()
        for i, filter in enumerate(lst):
            lst[i] = filter.export()
        return json.dumps(lst)


class _PbiFilterObject(_PbiObject):
    """
    A class for all object having a Power BI filter.
    """
    def format_filter(self):
        """
        Formats the filter to a Power BI filter Object
        :return: None
        """
        try:
            self['filters'] = PbiFilters.get_from_strg(self['filters'])
        except KeyError:
            pass

    def export_filters(self):
        """
        Exports the filter to a list of json strings
        :return: a list
        """
        return self['filters'].export()

    def get_filters(self, filter_name):
        """
        Returns the filter with the name given
        :param filter_name: a string
        :return: a list of Power Bi Filter Object
        """
        return self['filters'].get_filters(filter_name)

    def update_filter(self, filter_name, value):
        """
        Updates a filter to a new value
        :param filter_name: a string
        :param value: a string
        :return: None
        """
        filter = self.get_filters(filter_name)
        filter.update_value(value)

    def add_filters(self, filters):
        """
        Adds the given filters to the page.
        :param filters: a Power BI Filters Object
        :return: None
        """
        new_list = [filter.copy() for filter in filters]
        for filter in new_list:
            filter.update_name(self._generate_name())
        self['filters'] = PbiFilters(self['filters'] + new_list)
