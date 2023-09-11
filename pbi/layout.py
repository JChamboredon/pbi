import copy
import json

from pbi.bookmark import Bookmark
from pbi.config import _PbiConfigObject
from pbi.filter import _PbiFilterObject
from pbi.page import PbiPage


class PbiLayout(dict, _PbiFilterObject, _PbiConfigObject):
    """
    The Power BI Layout class.
    """
    def __init__(self, strg):
        """
        Creates a PbiLayout object with all appropriate json strings converted as dictionaries or similar objects
        :param strg: a json string that represents a Power BI layout
        """
        super().__init__(json.loads(strg))
        self.format_filter()
        self.format_config()
        try:
            self['config']['bookmarks'] = [Bookmark(item) for item in self['config']['bookmarks']]
        except KeyError:
            print("""
Warning: no bookmarks found in layout. Ignoring.
""")
        self._update_layout_objects()

    @property
    def pages(self):
        """
        Returns the list of page names in the layout
        :return:
        """
        return [section['displayName'] for section in self['sections']]

    def get_bookmarks(self, names=None):
        """
        Returns the bookmarks (with all their information) present at layout (report) level
        :param names: a string, a list of strings, or None
        :return: a list of dictionaries (representing bookmarks)
        """
        if type(names) == str:
            return self.get_bookmarks([names])
        return [bookmark for bookmark in self['config']['bookmarks'] if names is None or bookmark['displayName'] in names]

    def get_used_bookmark_names(self):
        """
        Collects the bookmarks that are effectively used in the design visuals
        :return: a set of bookmark names
        """
        res = set()
        for page in self['sections']:
            res = res.union(page.get_used_bookmark_names())
        return res

    def _update_layout_objects(self):
        """
        Updates layout strings into lists or dictionaries
        :return: None
        """
        for i, section in enumerate(self['sections']):
            self['sections'][i] = PbiPage(section)

    def replace(self, target, replacement):
        """
        Replace a string by another in all the layout.
        NB: Use with caution as the target string might appear in some unsuspected places
        NB: some strings appearing in filters, etc. need to have the single quote duplicated before running this
        :param target: the string to be replaced
        :param replacement: the replacement string
        :return: None
        """
        str_layout = self.export()
        self.__init__(str_layout.replace(target, replacement))

    def get_page(self, name=None):
        """
        returns the page of the report with the given name
        :param name: a string
        :return: a Power BI page
        """
        if name is None:
            name = ''
        page_list = [page for page in self['sections'] if page['displayName'].upper() == name.upper()]
        try:
            assert len(page_list) == 1
        except AssertionError:
            print(f'Page not found or several pages found: {name}')
            return None
        return page_list[0]

    def get_resource_package(self, name):
        """
        Returns the layout resource package of given name
        :param name: the resource package name (e.g. 'RegisteredResources', 'SharedResources')
        :return: the resource package itself (dictionary)
        """
        lst = self['resourcePackages']
        res_lst = [
            package['resourcePackage'] for package in lst
            if package['resourcePackage']['name'] == name
        ]
        assert len(res_lst) == 1
        return res_lst[0]

    def add_resource_packages(self, name, resource_package_item):
        """
        Adds a resource package item to the layout resource packages of given name
        :param name: the resource package name (e.g. 'RegisteredResources', 'SharedResources')
        :param resource_package_item: the resource package item (dictionary)
        :return: None
        """
        self.get_resource_package(name)['items'].append(resource_package_item)

    def export(self):
        """
        Converts appropriate dictionaries back to json strings
        :return: a json string
        """
        layout = copy.deepcopy(self)
        try:
            layout['filters'] = self.export_filters()
        except KeyError:
            pass
        layout['config'] = self.export_config()
        try:
            self['config']['bookmarks'] = [dict(bookmark) for bookmark in self['config']['bookmarks']]
        except KeyError:
            pass
        for i, section in enumerate(self['sections']):
            layout['sections'][i] = section.export()
        return json.dumps(layout)
