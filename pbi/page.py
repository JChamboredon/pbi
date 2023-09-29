import copy
import json

from pbi.config import PbiConfig
from pbi.container import PbiContainer
from pbi.filter import _PbiFilterObject


class PbiPage(dict, _PbiFilterObject):
    """
    The Power BI page (or section) class.
    """
    name_prefix = 'ReportSection'

    def __init__(self, *arg, **kw):
        """
        Creates a PbiPage object with all appropriate json strings converted as dictionaries or similar objects
        :param arg: list of arguments
        :param kw: dictionary of keywords
        """
        super().__init__(*arg, **kw)
        for i, container in enumerate(self['visualContainers']):
            self['visualContainers'][i] = PbiContainer(self['visualContainers'][i])
        self['config'] = PbiConfig(self['config'])
        self.format_filter()

    @property
    def display_name(self):
        """
        Returns the display name of the page
        :return: a string
        """
        return self['displayName']

    @property
    def name(self):
        """
        Returns the name of the page
        :return: a string
        """
        return self['name']

    def export(self):
        """
        Converts appropriate dictionaries to json strings
        :return: a dictionary
        """
        page = copy.deepcopy(self)
        for i, container in enumerate(page['visualContainers']):
            page['visualContainers'][i] = container.export()
        page['config'] = page['config'].export()
        page['filters'] = self.export_filters()
        return page

    def _get_visuals_from_name_set(self, name_set):
        """
        Gets all visuals which contains the names in the given set of names
        :param name_set: a set
        :return: a lit of visuals
        """
        return [
            container for container in self['visualContainers']
            if {name for name in name_set if name in json.dumps(container)}
        ]

    def get_visuals(self, vis_name=None):
        """
        returns the list of Power BI containers which name includes a given string
        :param vis_name: a string or None
        :return: a list of PbiContainers
        """
        if vis_name is None:
            name_set = set()
        else:
            name_set = {vis_name}
        return self._get_visuals_from_name_set(name_set)

    def get_visual_group(self, group_name=None):
        """
        returns all the visuals belonging to the group with given name
        :param group_name: a string or None
        :return: a list of visuals
        """
        found = self.get_visuals(group_name)
        previously_found = []
        while len(previously_found) < len(found):
            previously_found = found
            found = self._get_visuals_from_name_set({vis.name for vis in found})
        return found

    def remove_visuals(self, vis_name=None):
        """
        removes the Power BI containers which name include a given string
        :param vis_name: a string, a list of strings, or None
        :return: None
        """
        if not vis_name:
            return None
        if type(vis_name) == list:
            self.remove_visuals(vis_name[0])
            self.remove_visuals(vis_name[1:])
        else:
            self['visualContainers'] = [container for container in self['visualContainers']
                                        if vis_name not in json.dumps(container)]

    def add_visuals(self, visual_lst=None):
        """
        Adds the given visuals to the page layout.
        :param visual_lst: a list of Power BI containers or a Power BI container or None
        :return: the list of added visuals (with updated names)
        """
        if visual_lst is None:
            visual_lst = []
        if type(visual_lst) != list:
            visual_lst = [visual_lst]
        new_list = [visual.copy() for visual in visual_lst]
        for visual in new_list:
            visual.reset_name()
        for i, vis in enumerate(visual_lst):
            for j, vis2 in enumerate(visual_lst):
                if vis2.parent_name == vis.name:
                    new_list[j].update_parent_name(new_list[i].name)
        self['visualContainers'] += new_list
        return new_list

    def add_bookmarks(self, report, page_name, bookmark_names):
        """
        Adds all the visuals from the given report that are governed by the given bookmarks.
        :param report: a Power BI report
        :param page_name: the name of a page within the given report where to get the visuals to copy
        :param bookmark_names: a string, or a list of bookmark names
        :return: the list of added visuals (with updated names)
        """
        if type(bookmark_names) == str:
            return self.add_bookmarks(report, page_name, [bookmark_names])
        original_page = report.get_page(page_name)
        bookmark_list = report.get_bookmarks(bookmark_names)
        visual_names_in_bookmarks = set()
        for bookmark in bookmark_list:
            visual_names_in_bookmarks.union(set(bookmark.get_target_visuals()))
        visuals_to_copy = original_page._get_visuals_from_name_set(visual_names_in_bookmarks)

    def replace_visual_by_placeholder(self, visual, placeholder_visual):
        """
        Hides the visual and places a given placeholder visual on top
        :param visual: on visual on the page
        :param placeholder_visual: a visual
        :return: None
        """
        visual_name = visual.name
        current_visual = self._get_visuals_from_name_set({visual_name})[0]
        v_x = current_visual['x']
        v_y = current_visual['y']
        v_z = current_visual['z']
        v_width = current_visual['width']
        v_height = current_visual['height']
        current_visual.hide()

        added_visual_name = self.add_visuals(placeholder_visual)[0].name
        added_visual = self._get_visuals_from_name_set({added_visual_name})[0]
        nv_width = added_visual['width']
        nv_height = added_visual['height']

        added_visual.update_position(
            x=v_x + (v_width - nv_width) / 2,
            y=v_y + (v_height - nv_height) / 2,
            z=v_z + 1
        )
        added_visual.update_parent_name(visual.parent_name)

    def update_keep_layer_order(self):
        """
        Updates all single visuals to keep layer order and returns number of updates
        :return: the number of updates
        """
        res = 0
        for vis in self['visualContainers']:
            res += vis.update_keep_layer_order()
        return res

    def update_multiselect(self):
        """
        Updates all multiselect slicers on the page not to allow selection with CTRL key and returns number of updates
        :return: the number of updates
        """
        res = 0, 0, 0
        for vis in self['visualContainers']:
            if vis.type == 'slicer':
                res = tuple(map(sum, zip(res, vis.update_multiselect())))
        return res

    def disable_headers(self, types_to_filter=None):
        """
        Disables all headers on the page returns number of updates
        :types_to_filter: a list of Power BI visual types or None
        :return: the number of updates
        """
        res = 0
        for vis in self['visualContainers']:
            if types_to_filter is None or vis.type in types_to_filter:
                res += vis.disable_headers()
        return res

    def add_search(self):
        """
        Add search feature on all slicers on the page and returns number of updates
        :return: the number of updates
        """
        res = 0
        for vis in self['visualContainers']:
            if vis.type == 'slicer':
                res += vis.add_search()
        return res

    def remove_visuals_from_mobile(self):
        """
        Removes all the visuals from the mobile screen and returns number of updates
        :return: the number of updates
        """
        res = 0
        for vis in self['visualContainers']:
            res += vis.remove_from_mobile()
        return res

    def get_used_bookmark_names(self):
        """
        returns the set of bookmarks used in the visuals of the page
        :return: A set of bookmark names
        """
        res = set()
        for vis in self['visualContainers']:
            try:
                res.add(vis['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['bookmark']['expr']['Literal']['Value'])
            except (KeyError, IndexError):
                pass
        res = {name.replace("'", "") for name in res}
        return res

    def replace(self, str, new_str):
        """
        replaces every occurrence of a string with a new string and returns the number of updates
        :param str: a string
        :param new_str: a string
        :return: an integer
        """
        count = 0
        for vis in self['visualContainers']:
            try:
                if str in \
                        vis['config']['singleVisual']['objects']['text'][1]['properties']['text']['expr']['Literal'][
                            'Value']:
                    vis['config']['singleVisual']['objects']['text'][1]['properties']['text']['expr']['Literal'][
                        'Value'] = \
                    vis['config']['singleVisual']['objects']['text'][1]['properties']['text']['expr']['Literal'][
                        'Value'].replace(str, new_str)
                    count += 1
            except (KeyError, IndexError):
                pass
            try:
                if str in \
                        vis['config']['singleVisual']['vcObjects']['title'][0]['properties']['text']['expr']['Literal'][
                            'Value']:
                    vis['config']['singleVisual']['vcObjects']['title'][0]['properties']['text']['expr']['Literal'][
                        'Value'] = \
                    vis['config']['singleVisual']['vcObjects']['title'][0]['properties']['text']['expr']['Literal'][
                        'Value'].replace(str, new_str)
                    count += 1
            except (KeyError, IndexError):
                pass
        return count

    def set_visibility(self, value):
        """
        Sets the visibility of the page to the provided value
        :param value: 0 or 1
        :return: None
        """
        self['config']['visibility'] = value

    def hide(self):
        """
        Hides the page.
        :return: None
        """
        self.set_visibility(1)

    def unhide(self):
        """
        Unhides the page.
        :return: None
        """
        self.set_visibility(0)
