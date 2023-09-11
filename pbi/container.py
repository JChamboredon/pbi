import copy
import json

from pbi.config import PbiConfig
from pbi.filter import _PbiFilterObject


class PbiContainer(dict, _PbiFilterObject):
    """
    The PowerBI Container (or visual) class.
    """

    def __init__(self, *arg, **kw):
        """
        Creates a PbiContainer object with all appropriate json strings converted as dictionaries or similar objects
        :param arg: list of arguments
        :param kw: dictionary of key words
        """
        super().__init__(*arg, **kw)
        self['config'] = PbiConfig(self['config'])
        self.format_filter()
        try:
            self['query'] = json.loads(self['query'])
        except KeyError:
            pass
        try:
            self['dataTransforms'] = json.loads(self['dataTransforms'])
        except KeyError:
            pass

    @property
    def name(self):
        """
        returns the name of the PowerBI container
        :return: a string corresponding to a hexadecimal
        """
        return self['config']['name']

    def update_name(self, new_name):
        """
        Updates the name of the PowerBI container
        :param new_name: a string corresponding to a hexadecimal
        :return: None
        """
        self['config']['name'] = new_name

    def reset_name(self):
        """
        Resets the visual name to a random name
        :return: None
        """
        self.update_name(self._generate_name())

    @property
    def parent_name(self):
        """
        return the visual parent Name if it exsits
        :return: a visual name
        """
        try:
            return self['config']['parentGroupName']
        except KeyError:
            return None

    def update_parent_name(self, new_name):
        """
        Updates the parent name of the PowerBI container
        :param new_name: a string corresponding to a hexadecimal
        :return: None
        """
        self['config']['parentGroupName'] = new_name

    @property
    def is_group(self):
        """
        Returns true if the containers is a group, false if not
        :return: A boolean
        """
        return 'singleVisualGroup' in self['config'].keys()

    @property
    def display_name(self):
        """
        return the display name of the Power BI container
        :return: a string
        """
        try:
            if self.is_group:
                return self['config']['singleVisualGroup']['displayName']
            else:
                return self['config']['singleVisual']['vcObjects']['title'][0]['properties']['text']['expr']['Literal'][
                    'Value']
        except KeyError:
            return f'Unavailable (Unknown {self.type})'

    @property
    def type(self):
        """
        Returns the single visual type or visualGroup
        :return: a string
        """
        if self.is_group:
            return 'singleVisualGroup'
        else:
            return self['config']['singleVisual']['visualType']

    def get_font(self):
        """
        Returns the font used in the visual text.
        :return:
        """
        return self['config']['singleVisual']['objects']['text'][1]['properties']['fontFamily']['expr']['Literal'][
            'Value']

    def update_font(self, new_font):
        """
        Updates the font used in the visual text.
        :return:
        """
        self['config']['singleVisual']['objects']['text'][1]['properties']['fontFamily']['expr']['Literal'][
            'Value'] = new_font

    def copy(self):
        """
        Returns a new Power BI container.
        :return: a Power BI container
        """
        return PbiContainer(self.export())

    def export(self):
        """
        Converts appropriate dictionaries to json strings
        :return: a dictionary
        """
        container = copy.deepcopy(self)
        try:
            self['filters'] = self.export_filters()
        except KeyError:
            pass
        container['config'] = container['config'].export()
        try:
            self['query'] = json.dumps(self['query'])
        except KeyError:
            pass
        try:
            self['dataTransforms'] = json.dumps(self['dataTransforms'])
        except KeyError:
            pass

        return container

    def update_keep_layer_order(self):
        """
        Updates keep layer order if appropriate
        :return: a boolean whether the update actually happened
        """
        if self.is_group:
            return False
        else:
            if 'vcObjects' not in self['config']['singleVisual'].keys():
                self['config']['singleVisual']['vcObjects'] = {}
            if 'general' not in self['config']['singleVisual']['vcObjects'].keys():
                self['config']['singleVisual']['vcObjects']['general'] = []
            if not self['config']['singleVisual']['vcObjects']['general']:
                self['config']['singleVisual']['vcObjects']['general'].append({})
            if 'properties' not in self['config']['singleVisual']['vcObjects']['general'][0]:
                self['config']['singleVisual']['vcObjects']['general'][0]['properties'] = {}
            if 'keepLayerOrder' not in self['config']['singleVisual']['vcObjects']['general'][0]['properties']:
                self['config']['singleVisual']['vcObjects']['general'][0]['properties']['keepLayerOrder'] = {
                    'expr': {
                        'Literal': {
                            'Value': 'false'
                        }
                    }
                }
            if self['config']['singleVisual']['vcObjects']['general'][0]['properties']['keepLayerOrder']['expr'][
                'Literal']['Value'] == 'false':
                self['config']['singleVisual']['vcObjects']['general'][0]['properties']['keepLayerOrder']['expr'][
                    'Literal']['Value'] = 'true'
                return True
            else:
                return False

    def update_multiselect(self, allow_control=False, allow_all=True, unselect_all=True):
        """
        Updates multiselection not to allow CTRL key and returns a boolean
        :allow_control: a boolean to allow CTRL selection or not
        :allow_all: a boolean to allow 'select all' or not
        :select_all: a boolean to select all by default or not
        :return: the number of updates
        """
        res_ctrl = 0
        res_allow_all = 0
        res_unselect_all = 0
        try:
            if 'selection' not in self['config']['singleVisual']['objects']:
                self['config']['singleVisual']['objects']['selection'] = [
                    {
                        'properties': {
                            'strictSingleSelect': {'expr': {'Literal': {'Value': 'false'}}},
                            'singleSelect': {'expr': {'Literal': {'Value': str(allow_control).lower()}}},
                            'selectAllCheckboxEnabled': {'expr': {'Literal': {'Value': str(allow_all).lower()}}}
                        }
                    }
                ]
                res_ctrl += 1
                res_allow_all += 1
            elif 'strictSingleSelect' not in self['config']['singleVisual']['objects']['selection'][0]['properties'] or \
                    self['config']['singleVisual']['objects']['selection'][0]['properties']['strictSingleSelect'][
                        'expr']['Literal']['Value'] == 'true':
                return 0, 0, 0
        except KeyError as e:
            print(f'KeyError in updating multi-select for {self.display_name}: {e}.')
            return 0, 0, 0
        try:
            if 'singleSelect' in self['config']['singleVisual']['objects']['selection'][0]['properties']:
                if self['config']['singleVisual']['objects']['selection'][0]['properties']['singleSelect'][
                    'expr']['Literal']['Value'] != str(allow_control).lower():
                    self['config']['singleVisual']['objects']['selection'][0]['properties']['singleSelect']['expr'][
                        'Literal']['Value'] = str(allow_control).lower()
                    res_ctrl += 1
            else:
                self['config']['singleVisual']['objects']['selection'][0]['properties']['singleSelect'] = {
                    'exp': {
                        'Literal': {
                            'Value': str(allow_control).lower()
                        }
                    }
                }
                res_ctrl += 1
        except KeyError as e:
            print(f'KeyError in updating multi-select for {self.display_name} (CTRL): {e}.')
        try:
            if 'selectAllCheckboxEnabled' in self['config']['singleVisual']['objects']['selection'][0][
                'properties']:
                if self['config']['singleVisual']['objects']['selection'][0]['properties']['selectAllCheckboxEnabled'][
                    'expr'][
                    'Literal']['Value'] != str(allow_all).lower():
                    self['config']['singleVisual']['objects']['selection'][0]['properties']['selectAllCheckboxEnabled'][
                        'expr'][
                        'Literal']['Value'] = str(allow_all).lower()
                    res_allow_all += 1
            else:
                self['config']['singleVisual']['objects']['selection'][0]['properties']['selectAllCheckboxEnabled'] = {
                    'exp': {
                        'Literal': {
                            'Value': str(allow_all).lower()
                        }
                    }
                }
                res_allow_all += 1
        except KeyError as e:
            print(f'KeyError in updating multi-select for {self.display_name} (enabling select all): {e}.')
        if unselect_all:
            try:
                if self['config']['singleVisual']['objects']['general'][0]['properties']:
                    self['config']['singleVisual']['objects']['general'][0]['properties'] = {}
                    res_unselect_all += 1
            except KeyError as e:
                print(f'KeyError in updating multi-select {self.display_name} (unselect all): {e}.')
        return res_ctrl, res_allow_all, res_unselect_all

    def add_search(self):
        """
        Enables the search feature for a slicer
        :return: a boolean whether the update actually happened
        """
        if self.type == 'slicer':
            try:
                res = (self['config']['singleVisual']['objects']['general'][0]['properties']['selfFilterEnabled']['expr']['Literal']['Value']) == 'true'
                self['config']['singleVisual']['objects']['general'][0]['properties']['selfFilterEnabled'] = {
                    'expr': {'Literal': {'Value': 'true'}}
                }
                return res

            except KeyError as e:
                print(f'Error in enabling search for {self.display_name}: {e}.')
                return False

    def disable_headers(self):
        """
        Disables headers for the visual and returns a boolean
        :return: a boolean whether the update actually happened
        """
        try:
            if 'visualHeader' in self['config']['singleVisual']['vcObjects']:
                if \
                self['config']['singleVisual']['vcObjects']['visualHeader'][0]['properties']['show']['expr']['Literal'][
                    'Value'] == 'true':
                    self['config']['singleVisual']['objects']['selection'][0]['properties']['singleSelect']['expr'][
                        'Literal']['Value'] = 'false'
                    return True
                else:
                    return False
            else:
                self['config']['singleVisual']['vcObjects']['visualHeader'] = [
                    {
                        'properties':
                            {
                                'show':
                                    {
                                        'expr':
                                            {
                                                'Literal':
                                                    {
                                                        'Value':
                                                            'false'
                                                    }
                                            }
                                    }
                            }
                    }
                ]
                return True
        except KeyError:
            print(f'Error in disabling headers: {self.display_name}.')
            return False

    def remove_from_mobile(self):
        """
        Removes the visual from the mobile screen and returns a boolean
        :return: a boolean whether the update actually happened
        """
        try:
            layouts = self['config']['layouts']
            if len(layouts) == 2:
                self['config']['layouts'] = layouts[:1]
                return True
            assert len(layouts) == 1
            return False
        except (KeyError, AssertionError):
            print(f'Error in removing visual from mobile: {self.display_name}.')
            return False

    def hide(self):
        """
        Hides the visual
        :return: None
        """
        if self.is_group:
            self['config']['singleVisualGroup']['isHidden'] = True
        else:
            self['config']['singleVisual']['display'] = {'mode': 'hidden'}

    def update_position(
            self,
            x=None,
            y=None,
            z=None,
            width=None,
            height=None
    ):
        """
        Updates the position of the visual
        :param x: new x or None
        :param y: new y or None
        :param z: new z or None
        :param width: new width or None
        :param height: new height or None
        :return: None
        """
        if x is not None:
            self['config']['layouts'][0]['position']['x'] = x
            self['x'] = round(x, 2)
        if y is not None:
            self['config']['layouts'][0]['position']['y'] = y
            self['y'] = round(y, 2)
        if z is not None:
            self['config']['layouts'][0]['position']['z'] = z
            self['z'] = round(z, 2)
        if width is not None:
            self['config']['layouts'][0]['position']['width'] = width
            self['width'] = round(width, 2)
        if height is not None:
            self['config']['layouts'][0]['position']['height'] = height
            self['height'] = round(height, 2)

    def update_page_link(self, page):
        """
        Update a link to point to a page.
        :param page: a Power BI page object
        :return: None
        """
        self['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['show']['expr']['Literal'][
            'Value'] = "true"
        self['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['type']['expr']['Literal'][
            'Value'] = "'PageNavigation'"
        try:
            self['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['navigationSection']['expr'][
                'Literal']['Value'] = f"'{page.name}'"
        except KeyError:
            try:
                self['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['navigationSection'] = {
                    'expr': {
                        'Literal': {
                            'Value': f"'{page.name}'"
                        }
                    }
                }
            except KeyError:
                print(f'Coud not update button link for {self.display_name}.')

    def remove_link(self, update_style=False, hide=False):
        """
        Removes a link from an action button visual and updates look and feel by setting font and border to grey or
        hiding the visual altogether if required
        :param update_style: a boolean
        :param hide: a boolean
        :return: None
        """
        self['config']['singleVisual']['vcObjects']['visualLink'][0]['properties']['show']['expr']['Literal'][
            'Value'] = "false"
        if update_style:
            for outline_state in self['config']['singleVisual']['objects']['text']:
                if 'fontColor' in outline_state['properties']:
                    outline_state['properties']['fontColor'] = {
                        'solid': {'color': {'expr': {'Literal': {'Value': "'#CCCCCC'"}}}}
                    }
            for outline_state in self['config']['singleVisual']['objects']['outline']:
                if 'lineColor' in outline_state['properties']:
                    outline_state['properties']['lineColor'] = {
                        'solid': {'color': {'expr': {'Literal': {'Value': "'#CCCCCC'"}}}}
                    }
        if hide:
            self.hide()
