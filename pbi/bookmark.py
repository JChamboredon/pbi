from pbi.object import _PbiObject


class Bookmark(dict, _PbiObject):
    """
    The Power BI bookmark class.
    """
    name_prefix = 'Bookmark'

    def __init__(self, *arg, **kw):
        """
        Creates a PbiPage bookmark with all appropriate json strings converted as dictionaries or similar objects
        :param arg: list of arguments
        :param kw: dictionary of keywords
        """
        super().__init__(*arg, **kw)

    def get_target_visuals(self):
        """
        Return the names of the visuals targeted by the bookmark
        :return: a list of strings (potentially empty)
        """
        try:
            return self['options']['targetVisualNames']
        except KeyError:
            print("""
Warning: could not find bookmark target visuals.
""")
            return []
