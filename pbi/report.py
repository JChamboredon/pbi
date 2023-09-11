import json
import os

import pandas as pd
import shutil

from pbi.layout import PbiLayout
from pbi.utils import run_ps_script


class PbiReport:
    """
    The Power BI report class.
    """
    ext = 'pbix'
    archive_format = 'zip'

    def __init__(self, folder, filename):
        """
        Initiates a Power BI Report object.
        :param folder: a path to a local folder
        :param filename: a string
        """
        self.folder = folder
        self.filename = filename
        self._open()
        with open(self.temp_folder + '\\Report\\Layout', 'r', encoding='utf-16-le') as file:
            layout_str = file.read()
        self.layout = PbiLayout(layout_str)
        try:
            with open(self.temp_folder + '\\Connections', 'r') as file:
                connections_str = file.read()
            self.connections = json.loads(connections_str)
        except FileNotFoundError:
            self.connections = None
            print("Warning: No connection file found.")
        self._close()

    @property
    def path(self):
        """
        Returns the full path to the Power BI report.
        :return: a string
        """
        return f'{self.folder}\\{self.filename}.{self.ext}'

    @property
    def temp_folder(self):
        """
        returns a temporary folder name (will be used to unzip the .pbix file)
        :return: a string
        """
        return f'Temp_{self.filename}'

    @property
    def filters(self):
        return self.layout['filters']

    def copy(self, new_name=None, new_folder=None):
        """
        Copies a Power BI report and returns the corresponding PbiReport
        :param new_name: the new name of the copied file (string)
        :param new_folder: the name of the folder where to place the copied file (string)
        :return: a PbiReport object
        """
        if new_name is None:
            new_name = f'{self.filename}_copy'
        if new_folder is None:
            new_folder = f'{self.folder}'
        new_path = f'{new_folder}\\{new_name}.{self.ext}'
        shutil.copyfile(self.path, new_path)
        return PbiReport(new_folder, new_name)

    def _open(self):
        """
        Decomposes the PBI Report in its Temp folder (unzip)
        :return: None
        """
        shutil.rmtree(self.temp_folder, ignore_errors=True)
        os.mkdir(self.temp_folder)
        shutil.unpack_archive(self.path, self.temp_folder, self.archive_format)

    def _close(self):
        """
        Saves the PBI Report as a .pbix (zips) and removes Temp folder
        :return: None
        """
        os.remove(self.path)
        shutil.make_archive(f'{self.path}', self.archive_format, self.temp_folder)
        os.rename(
            f'{self.path}.{self.archive_format}',
            f'{self.path}'
        )
        shutil.rmtree(self.temp_folder)

    def save(self, dataset_id_from=None, dataset_id_to=None):
        """
        Saves the Python Power BI Report as a .pbix file.
        :param dataset_id_from: a string
        :param dataset_id_to: a string
        :return: None
        """
        self._open()
        self.tidy_bookmarks()
        layout = self.layout.export()
        os.remove(self.temp_folder + '\\Report\\Layout')
        new_layout_file = open(self.temp_folder + '\\Report\\Layout', "a", encoding='utf-16-le')
        new_layout_file.write(layout)
        new_layout_file.close()

        if dataset_id_from is not None and dataset_id_to is not None and dataset_id_to != dataset_id_from:
            with open(self.temp_folder + '\\Connections', 'r') as connection_file:
                connection_str = connection_file.read()
                new_connection_str = connection_str.replace(dataset_id_from, dataset_id_to)
            connection_file.close()
            os.remove(self.temp_folder + '\\Connections')

            new_connections_file = open(self.temp_folder + '\\Connections', "a")
            new_connections_file.write(new_connection_str)
            new_connections_file.close()

        # os.remove(self.temp_folder + '\\SecurityBindings')
        self._close()

    def get_page(self, page_name):
        """
        Returns the page with the given name
        :param page_name: a string
        :return: a power BI Page object
        """
        return self.layout.get_page(page_name)

    def select_pages(self, page_list):
        """
        Select the sections (pages) from the PBI Report which names appear in page_list.
        :param page_list: a list of strings
        :return: None
        """
        self.layout['sections'] = [
            section for section in self.layout['sections']
            if section['displayName'] in page_list
        ]
        self.tidy_bookmarks()

    def get_bookmarks(self, names=None):
        """
        Returns the list of bookmarks used in the report
        :param names: a list of bookmark names, a string, or None
        :return: a list of dictionaries representing bookmarks
        """
        return self.layout.get_bookmarks(names)

    def tidy_bookmarks(self):
        """
        removes bookmarks from report list if they are not used anywhere in the page visuals
        :return: None
        """
        used_bookmarks = self.layout.get_used_bookmark_names()
        if 'bookmarks' not in self.layout['config']:
            return None
        self.layout['config']['bookmarks'] = [
            bookmark for bookmark in self.layout['config']['bookmarks']
            if bookmark['name'] in used_bookmarks or (
                'children' in bookmark.keys()
                and {
                    sb['name'] for sb in bookmark['children']
                }.intersection(used_bookmarks)
            )
        ]

    def merge(self, report):
        """
        Merges two reports
        :param report: the PBI Report from which pages have to be merged or None
        :return: None
        """
        if report is None:
            return None
        report.tidy_bookmarks()
        self.layout['sections'] += report.layout['sections']
        self._update_section_id()
        bookmarks_to_add = []
        for page in report.layout['sections']:
            page_bookmarks_in_visuals = page.get_used_bookmark_names()
            if 'bookmarks' in report.layout['config']:
                bookmarks_to_add += [
                    bookmark
                    for bookmark in report.layout['config']['bookmarks']
                    if (
                            bookmark['name'] in page_bookmarks_in_visuals
                            or (
                                    'children' in bookmark.keys()
                                    and {
                                        sb['name'] for sb in bookmark['children']
                                    }.intersection(page_bookmarks_in_visuals)
                            )
                    )
                ]
        if 'bookmarks' not in self.layout['config']:
            self.layout['config']['bookmarks'] = []
        self.layout['config']['bookmarks'] += bookmarks_to_add

    def _update_section_id(self, id_list=None):
        """
        Updates the ids of the report sections (pages).
        NB: Whether these ids are used anywhere is currently unclear.
        :param id_list: a list of integers
        :return: None
        """
        if id_list is None:
            id_list = list(range(len(self.layout['sections'])))
        for i, section in enumerate(self.layout['sections']):
            section['id'] = id_list[i]
            section['ordinal'] = section['id']

    def update_keep_layer_order(self):
        """
        Sets the 'keepLayerOrder' to all appropriate visuals to 'true'
        :return: None
        """
        updates = []
        for page in self.layout['sections']:
            updates += [[page.display_name, page.update_keep_layer_order()]]
        print(f"""
Number of visuals updated per page (keep layer order):
{pd.DataFrame(updates, columns=['Page', 'Updates'])}
""")

    def update_multiselect(self):
        """
        Sets the multiselect slicers not to allow selection with CTRL key
        :return: None
        """
        updates = []
        for page in self.layout['sections']:
            updates += [[page.display_name, *page.update_multiselect()]]
        print(f"""
Number of multi-select slicers updated per page:
{pd.DataFrame(updates, columns=['Page', 'CTRL Updates', 'Allow all Updates', 'Unselect all Updates'])}
""")

    def remove_visuals_from_mobile(self):
        """
        Removes all the visuals from the mobile screen
        :return: None
        """
        updates = []
        for page in self.layout['sections']:
            updates += [[page.display_name, page.remove_visuals_from_mobile()]]
        print(f"""
Number of visuals removed from mobile screen:
{pd.DataFrame(updates, columns=['Page', 'Updates'])}
""")

    def reset_mobile_screen(self, default_message_report=None, pages=None):
        """
        Resets the mobile screens for the report.
        :param default_message_report: a Power BI report containing the default screen message to display
        :param pages: a list of page names to apply the default mobile message too (exclude tooltip pages, etc.);
        default is None (all pages)
        :return: None
        """
        if pages is None:
            pages = self.layout.pages
        self.remove_visuals_from_mobile()
        if default_message_report is not None:
            default_visuals = default_message_report.layout['sections'][0]['visualContainers']
            for page_name in pages:
                page = self.get_page(page_name)
                if page is not None:
                    page.add_visuals(default_visuals)

    def disable_headers(self, **kwargs):
        """
        Disables headers for all visuals
        :return: None
        :kwargs:
        """
        updates = []
        for page in self.layout['sections']:
            updates += [[page.display_name, page.disable_headers(**kwargs)]]
        print(f"""
Number of headers disabled per page:
{pd.DataFrame(updates, columns=['Page', 'Updates'])}
""")

    def add_search(self, **kwargs):
        """
        Add search feature for all slicers
        :return: None
        :kwargs:
        """
        updates = []
        for page in self.layout['sections']:
            updates += [[page.display_name, page.add_search(**kwargs)]]
        print(f"""
Number of search features enabled per page:
{pd.DataFrame(updates, columns=['Page', 'Updates'])}
""")

    def add_resource_package(self, report, name, item):
        """
        Updates and saves the report to add the resource package (both in layout and linked files)
        :param report: the report with the resource package to add
        :param name: the name of the resource package
        :param item: the actual resource package item (dictionary)
        :return: None
        """
        self.layout.add_resource_packages(name, item)
        report._open()
        self._open()
        src_file = report.temp_folder + f"/Report/StaticResources/{name}/{item['name']}"
        tgt_file = self.temp_folder + f"/Report/StaticResources/{name}/{item['name']}"
        shutil.copyfile(src_file, tgt_file)
        report._close()
        self._close()

    @classmethod
    def _get_download_script(cls, name, workspace_id, destination):
        """
        generates a PowerShell script to download the report from a workspace
        :param name: the name of the report to download
        :param workspace_id: the id of the Power BI workspace to download from
        :param destination: the destination folder
        :return: a string
        """
        return f'''
Connect-PowerBIServiceAccount
$path = "{destination}"
$workspaceid = "{workspace_id}"
$report_details= Get-PowerBIReport -WorkspaceId "$workspaceid"
$report_ids = $report_details.Id
$report_names = $report_details.Name

$i=0
foreach ($report in $report_ids)''' + '''
{''' + f'''
    $current_name = $report_names[$i]
    if ($current_name -eq "{name}")''' + '''
    {''' + '''
        Write-Host "Downloading report $report and name $current_name..."
        Export-PowerBIReport -WorkspaceId $workspaceid -Id $report -Outfile $path\\$current_name".pbix"
    }
    $i=$i + 1''' + '''
}
'''

    @classmethod
    def download(cls, name, workspace_id, destination):
        """
        Downloads a report from a Power BI workspace.
        :param name: the report name
        :param workspace_id: the Power BI workspace ID
        :param destination: the destination folder
        :return: A Power BI report object
        """
        ps_script = PbiReport._get_download_script(name, workspace_id, destination)
        run_ps_script(ps_script)

        return PbiReport(
            destination,
            name
        )

    def update_names(self, dct):
        """
        Updates the hardcoded names in the report: e.g. filter names, etc.
        :param dct: A dictionary with old and new names e.g. {'old_name': 'new_name'}
        :return: None
        """
        for old_name in dct:
            new_name = dct[old_name]
            old_name = old_name.replace("'", "''")
            new_name = new_name.replace("'", "''")
            self.layout.replace(f"'{old_name}'", f"'{new_name}'")

    def set_landing_page(self, landing_page):
        """
        Sets the landing page of the report
        :page: a page of the report
        :return: None
        """
        self.layout['sections'] = [landing_page] + [
            page for page in self.layout['sections']
            if page.display_name != landing_page.display_name
        ]
        self._update_section_id()
