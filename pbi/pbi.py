from pbi.utils import run_ps_script, message_box


class PowerBI:
    def __init__(self, config):
        """
        Initiates a Power BI object with the dictionary of given configuration (workspace IDs, dataset IDs, etc.)
        :config: a dictionary
        :return: None
        """
        self.config = config
        self.download_folder = self.config['DOWNLOAD_FOLDER']
        self.upload_folder = self.config['UPLOAD_FOLDER']

    @staticmethod
    def _get_download_script(strg, destination, workspace_id=None):
        """
        Returns a PowerShell script to download reports from a workspace
        :param strg: a string to look for in the report names
        :param destination: the destination folder
        :param workspace_id: the id of a workspace or None
        :return: a string
        """

        workspace_id_filter = ''
        if workspace_id:
            workspace_id_filter = f'-WorkspaceId "{workspace_id}" '

        return f'''
$path = "{destination}"''' + '''
if (!(Test-Path $path -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $path
}
Connect-PowerBIServiceAccount

New-Item $path -itemType Directory''' + f'''
$report_list= Get-PowerBIReport {workspace_id_filter}

Write-Host $report_list.id

foreach ($report in $report_list)''' + '''
{''' + f'''
    $report_id = $report.id
    $report_name = $report.name
    if ($report_name.contains("{strg}"))''' + '''
    {''' + f'''
        Write-Host "Importing report $report_name ($report_id)..." 
        Export-PowerBIReport {workspace_id_filter}-Id $report_id -Outfile $path\\"temp.pbix"
        Move-Item -Path $path\\"temp.pbix" -Destination $path\\$report_name".pbix" -Force''' + '''
    }
}'''

    @staticmethod
    def _get_backup_script(**kwargs):
        """
        Returns a PowerShell script to back up a workspace
        :return: a string
        """
        return PowerBI._get_download_script(
            '',
            **kwargs
        )

    @staticmethod
    def download(*args, **kwargs):
        """
        Downloads the reports and dataset from the workspace in the destination folder
        :return: None
        """
        run_ps_script(PowerBI._get_download_script(*args, **kwargs))

    @staticmethod
    def get_backup(**kwargs):
        """
        Copies the reports and dataset from the workspace in the destination folder
        :return: None
        """
        run_ps_script(PowerBI._get_backup_script(**kwargs))

    @staticmethod
    def _get_upload_script(
            source,
            workspace_id=None
    ):
        """
        Return the PowerShell script to publish to the target workspace
        :param source: the source folder (string)
        :param workspace_id: a workspace id (string) or None
        :return: a string
        """
        workspace_id_filter = ''
        if workspace_id:
            workspace_id_filter = f'-WorkspaceId "{workspace_id}" '

        return f'''
$pathinput = "{source}"''' + '''
if (!(Test-Path $pathinput -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $pathinput
}
Connect-PowerBIServiceAccount
$bslash = "\\"
foreach ($report in Get-ChildItem -Path $pathinput)''' + '''
    {
        $reportname = "$pathinput$bslash$report"''' + f'''
        New-PowerBIReport -Path $reportname {workspace_id_filter}-ConflictAction CreateOrOverwrite -ErrorAction Stop''' + '''
        Write-Host "Published $pathinput$bslash$report."
    }
'''

    def upload(
            self, *args, **kwargs
    ):
        """
        publishes the given reports to the target workspace
        :return: a string
        """
        run_ps_script(
            self._get_upload_script(
                *args,
                **kwargs
            )
        )
        message_box('Power BI upload', 'Done uploading.', 0)
        print('Done.')
