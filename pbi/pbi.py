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
    def _get_download_script(strg, workspace_id, destination):
        """
        Returns a PowerShell script to download reports from a workspace
        :param strg: a string to look for in the report names
        :param workspace_id: the id of the workspace
        :param destination: the destination folder
        :return: a string
        """
        return f'''
Connect-PowerBIServiceAccount

$path="{destination}"
$workspaceid_backup= "{workspace_id}"

New-Item $path -itemType Directory
$report_details= Get-PowerBIReport -WorkspaceId "$workspaceid_backup"
$report_ids=$report_details.Id
$report_names=$report_details.Name

$i=0
foreach ($report in $report_ids)''' + '''
{''' + f'''
    $current_name = $report_names[$i]
    $safe_name = "temp"
    if ($current_name.contains("{strg}"))''' + '''
    {''' + '''
        Write-Host "Importing report $report and name $current_name..." 
        Export-PowerBIReport -WorkspaceId $workspaceid_backup -Id $report -Outfile $path\\"temp.pbix"
        Move-Item -Path $path\\"temp.pbix" -Destination $path\\$current_name".pbix" -Force
    }
    $i=$i + 1
}
'''

    @staticmethod
    def _get_backup_script(workspace_id, destination):
        """
        Returns a PowerShell script to backup a workspace
        :param workspace_id: the id of the workspace
        :param destination: the destination folder
        :return: a string
        """
        return PowerBI._get_download_script(
            '',
            workspace_id,
            destination
        )

    @staticmethod
    def download(strg, workspace_id, destination):
        """
        Dowloads the reports and dataset from the workspace in the destination folder
        :param strg: A string to look for in the report names
        :param workspace_id: The id of the workspace
        :param destination: the destination folder
        :return: None
        """
        run_ps_script(PowerBI._get_download_script(strg, workspace_id, destination))

    @staticmethod
    def get_backup(workspace_id, destination):
        """
        Copies the reports and dataset from the workspace in the destination folder
        :param workspace_id: The id of the workspace
        :param destination: the destination folder
        :return: None
        """
        run_ps_script(PowerBI._get_backup_script(workspace_id, destination))

    @staticmethod
    def _get_upload_script(
            source,
            workspace_id_to
    ):
        """
        Return the PowerShell script to publish to the target workspace
        :param source: the source folder (string)
        :param workspace_id_to: a workspace id (string)
        :return: a string
        """
        return f'''
Connect-PowerBIServiceAccount
$pathinput = "{source}"
$workspaceid = "{workspace_id_to}"
$bslash = "\\"
foreach ($report in Get-ChildItem -Path $pathinput)''' + '''
    {
        $reportname = "$pathinput$bslash$report"
        New-PowerBIReport -Path $reportname -WorkspaceId $workspaceid -ConflictAction CreateOrOverwrite -ErrorAction Stop
        Write-Host "Published $pathinput$bslash$report."
    }
'''

    def upload(
            self,
            source,
            workspace_id_to
    ):
        """
        publishes the given reports to the target workspace
        :param source: the source folder (string)
        :param workspace_id_to: a workspace id (string)
        :return: a string
        """
        run_ps_script(
            self._get_upload_script(
                source,
                workspace_id_to
            )
        )
        message_box('Power BI upload', 'Done uploading.', 0)
        print('Done.')
