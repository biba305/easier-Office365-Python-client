import os
from dataclasses import dataclass, field
from decorators import decorator_root_folder
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential


__version__ = "1.0.0"
__author__ = "Khilal-Shpiro Mukhammed"


@dataclass
class SharePoint:
    """
    Module to make working with SharePoint more easy and comfortable

    Current functionality:
        - Authorize to SharePoint via your username and password = when init the class
        - Upload file to specific folder = upload_file
        - Download file to your local folder = download_file
        - Create folder in Sharepoint = create_folder
        - Get file properties in specific folder = get_file_properties_from_folder

    __init__ parametrs:
        ::url - url to the needed SharePoint, example: https://<company_name>.sharepoint.com
        ::username - your username
        ::password - your password
        ::site_name - you can find it in the web-browser, example: ps.all
        ::doc - default "Shared Documents"
        ::session - your login session
    """
    url: str
    username: str
    password: str
    site_name: str
    doc: str = "Shared Documents"
    session: ClientContext = field(init=False)

    def __post_init__(self):
        """
        Description:
            Function to authenticate to Sharepoint using credentials.
            It then saves your session to self.session so you will have no need to authenicate 100 times
        """
        self.session = ClientContext(f"{self.url}/sites/{self.site_name}").with_credentials(
            UserCredential(self.username, self.password)
        )

    def upload_file(self, local_file_path: str, sharepoint_destination: str) -> None:
        """
        Description:
            Uploads file to specific Sharepoint folder

        Parametrs:
            ::local_file_path - path to local file. Can be absolute path or relative
            ::sharepoint_destination - folder in sharepoint where the file you want should be located

        Example:
            share.upload_file(
                local_file_path="examples/test.csv", 
                sharepoint_destination="General/Документы"
            )

        Return:
            This function does not return anything, it just uploads the file to specified directory
        """
        file_content = self._get_local_file_content(local_file=local_file_path)
        self._send_file_to_sharepoint(
            file_name=os.path.basename(local_file_path), folder_name=sharepoint_destination, file_content=file_content
        )

    def download_file(self, sharepoint_source: str, destination: str) -> None:
        """
        Description:
            Download any file from sharepoint

        Parametrs:
            ::sharepoint_source - path to file in SharePoint
            ::destination - where to download the needed file

        Example:
            share.download_file(
                sharepoint_source="General/Документы/test.csv", 
                destination="downloads"
            )
            The sharepoint_source is like General/Документы/file_name, 
            because we already assigned /sites/ps.all/Shared Documents in the init method

        Return:
            This function does not return anything, it just downloads a file to the specified directory
        """
        file_obj = self._get_sharepoint_file_content(source=sharepoint_source)
        self._save_file(
            file_name=os.path.basename(sharepoint_source), file_obj=file_obj, destination=destination
        )

    @decorator_root_folder
    def create_folder(self, parent_folder, new_folder: str) -> None:
        """
        Description:
            Create new folder under some parent folder

        Parametrs:
            ::parent_folder - path to the parent folder
            ::new_folder - name of the new folder

        Example:
            share.create_folder(
                parent_folder="General/Документы", 
                new_folder="test"
            )
            Now we created folder "test" in "Документы" folder

        Return:
            This function does not return anything, it just creates folder under it's parent folder
        """
        new_folder = parent_folder.folders.add(new_folder).execute_query()

    @decorator_root_folder
    def get_files_list(self, parent_folder) -> list:
        """
        Description:
            Get list of files in the specified path

        Parametrs:
            ::folder_name - folder in which we want to get a list of files

        Example:
            get_files_list("General/Документы")

        Returns:
            [<office365.sharepoint.files.file.File object at 0x7f7ab305ece0>]

            we recived one object, because we have one file in 
            "/sites/ps.all/Shared Documents/General/Документы"
        """
        parent_folder.expand(["Files", "Folders"]).get().execute_query()

        return parent_folder.files

    @decorator_root_folder
    def get_folder_list(self, parent_folder: str) -> list:
        """
        Description:
            Get list of folders in the specified path

        Parametrs:
            ::folder_name - folder in which we want to get a list of folders

        Example:
            get_folder_list("General")

        Returns:
            [<office365.sharepoint.folders.folder.Folder object at 0x7fc8fd266d10>, 
            <office365.sharepoint.folders.folder.Folder object at 0x7fc8fd267790>]

            we recived two objects, because we have two folders in 
            "/sites/ps.all/Shared Documents/General"
        """
        parent_folder.expand(["Folders"]).get().execute_query()

        return parent_folder.folders

    def get_file_properties_from_folder(self, folder_name: str) -> list:
        """
        Description:
            Get file properties from specific folder

        Parametrs:
            ::folder_name - folder that contain files that we want to get info about them

        Usage:
            info = share.get_file_properties_from_folder(
                folder_name="General"
            )

            print(info)

        Return:
            [
                {
                    "file_id": "6956...4981",
                    "file_name": "test.jpg",
                    "major_version": 1,
                    "minor_version": 0,
                    "file_size": 2520331,
                    "time_created": "2021-12-24T14:39:30Z",
                    "time_last_modified": "2021-12-24T14:39:35Z"
                },
                ...
                {
                    "file_id": "3ddb...8acb",
                    "file_name": "test.csv",
                    "major_version": 29,
                    "minor_version": 0,
                    "file_size": 32448,
                    "time_created": "2023-03-30T06:11:29Z",
                    "time_last_modified": "2023-04-07T13:15:08Z"
                }
            ]
        """
        files_list = self.get_files_list(folder_name)

        def name(files):
            for file in files:
                yield {
                    'file_id': file.unique_id,
                    'file_name': file.name,
                    'major_version': file.major_version,
                    'minor_version': file.minor_version,
                    'file_size': file.length,
                    'time_created': file.time_created,
                    'time_last_modified': file.time_last_modified,
                }

        return list(name(files=files_list))

    def _get_sharepoint_file_content(self, source: str) -> bytes:
        """
        Description:
            Get the file from specific folder. Reads it in binary so it will be possible to save it
        """
        file = File.open_binary(
            self.session, 
            f'/sites/{self.site_name}/{self.doc}/{source}'
        )

        return file.content

    @staticmethod
    def _save_file(file_name: str, file_obj: bytes, destination: str) -> None:
        """
        Description:
            Saves file to specific local folder

        Parametrs:
            ::file_name - how the file will be named as saved
            ::file_obj - result of _get_file function (binary data)
            ::destination - where the file will be downloaded
        """
        if not os.path.exists(destination):
            os.makedirs(destination, exist_ok=True)

        file_dir_path = os.path.join(destination, file_name)
        with open(file_dir_path, 'wb') as f:
            f.write(file_obj)

    def _get_local_file_content(self, local_file: str) -> bytes:
        """
        Description:
            Get the file from specific folder. Reads it in binary so it will be possible to save it

        Parametrs:
            ::local_file - Path to local file
        """
        with open(local_file, 'rb') as lf:
            return lf.read()

    def _send_file_to_sharepoint(self, file_name: str, folder_name: str, file_content: bytes) -> None:
        """
        Description:
            function that uploads a file

        Parametrs:
            ::file_name - name of the uploaded file in SharePoint (will be the same as the local file)
            ::folder_name - folder where the file will be uploaded
            ::file_content - binary content of the file
        """
        target_folder = self.session.web.get_folder_by_server_relative_path(
            f'/sites/{self.site_name}/{self.doc}/{folder_name}'
        )
        target_folder.upload_file(file_name, file_content).execute_query()
