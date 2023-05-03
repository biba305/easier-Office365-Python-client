from functools import wraps
from sharepoint import SharePoint

def decorator_root_folder(func):
    @wraps(func)
    def wrapper(self: SharePoint, parent_folder: str, *args, **kwargs):
        folder_name = self.session.web.get_folder_by_server_relative_url(
            f"{self.doc}/{parent_folder}"
        )

        return func(self, folder_name, *args, **kwargs)

    return wrapper
