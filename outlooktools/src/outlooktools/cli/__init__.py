# SPDX-FileCopyrightText: 2023-present richmr <richmr@users.noreply.github.com>
#
# SPDX-License-Identifier: MIT
import typer

from outlooktools.__about__ import __version__
from outlooktools.outlook_connection import OutlookConnection

app = typer.Typer(name="OutlookTools")

@app.command()
def count_items(folder_path:str):
    """
    Returns a count of items in the designated folder
    Expects a folder_path like Inbox/Folder1/Subfolder1 or Inbox\\Folder1\\Subfolder 2
    """
    oc = OutlookConnection()
    folder = oc.traverseFolders(folder_path)
    messages = list(folder.Items)
    print(f"There are {len(messages)} messages in {folder_path}")
    print(f"First subject: {messages[0].Subject}")

@app.command()
def list_mailboxes():
    """
    Prints a list of all available mailboxes for this Outlook instance
    """
    oc = OutlookConnection()
    print("Available mailboxes:")
    for folder in oc.outlook.Folders:
        print(f"\t{folder.name}")
    print("")
    print("Prepend this to your folder path to delve into correct mailbox")


def outlooktools():
    app()