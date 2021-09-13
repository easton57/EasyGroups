"""
Class to do everything that's needed with folders and directories
by: Easton Seidel
2/23/21
"""

from os import *
from shutil import *


class DirectoryServices:

    # Declare Directories
    delivered_dir = "\\\\ccm-file-10\\corporate$\\New Hire Onboarding\\Delivered Letters\\"
    undelivered_dir = "\\\\ccm-file-10\\corporate$\\New Hire Onboarding\\Undelivered Letters\\"

    def delivered(self):
        return self.delivered_dir

    def undelivered(self):
        return self.undelivered_dir

    def get_names(self):
        return [f for f in listdir(self.undelivered_dir) if path.isdir(path.join(self.undelivered_dir, f))]

    def create_directories(self, names):

        # Create a new list
        directories = []

        # iterate through the names and create a directory list
        for i in names:
            directories.append(self.undelivered_dir + i)

        return directories

    def move_folder(self, name):

        # Try to move the folder
        try:
            move(self.undelivered_dir + name, self.delivered_dir)
        except Error:
            print("Folder already exists! Saving with .new!")
            move(self.undelivered_dir + name, self.undelivered_dir + name + ".new")
            move(self.undelivered_dir + name + ".new", self.delivered_dir)
