"""
Just a class
A lovely class to add additional groups
By: Me, Easton Seidel
"""

import pyad
from pyad import adquery, adgroup, aduser
import keyring
import io
from datetime import datetime
from DirectoryServices import *
from getpass import getpass
from time import sleep


class EasyGroups:

    def __init__(self, auto):
        """ Initialization for the class """
        # Query and other AD info
        self.q = adquery.ADQuery()
        self.domain_name = "domain"
        self.top_level_domain = "local"

        # Creds info
        self.user = ""
        self.password = keyring.get_password("system", self.user)

        # Get creds if not automatic
        if not auto:
            self.user = input("Enter your username: ")
            self.password = getpass("Enter your password: ")

        # Initialize AD
        pyad.pyad_setdefaults(ldap_server=f"{self.domain_name}.{self.top_level_domain}", username=self.user,
                              password=self.password)

        # Initialize Directory Services
        self.ds = DirectoryServices()

        # Variable for who was done and who was done previously
        self.previous = []
        self.done = []
        self.fail = []

        # Get old log data
        self.old_log()

        # Date Variables
        now = datetime.now()
        self.date = now.strftime("%m-%d-%y")

        # Initialize group values
        self.init_groups()

    def add_groups(self):
        """ Method to add AD groups to new hires """
        # Get the new hires
        users_og = self.ds.get_names()
        users = []

        # Filter out users that already have their groups
        for i in users_og:
            if i not in self.previous:
                users.append(i)

        # Loop through the users and add the groups
        for i in users:
            # Wait a little bit of time
            sleep(2)

            # Assign the user
            user = pyad.aduser.ADUser.from_cn(i)

            # Query for the users title
            self.q.execute_query(
                attributes=["description", "distinguishedName"],
                where_clause="displayName = '{}'".format(i),
                base_dn=f"DC={self.domain_name}, DC={self.top_level_domain}"
            )

            # Break if the user doesn't exist anymore
            if self.q.get_row_count() <= 0:
                print(i + " does not exist on " + self.domain_name + "." + self.top_level_domain + "\n"
                      "Please verify that the user is still scheduled for onboarding.")
            else:
                # Separate groups
                for row in self.q.get_results():
                    long_name = row["distinguishedName"]
                    title = row["description"]

                    # Get the user location
                    user_loc = str(row['distinguishedName']).replace(",OU=", ",DC=")
                    user_loc = user_loc.split(',DC=')
                    user_loc = user_loc[2]

                # Convert title to a string
                title = title[0]

                # Add to archive if not corp
                if user_loc != 'Corporate - 1001 (Salt Lake City\\, UT)':
                    user.add_to_group(self.archive)

                # Make sure the user hasn't already been done
                if i not in self.previous:
                    # Try this and print an error otherwise
                    try:
                        # Loop through the list with the matching title
                        if title in ["Loan Officer", "Loan Officer Assistant", "Loan Coordinator", "Production Coordinator"]:
                            # Add to groups
                            user.add_to_group(self.lo)
                            user.add_to_group(self.production)
                            user.add_to_group(self.non_exempt)

                            # Message for preset
                            message = f"User {i} has been updated using Loan Officer/Coordinator and Production Coordinator preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Loan Processor":
                            # Add to groups
                            user.add_to_group(self.lo)
                            user.add_to_group(self.non_exempt)
                            user.add_to_group(self.processor)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Loan Processor preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Branch Manager":
                            # Add to groups
                            user.add_to_group(self.bm)
                            user.add_to_group(self.exempt)
                            user.add_to_group(self.receptionist)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.lo)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Branch Manager preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Receptionist":
                            # Add to groups
                            user.add_to_group(self.non_exempt)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Receptionist preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title in ["Closer", "Insuring Specialist", "Funder", "Funding Assistant"]:
                            # Add to groups
                            user.add_to_group(self.cie)
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.ops)
                            user.add_to_group(self.ops_fs)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.sharefile)

                            # Message for preset
                            message = f"User {i} has been updated using a Closing preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title in ["Final Docs Specialist", "Shipper", "Post-Closing Specialist",
                                       "Post-Closing Specialist I", "Post-Closing Specialist II",
                                       "Post-Closing Specialist III", "Post Closing Specialist",
                                       "Post Closing Specialist I", "Post Closing Specialist II",
                                       "Post Closing Specialist III"]:
                            # Add to groups
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.filezilla)
                            user.add_to_group(self.ops)
                            user.add_to_group(self.ops_fs)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.sharefile)

                            # Message for preset
                            message = f"User {i} has been updated using an Operations preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Collateral Specialist":
                            # Add to groups
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.filezilla)

                            # Message for Preset
                            message = f"User {i} has been updated using Collateral Specialist preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        else:
                            # Message for Preset
                            message = f"User {i} did not match any presets"

                            # Print out what was done
                            print(message)

                            self.fail.append(i)
                    except Exception as e:
                        message = f"There was an error with adding groups for {i}!"
                        print(message)
                else:
                    message = f"User {i} already has their groups."
                    print(message)

            self.run_log(message)

        # Write the log
        self.finished_logs()

    def run_log(self, message):
        with io.open(f"H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Logs\\{self.date}.txt", "a") as log:
            log.write(message + "\n")

    def finished_logs(self):
        # Open the log file
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Finished.txt", "a") as log:
            # Loop through the users and add them
            for i in self.done:
                log.write(f"\n{i}")

        # Create the failed log
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Failed.txt", "a") as log:
            # Loop through the users and add them
            for i in self.fail:
                log.write(f"\n{i}")

    def old_log(self):
        # open the log file and read all the lines
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Finished.txt", "r") as log:
            for line in log:
                self.previous.append(line.rstrip('\n'))

    def init_groups(self):
        # MailArchive
        self.archive = pyad.adgroup.ADGroup.from_dn("CN=MailArchive,OU=Security Groups,OU=Exchange Objects,"
                                                    "OU=Offices,DC=domain,DC=local")
        # LoanOfficers
        self.lo = pyad.adgroup.ADGroup.from_dn("CN=Loan Officers,OU=Distribution Group,OU=Exchange Objects,"
                                               "OU=Offices,DC=domain,DC=local")
        # ProductionUsers
        self.production = pyad.adgroup.ADGroup.from_dn("CN=Production Users,OU=Citrix,OU=Groups,"
                                                       "DC=domain,DC=local")
        # Non-ExemptEmployee
        self.non_exempt = pyad.adgroup.ADGroup.from_dn("CN=Non-Exempt Employee,OU=Distribution Group,"
                                                       "OU=Exchange Objects,OU=Offices,DC=domain,DC=local")
        # Processors
        self.processor = pyad.adgroup.ADGroup.from_dn("CN=Processors,OU=Distribution Group,OU=Exchange Objects,"
                                                      "OU=Offices,DC=domain,DC=local")
        # Branch Managers
        self.bm = pyad.adgroup.ADGroup.from_dn("CN=Branch Managers,OU=Distribution Group,OU=Exchange Objects,"
                                               "OU=Offices,DC=domain,DC=local")
        # Exempt
        self.exempt = pyad.adgroup.ADGroup.from_dn("CN=Exempt Employee,OU=Distribution Group,OU=Exchange Objects,"
                                                   "OU=Offices,DC=domain,DC=local")
        # FS-Receptionist
        self.receptionist = pyad.adgroup.ADGroup.from_dn("CN=FS-Reception,OU=File Server,OU=Groups,"
                                                         "DC=domain,DC=local")
        # FS-HDriveAccess
        self.hdrive = pyad.adgroup.ADGroup.from_dn("CN=FS-HDriveAccess,OU=File Server,OU=Groups,"
                                                   "DC=domain,DC=local")
        # CIE
        self.cie = pyad.adgroup.ADGroup.from_dn("CN=CIE,OU=Distribution Group,OU=Exchange Objects,"
                                                "OU=Offices,DC=domain,DC=local")
        # Foxit Users
        self.foxit = pyad.adgroup.ADGroup.from_dn("CN=Foxit Users,OU=Citrix,OU=Groups,"
                                                  "DC=domain,DC=local")
        # Ops Department
        self.ops = pyad.adgroup.ADGroup.from_dn("CN=OPS,OU=RDS,OU=Groups,DC=domain,DC=local")

        # RDFilezilla
        self.filezilla = pyad.adgroup.ADGroup.from_dn("CN=RDFilezilla,OU=RDS,OU=Groups,"
                                                      "DC=domain,DC=local")
        # Ops_Department
        self.ops_fs = pyad.adgroup.ADGroup.from_dn("CN=FS-Ops_Department,OU=File Server,OU=Groups,"
                                                   "DC=domain,DC=local")
        # Sharefile
        self.sharefile = pyad.adgroup.ADGroup.from_dn("CN=Sharefile,OU=Citrix,OU=Groups,"
                                                      "DC=domain,DC=local")


"""
email what users were done and what preset was applied
Add groups to AD
https://stackoverflow.com/questions/65504041/add-group-membership-to-ad-with-pyad
"""