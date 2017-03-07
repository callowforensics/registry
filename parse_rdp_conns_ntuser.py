#!/usr/bin/env python
from Registry import Registry
import os
import sys

if __name__ == "__main__":
    # Check that the user has enter enough args.
    if len(sys.argv) != 4:
        print("Loops through a given directory or directories and parses out RDP connections.")
        print("Usage: input out_dir system_name")
        print("\nExample:")
        print("-" * 50)
        print(r"c:\exported_ntuser_files c:\results testsys123")
        sys.exit()

    in_path = sys.argv[1]
    out_path = sys.argv[2]
    system_in_focus = sys.argv[3]

    if not os.path.exists(out_path):
        os.makedirs(out_path)

    # Set a dictionary to store the results.
    data = {}
    for path, dirs, files in os.walk(in_path):
        for filename in files:
            fullpath = os.path.join(path, filename)
            if out_path not in fullpath:
                reg = Registry.Registry(fullpath)
                # Default if the user name can't be found (see below).
                user_name = "Unknown"
                # noinspection PyBroadException
                try:
                    # Try to get the user name from the NTUSER.DAT file.
                    get_user = reg.open("Software\Microsoft\Windows\CurrentVersion\Explorer\shell Folders")
                    for value in get_user.values():
                        if value.name().lower() == "desktop":
                            user_name = value.value().split("\\")[2]
                            if user_name not in data:
                                data[user_name] = {}
                                data[user_name]["server"] = []
                                data[user_name]["default"] = []
                except Exception:
                    pass

                # noinspection PyBroadException
                try:
                    # Get the server key entries.
                    ts_key_servers = reg.open("Software\Microsoft\Terminal Server Client\Servers")
                    for system in ts_key_servers.subkeys():
                        # noinspection PyBroadException
                        try:
                            pwhint = system.value("UserNameHint").value()
                        except:
                            pwhint = "-"
                        # We split the key name as sometimes it can have /console etc. in it.
                        data[user_name]["server"].append((system.name().split()[0], pwhint, system.timestamp()))
                except Exception:
                    pass

                # noinspection PyBroadException
                try:
                    # Get the default key entries.
                    ts_key_default = reg.open("Software\Microsoft\Terminal Server Client\Default")
                    for name in ts_key_default.values():
                        data[user_name]["default"].append(name.value().split()[0])
                except Exception:
                    pass

    # Get a unique list of RDP conns.
    all_server_rdp_conns = sorted(list(set(data[user]["server"][i][0]
                                           for user in data for i in range(len(data[user]["server"])))))
    all_default_rdp_conns = sorted(list(set(data[user]["default"][i]
                                            for user in data for i in range(len(data[user]["default"])))))
    all_rdp_conns = sorted(list(set(all_default_rdp_conns + all_server_rdp_conns)))

    # Write the results.
    with open(os.path.join(out_path, "results.txt"), "w") as f:
        summary_header = "The following {} RDP connections have been made from the system: \"{}\":\n"\
            .format(len(all_rdp_conns), system_in_focus)

        f.write(summary_header)
        f.write("#" * len(summary_header.strip()) + "\n\n")

        for index, system in enumerate(all_rdp_conns, 1):
            f.write(str(index) + ") " + system + "\n")

        summary_header_2 = "\nBreakdown of RDP connections per user:\n"
        f.write(summary_header_2)
        f.write("#" * len(summary_header_2.strip()) + "\n\n")

        for user in data:
            if data[user]["server"] or data[user]["default"]:
                f.write("User \"{}\":\n\n".format(user))

                if data[user]["server"]:
                    f.write("Entries found within \"Software\Microsoft\Terminal Server Client\Servers\":\n\n")
                    for index, system in enumerate(sorted(data[user]["server"], 1)):
                        host, hint, time = system
                        f.write("{})".format(str(index)))
                        f.write("\nHost: {}".format(host))
                        f.write("\nKey last written time: {}".format(str(time)[:19]))
                        f.write("\nPassword hint: {}\n".format(hint))

                if data[user]["default"]:
                    f.write("Entries found within \"Software\Microsoft\Terminal Server Client\Default\":\n\n")
                    for index, system in enumerate(sorted(data[user]["default"]), 1):
                        f.write(str(index) + ") " + system + "\n")

                f.write("-" * 100 + "\n")
