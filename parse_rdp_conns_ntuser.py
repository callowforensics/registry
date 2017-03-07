#!/usr/bin/env python
from Registry import Registry
import os
import sys
from collections import Counter

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

    # Get a unique list of RDP conns plus a frequency analysis of them.
    all_server_rdp_conns = [data[user]["server"][i][0].lower() for user in data for i in range(len(data[user]["server"]))]
    all_default_rdp_conns = [data[user]["default"][i].lower() for user in data for i in range(len(data[user]["default"]))]
    all_rdp_conns = dict(Counter(list(all_default_rdp_conns + all_server_rdp_conns)))

    # Write the results.
    with open(os.path.join(out_path, "{}_rdp_analysis.txt".format(system_in_focus)), "w") as f:
        summary_header = "The following {} systems have been RDPd to from the system " \
                         "\"{}\" (frequency of key occurrence in brackets)\n"\
            .format(len(all_rdp_conns), system_in_focus)
        f.write("/" * len(summary_header.strip()) + "\n")
        f.write(summary_header)
        f.write("/" * len(summary_header.strip()) + "\n\n")

        for system, frequency in sorted(dict.items(all_rdp_conns)):
            f.write(system + " ({})".format(str(frequency)) + "\n")

        summary_header_2 = "Breakdown of RDP connections per user\n".upper()
        f.write("\n" + "/" * len(summary_header_2) + "\n")
        f.write(summary_header_2)
        f.write("/" * len(summary_header_2) + "\n\n")

        for user in sorted(data):
            if data[user]["server"] or data[user]["default"]:
                f.write("User \"{}\":\n\n".format(user))

                if data[user]["server"]:
                    focus = data[user]["server"]
                    f.write("{} system(s) found within the key "
                            "\"Software\Microsoft\Terminal Server Client\Servers\":\n\n"
                            .format(len(focus)))
                    for index, system in enumerate(sorted(focus), 1):
                        host, hint, time = system
                        f.write("{})".format(str(index)))
                        f.write("\nHost: {}".format(host))
                        f.write("\nKey last written time: {}".format(str(time)[:19]))
                        f.write("\nPassword hint: {}\n".format(hint))

                if data[user]["default"]:
                    focus = sorted(list(set(data[user]["default"])))
                    if data[user]["server"]:
                        f.write("\n")
                    f.write("{} system(s) found within the key "
                            "\"Software\Microsoft\Terminal Server Client\Default\":\n\n"
                            .format(len(focus)))
                    for index, system in enumerate(focus, 1):
                        f.write(str(index) + ") " + system + "\n")

                f.write("-" * 100 + "\n")
