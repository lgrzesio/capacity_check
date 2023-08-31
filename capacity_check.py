from getpass import getpass
from threading import Thread
from queue import Queue
from collections import defaultdict
from jnpr.junos import Device
from xlsxwriter import Workbook
from xlsxwriter.workbook import Worksheet
from zipfile import ZipFile
from lxml import etree
from jnpr.junos.exception import ConnectTimeoutError
import os
import re
import yaml


username = input("Username: ")
password = getpass()


class Capacity:
    def __init__(
        self, hosts_file="hosts.yaml", output_file="bandwidth_data.xlsx", log_rpc=True
    ):
        self.HOSTS_FILE = hosts_file
        self.DATA_FILENAME = output_file
        self.HEADERS = [
            "Hostname",
            "Version",
            "Hardware",
            "Model",
            "Serial",
            "Speed",
            "Channelized Ports",
            "Ports In Use",
            "Ports Installed",
            "Calculated Capacity",
            "Used (gbps)",
            "Available (gbps)",
            "Remaining (gbps)",
        ]
        self.workbook = None
        self.worksheet = None
        self.entry_row = 1
        self.log_rpc = log_rpc
        self.queue = Queue()

    def initialize_workbook(self):
        self.workbook = Workbook(self.DATA_FILENAME)
        self.worksheet = self.workbook.add_worksheet()
        bold = self.workbook.add_format({"bold": True})

        for col, header in enumerate(self.HEADERS):
            self.worksheet.write(0, col, header, bold)

        self.worksheet.autofilter(0, 0, 0, 6)

        return (self.workbook, self.worksheet)

    def get_devices(self):
        with open("hosts.yaml") as hosts:
            try:
                host_list = yaml.safe_load(hosts)
                host_list = host_list["hosts"]
            except yaml.YAMLError as e:
                print(e)
                host_list = []

        return host_list

    def get_capacity_usage(self):
        device_list = self.get_devices()

        workbook, worksheet = self.initialize_workbook()
        writer_thread = Thread(target=self.write_data, args=(), daemon=True)
        writer_thread.start()

        devices = [
            JunosDevice(device_name, self.queue, self.log_rpc)
            for device_name in device_list
        ]
        threads = [Thread(target=device.check_bandwidth, args=()) for device in devices]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        self.queue.put(None)
        self.queue.join()

        if self.log_rpc:
            with ZipFile(f"rpc_logs.zip", "w") as zip_file:
                missing_devices = []
                for device_name in device_list:
                    try:
                        zip_file.write(f"{device_name}_interface_config.xml")
                        zip_file.write(f"{device_name}_interface_media.xml")
                        zip_file.write(f"{device_name}_chassis_hardware.xml")
                        os.remove(f"{device_name}_interface_config.xml")
                        os.remove(f"{device_name}_interface_media.xml")
                        os.remove(f"{device_name}_chassis_hardware.xml")

                    except FileNotFoundError as e:
                        # Add missing devices to separate file
                        missing_devices.append(device_name)
                with open("missing_devices.txt", "w") as missing:
                    missing.write("\n".join(missing_devices))
                zip_file.write("missing_devices.txt")

        worksheet.autofit()
        workbook.close()

    def write_data(self):
        while True:
            entry = self.queue.get()
            if entry is None:
                self.queue.task_done()
                break

            col = 0
            for value in entry["chassis"]:
                self.worksheet.write(self.entry_row, col, value)
                col += 1
            self.entry_row += 1

            for _, info in entry["linecards"].items():
                # Iterate over linecard details first
                col = 0
                for value in info["details"]:
                    self.worksheet.write(self.entry_row, col, value)
                    col += 1
                self.entry_row += 1

                # Iterate over each PIC
                for pic, xcvrs in info["interfaces"].items():
                    for xcvr_name, channel_details in xcvrs.items():
                        for channel, xcvr_details in channel_details.items():
                            name = f"{pic} / {xcvr_name}"
                            if channel != "default":
                                name = f"{pic} / {xcvr_name}:{channel}"
                            model = xcvr_details["model"]
                            serial = xcvr_details["serial"]
                            speed = xcvr_details["speed"]

                            self.worksheet.write(
                                self.entry_row, 0, entry["chassis"][0]
                            )  # host
                            self.worksheet.write(self.entry_row, 2, name)
                            self.worksheet.write(self.entry_row, 3, model)
                            self.worksheet.write(self.entry_row, 4, serial)
                            self.worksheet.write(self.entry_row, 5, speed)
                            self.entry_row += 1

            self.queue.task_done()


class JunosDevice:
    def __init__(self, device_name, queue, log_rpc):
        self.device_name = device_name
        self.queue = queue
        self.device = Device(host=device_name, user=username, password=password)
        self.log_rpc = log_rpc

        self.NAG_VERSION = "22.2R1"
        self.EVO_NAG_VERSION = "21.1"
        self.INTERFACE_PREFIXES = ["et", "ge", "xe", "xle", "fte"]  # ,"ae"]

    def connect(self):
        try:
            self.device.open()
            return True
        except ConnectTimeoutError as e:
            print(f"Could not connect to {self.device_name}")
            return False

    def disconnect(self):
        self.device.close()

    def __enter__(self):
        self.device.open()

    def __exit__(self, type, value, traceback):
        self.device.close()

    def _validate_prefix(self, name):
        for prefix in self.INTERFACE_PREFIXES:
            if name.startswith(prefix):
                return True
        return False

    def _get_field(self, xml_data, name, default_value=None):
        value = getattr(xml_data.find(name), "text", default_value)
        if value is not None:
            value = value.strip()
        return value

    def _get_layer(self, layer_config):
        if (
            "ethernet-switching interface-mode trunk" in layer_config
            and "ethernet-switching vlan members" in layer_config
        ) or ("family mpls" in layer_config):
            return 2

        if (
            "ether-options 802.3ad" in layer_config
            or "family inet address" in layer_config
            or "family inet6 address" in layer_config
        ):
            return 3

        return None

    def get_linecard_capacity(self, linecard_interfaces, linecard_details):
        """
        print("Checking the following linecard interfaces:")
        pprint(linecard_interfaces)
        print("-" * 10)
        """
        channelized_ports = 0
        ports_in_use = set()
        linecard_capacity = 0

        for interface, info in linecard_interfaces.items():
            """
            Interface needs to be configured as Up (operational status does not seem to matter)
            and an address must exist on at least one of the subunits
            """
            speed = "N/A"
            channel = "default"
            _, pic, slot = interface.split("-")[1].split("/")
            # check if we're dealing with channelized interfaces
            if ":" in slot:
                slot, channel = slot.split(":")
                channelized_ports += 1
            if info["admin_status"] == "up" and (
                info["layer"] == 2
                or (info["layer"] == 3 and info["address"] is not None)
            ):
                ports_in_use.add(f"{pic}/{slot}")
                speed = int(re.findall(r"\d+", info["speed"])[0])
                linecard_capacity += speed

            # copy over the model / serial information
            if channel != "default":
                if (
                    channel
                    not in linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"]
                ):
                    linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"][
                        channel
                    ] = {}
                main_channel = linecard_details["interfaces"][f"PIC {pic}"][
                    f"Xcvr {slot}"
                ]["default"]
                linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"][channel]
                linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"][channel][
                    "model"
                ] = main_channel["model"]
                linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"][channel][
                    "serial"
                ] = main_channel["serial"]

            linecard_details["interfaces"][f"PIC {pic}"][f"Xcvr {slot}"][channel][
                "speed"
            ] = speed

        return (channelized_ports, len(ports_in_use), linecard_capacity)

    def get_address(self, interface_info, physical_interface):
        for logical_interface in interface_info.xpath(
            f".//physical-interface/logical-interface[starts-with(name,'{physical_interface}')]"
        ):
            name = self._get_field(logical_interface, "name")

            # check to see if the interface belongs to an AE bundle
            ae_bundle = self._get_field(
                logical_interface, ".//address-family/ae-bundle-name"
            )

            if ae_bundle is not None:
                address = self.get_address(interface_info, ae_bundle)
            else:
                address = self._get_field(
                    logical_interface, ".//address-family/interface-address/ifa-local"
                )
            if address is not None:
                return name, address
        return None, None

    def _is_active_configuration(self, config, deactivate_config):
        for dc in deactivate_config:
            if dc[len("deactivate ") :] in config:
                return False
        return True

    def get_interface_configuration(self):
        interface_config = self.device.rpc.get_config(
            filter_xml="interfaces", options={"format": "set"}
        )
        # Log the RPC output for later review
        if self.log_rpc:
            with open(f"{self.device_name}_interface_config.xml", "w+") as rpc_log:
                rpc_log.write(
                    etree.tostring(interface_config, method="xml", encoding="unicode")
                )

        interface_config = interface_config.findall(".")[0].text.splitlines()

        # Parse out any deactivated statements
        deactivate_config = [
            config for config in interface_config if config.startswith("deactivate")
        ]
        interface_config = [
            config for config in interface_config if config not in deactivate_config
        ]

        new_config = [
            config
            for config in interface_config
            if self._is_active_configuration(config, deactivate_config)
        ]

        return "\n".join(new_config)

    def get_interface_details(self, interface_config):
        interface_details = {}
        ae_members = defaultdict(list)
        interface_info = self.device.rpc.get_interface_information(
            media=True, detail=True, normalize=True
        )
        if self.log_rpc:
            # Log the RPC output for later review
            with open(f"{self.device_name}_interface_media.xml", "w+") as rpc_log:
                rpc_log.write(
                    etree.tostring(interface_info, method="xml", encoding="unicode")
                )

        for physical_interface in interface_info.xpath(".//physical-interface"):
            name = self._get_field(physical_interface, "name")

            if not self._validate_prefix(name):
                continue

            admin_status = self._get_field(physical_interface, "admin-status")
            oper_status = self._get_field(physical_interface, "oper-status")
            speed = self._get_field(physical_interface, "speed", "0")
            logical_interface = None
            address = None
            ae = None

            # determine if it's configured as layer 2 or 3
            layer_config = "\n".join(
                [config for config in interface_config.splitlines() if name in config]
            )
            layer = self._get_layer(layer_config)

            if layer == 3:
                logical_interface, address = self.get_address(interface_info, name)

            if type(address) is tuple:
                ae, address = address
                ae_members[ae].append(name)

            interface_details[name] = {
                "admin_status": admin_status.lower(),
                "oper_status": oper_status.lower(),
                "speed": speed,
                "logical_interface": logical_interface,
                "layer": layer,
                "address": address,
                "ae": ae,
            }

        # in case of mixed optics on an AE make sure to override the speed of all AE members to reflect the lowest speed
        for details in interface_details.values():
            ae = details["ae"]
            if ae is not None:
                min_speed = 10000
                for link in ae_members[ae]:
                    speed = interface_details[link]["speed"]
                    speed = int(re.findall(r"\d+", speed)[0])
                    if speed < min_speed:
                        min_speed = speed
                for link in ae_members[ae]:
                    interface_details[link]["speed"] = f"{min_speed}Gbps"

        return interface_details

    def get_license_usage(self):
        license_summary = self.device.rpc.get_license_summary_information(
            normalize=True
        )
        feature_summary_path = (
            './/feature-summary[description="Port Bandwidth Usage (PAYG license)"]'
        )
        used_license = int(
            self._get_field(
                license_summary, f"{feature_summary_path}/used-licensed", "0"
            )
        )
        available_license = int(
            self._get_field(license_summary, f"{feature_summary_path}/licensed", "0")
        )

        return (used_license, available_license)

    def get_chassis_info(self):
        chassis_details = {}

        chassis_hardware = self.device.rpc.get_chassis_inventory(normalize=True)
        # Log the RPC output for later review
        if self.log_rpc:
            with open(f"{self.device_name}_chassis_hardware.xml", "w+") as rpc_log:
                rpc_log.write(
                    etree.tostring(chassis_hardware, method="xml", encoding="unicode")
                )

        chassis_info = chassis_hardware.xpath(".//chassis")[0]
        chassis_model = self._get_field(chassis_info, "description")
        chassis_serial = self._get_field(chassis_info, "serial-number")

        chassis_details["model"] = chassis_model
        chassis_details["serial"] = chassis_serial
        chassis_details["linecards"] = {}

        fpc_xpaths = [
            ".//chassis-module",
            ".//chassis-sub-module",
            ".//chassis-sub-sub-module",
        ]
        for fpc_xpath in fpc_xpaths:
            for module in chassis_hardware.xpath(fpc_xpath):
                name = self._get_field(module, "name")
                if name.startswith("FPC"):
                    serial = self._get_field(module, "serial-number")
                    model = self._get_field(module, "description")
                    version = self._get_field(module, "version", "")

                    pics = defaultdict(dict)
                    for sub_module in chassis_hardware.xpath(
                        f"{fpc_xpath}[name='{name}']//name"
                    ):
                        sub_module_name = self._get_field(
                            sub_module.getparent(), "name", "N/A"
                        )
                        if sub_module_name.startswith("Xcvr"):
                            parent = sub_module.getparent()
                            port_name = self._get_field(parent, "name", "N/A")
                            port_model = self._get_field(parent, "description", "N/A")
                            port_serial = self._get_field(
                                parent, "serial-number", "N/A"
                            )
                            pic_name = self._get_field(
                                parent.getparent(), "name", "N/A"
                            )

                            # assign 0 as a port channel by default
                            pics[pic_name][port_name] = {
                                "default": {
                                    "model": port_model,
                                    "serial": port_serial,
                                    "speed": "N/A",
                                }
                            }

                    chassis_details["linecards"][name] = {
                        "model": model,
                        "serial": serial,
                        "version": version,
                        "ports_installed": sum(
                            [len(ports) for _, ports in pics.items()]
                        ),
                        "interfaces": pics,
                    }

        return chassis_details

    def check_license_usage(self):
        used, available = self.get_license_usage()
        remaining = available - used

        return used, available, remaining

    def check_bandwidth(self):
        if not self.connect():
            return

        # check version to see if license nag is present
        version = self.device.facts["version"]
        is_evo = "EVO" in version.upper()

        # if we are on a version before the license nag, let's manually check port usage
        if (is_evo and version < self.EVO_NAG_VERSION) or (
            not is_evo and version < self.NAG_VERSION
        ):
            used, available, remaining = (0, 0, 0)
        else:
            used, available, remaining = self.check_license_usage()
            # Compare this to the manual calculations to determine any discrepenacies

        interface_config = self.get_interface_configuration()
        chassis_details = self.get_chassis_info()
        interface_details = self.get_interface_details(interface_config)

        total_ports = 0
        total_installed = 0
        total_capacity = 0

        entry = {
            "chassis": [
                self.device_name,
                version,
                "Chassis",
                chassis_details["model"],
                chassis_details["serial"],
                "",
                "",
            ],
            "linecards": {},
        }

        for linecard, linecard_details in chassis_details["linecards"].items():
            linecard_interfaces = {}
            for interface, details in interface_details.items():
                prefix, slot_info = interface.split("-")
                fpc, pic, xcvr = slot_info.split("/")
                if ":" in xcvr:
                    xcvr, channel = xcvr.split(":")
                pic_key = f"PIC {pic}"
                xcvr_key = f"Xcvr {xcvr}"
                if (
                    f"-{linecard.split(' ')[1]}/" in interface
                    and pic_key in linecard_details["interfaces"]
                    and xcvr_key in linecard_details["interfaces"][pic_key]
                ):
                    linecard_interfaces[interface] = details

            (
                channelized_ports,
                ports_in_use,
                linecard_capacity,
            ) = self.get_linecard_capacity(linecard_interfaces, linecard_details)

            entry["linecards"][linecard] = {
                "details": [
                    self.device_name,
                    linecard_details["version"],
                    linecard,
                    linecard_details["model"],
                    linecard_details["serial"],
                    "",
                    channelized_ports,
                    ports_in_use,
                    linecard_details["ports_installed"],
                    linecard_capacity,
                ],
                "interfaces": linecard_details["interfaces"],
            }
            total_ports += ports_in_use
            total_installed += linecard_details["ports_installed"]
            total_capacity += linecard_capacity

        entry["chassis"].append(total_ports)
        entry["chassis"].append(total_installed)
        entry["chassis"].append(total_capacity)
        entry["chassis"].append(used)
        entry["chassis"].append(available)
        entry["chassis"].append(remaining)

        self.queue.put(entry)
        self.disconnect()


def main():
    capacity = Capacity(
        hosts_file="hosts.yaml", output_file="bandwidth_data.xlsx", log_rpc=True
    )
    capacity.get_capacity_usage()


if __name__ == "__main__":
    main()
