#!/usr/bin/env python3

from pathlib import Path
import argparse
import re
import pandas as pd
from ipaddress import ip_network
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

rows = []


def clean(value):
    if not value:
        return ""
    return value.strip().strip(";").strip('"').strip("'")


def get_hostname(content, file_path):
    match = re.search(r'host-name\s+([^;\n]+);', content)
    if match:
        return clean(match.group(1))
    return file_path.stem.split("_")[0]


def add_row(hostname, source, interface, unit, subnet_name, ip_address, subnet):
    rows.append({
        "Hostname": hostname,
        "Source": source,
        "Interface": interface,
        "Unit": unit,
        "Subnet Name": subnet_name,
        "Interface IP": ip_address,
        "Subnet": subnet
    })


def cidr_from_ip(ip):
    try:
        return str(ip_network(ip, strict=False))
    except ValueError:
        return ""


def extract_named_block(content, name):
    pattern = re.search(rf'\b{name}\s*\{{', content)
    if not pattern:
        return ""

    start = pattern.end()
    depth = 1
    i = start

    while i < len(content):
        if content[i] == "{":
            depth += 1
        elif content[i] == "}":
            depth -= 1
            if depth == 0:
                return content[start:i]
        i += 1

    return ""


def extract_child_blocks(block):
    children = []
    i = 0

    while i < len(block):
        match = re.search(r'([A-Za-z0-9_\-./:<>\[\]*]+)\s*\{', block[i:])
        if not match:
            break

        name = match.group(1)
        start = i + match.end()
        depth = 1
        j = start

        while j < len(block):
            if block[j] == "{":
                depth += 1
            elif block[j] == "}":
                depth -= 1
                if depth == 0:
                    children.append((name, block[start:j]))
                    i = j + 1
                    break
            j += 1
        else:
            break

    return children


def parse_set_format(content, hostname):
    descriptions = {}

    for match in re.finditer(
        r'^set\s+interfaces\s+(\S+)\s+unit\s+(\S+)\s+description\s+(.+)$',
        content,
        re.MULTILINE
    ):
        interface, unit, desc = match.groups()
        descriptions[(interface, unit)] = clean(desc)

    for match in re.finditer(
        r'^set\s+interfaces\s+(\S+)\s+unit\s+(\S+).*?\bfamily\s+inet\s+address\s+(\d+\.\d+\.\d+\.\d+/\d+)',
        content,
        re.MULTILINE
    ):
        interface, unit, ip_address = match.groups()
        subnet = cidr_from_ip(ip_address)

        if subnet:
            subnet_name = descriptions.get((interface, unit), f"{interface}.{unit}")

            add_row(
                hostname,
                "interfaces",
                interface,
                unit,
                subnet_name,
                ip_address,
                subnet
            )


def parse_hierarchical_interfaces(content, hostname):
    interfaces_block = extract_named_block(content, "interfaces")

    if not interfaces_block:
        return

    for interface, interface_block in extract_child_blocks(interfaces_block):
        for unit, unit_block in extract_child_blocks(interface_block):
            if unit != "unit":
                continue

        unit_matches = re.finditer(r'\bunit\s+([^\s{]+)\s*\{', interface_block)

        for unit_match in unit_matches:
            unit_id = unit_match.group(1)
            start = unit_match.end()
            depth = 1
            i = start

            while i < len(interface_block):
                if interface_block[i] == "{":
                    depth += 1
                elif interface_block[i] == "}":
                    depth -= 1
                    if depth == 0:
                        unit_block = interface_block[start:i]
                        break
                i += 1
            else:
                continue

            desc_match = re.search(
                r'\bdescription\s+("[^"]+"|[^;\n]+);',
                unit_block
            )

            subnet_name = clean(desc_match.group(1)) if desc_match else f"{interface}.{unit_id}"

            addresses = re.findall(
                r'\baddress\s+(\d+\.\d+\.\d+\.\d+/\d+)\b',
                unit_block
            )

            for ip_address in addresses:
                subnet = cidr_from_ip(ip_address)

                if subnet:
                    add_row(
                        hostname,
                        "interfaces",
                        interface,
                        unit_id,
                        subnet_name,
                        ip_address,
                        subnet
                    )


def parse_address_assignment_pools(content, hostname):
    access_block = extract_named_block(content, "access")

    if not access_block:
        return

    for pool_match in re.finditer(
        r'\b(?:inactive:\s*)?pool\s+([^\s{]+)\s*\{',
        access_block
    ):
        pool_name = clean(pool_match.group(1))
        start = pool_match.end()
        depth = 1
        i = start

        while i < len(access_block):
            if access_block[i] == "{":
                depth += 1
            elif access_block[i] == "}":
                depth -= 1
                if depth == 0:
                    pool_block = access_block[start:i]
                    break
            i += 1
        else:
            continue

        networks = re.findall(
            r'\bnetwork\s+(\d+\.\d+\.\d+\.\d+/\d+);',
            pool_block
        )

        for subnet in networks:
            add_row(
                hostname,
                "dhcp-pool",
                "",
                "",
                pool_name,
                "",
                subnet
            )


def parse_junos_file(file_path):
    content = file_path.read_text(errors="ignore")
    hostname = get_hostname(content, file_path)

    print(f"[+] Parsing {file_path.name}")

    parse_set_format(content, hostname)
    parse_hierarchical_interfaces(content, hostname)
    parse_address_assignment_pools(content, hostname)


def export_excel(output_file):
    df = pd.DataFrame(rows)

    if df.empty:
        print("[-] No JUNOS subnets found.")
        return

    df = df.drop_duplicates()
    df = df.sort_values([
        "Hostname",
        "Source",
        "Interface",
        "Unit",
        "Subnet"
    ])

    df.to_excel(output_file, index=False)

    wb = load_workbook(output_file)
    ws = wb.active
    ws.title = "JUNOS Subnets"

    for column in ws.columns:
        max_length = max(
            len(str(cell.value)) if cell.value else 0
            for cell in column
        )
        ws.column_dimensions[get_column_letter(column[0].column)].width = max_length + 4

    table_range = f"A1:G{ws.max_row}"

    table = Table(
        displayName="JunosSubnetTable",
        ref=table_range
    )

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    ws.add_table(table)
    ws.freeze_panes = "A2"

    wb.save(output_file)

    print(f"[+] Exported {len(df)} subnets to {output_file}")


def main():
    parser = argparse.ArgumentParser(
        description="Parse JUNOS config files and export subnet names/subnets to Excel"
    )

    parser.add_argument(
        "-i",
        "--input",
        required=True,
        help="Input JUNOS .conf file or folder"
    )

    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output Excel file"
    )

    args = parser.parse_args()

    input_path = Path(args.input)

    if not input_path.exists():
        print("[-] Input path does not exist.")
        return

    if input_path.is_file():
        parse_junos_file(input_path)

    elif input_path.is_dir():
        conf_files = sorted(input_path.glob("*.conf"))

        if not conf_files:
            print("[-] No .conf files found.")
            return

        for conf_file in conf_files:
            parse_junos_file(conf_file)

    export_excel(args.output)


if __name__ == "__main__":
    main()
