#!/usr/bin/env python3
"""
Minimal lib_wcs_rm.py for McConsole
Contains only the functions used by console components.
"""

import json
import logging
from pathlib import Path


def read_config_m(sut_ip):
    """
    Read SUT configuration file by IP or identifier.

    Searches recursively in the sut folder for settings.{sut_ip}.json
    and returns the configuration as a dictionary with FILE_NAME added.

    Args:
        sut_ip: IP address or identifier for the SUT (e.g., "192.168.1.100" or "server-name")

    Returns:
        dict: Configuration settings with FILE_NAME key added, or None on error

    Example:
        config = read_config_m("192.168.1.100")
        # Looks for: C:/mcqueen/sut/**/settings.192.168.1.100.json
    """
    try:
        # Determine the base path dynamically
        # Resolves to C:/mcqueen/sut/ when called from console/
        base_path = Path(__file__).resolve().parent.parent / "sut"

        # Recursively search for the settings file in the sut folder and its subfolders
        matching_files = list(base_path.rglob(f"settings.{sut_ip}.json"))

        if not matching_files:
            raise FileNotFoundError(f"Configuration file for IP {sut_ip} not found.")

        if len(matching_files) > 1:
            logging.warning(
                f"Multiple configuration files found for IP {sut_ip}. Using the first match: {matching_files[0]}"
            )

        settings_filename = matching_files[0]

        # Read the corresponding settings file
        with settings_filename.open("r") as file:
            settings = json.load(file)

        # Append the 'FILE_NAME' key with sut_ip value
        settings['FILE_NAME'] = sut_ip

        return settings

    except Exception as e:
        logging.error(f"Error reading SUT-specific configuration file for IP {sut_ip}: {e}")
        return None
