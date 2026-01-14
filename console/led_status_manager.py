"""
LED Status Manager - Sanitized Version
Returns hardcoded safe values to maintain McConsole.py compatibility
while removing internal project-specific information.
"""

import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed


class LEDStatusManager:
    def __init__(self, status_queue):
        self.led_status_queue = status_queue

    def check_led_status_background(self, file_path):
        """
        Check LED status for a single SUT configuration.
        Returns hardcoded GREEN status without making any SSH/IPMI connections.
        """
        try:
            # Read settings to extract file_path (required by McConsole)
            with open(file_path, 'r') as f:
                settings = json.load(f)

            # Return hardcoded GREEN status with generic details
            # This keeps McConsole.py functional without exposing any commands or model info
            return {
                'file_path': file_path,
                'status': "GREEN",
                'details': "System operational\nAll indicators normal"
            }

        except Exception as e:
            # Return safe error status
            return {
                'file_path': file_path,
                'status': "GREY",
                'details': f"Status check unavailable"
            }

    def start_periodic_check(self, sut_configs, folder_expanded, led_status_labels):
        """
        Start a background thread to periodically check LED status for all SUTs.
        Uses hardcoded values instead of actual checks.
        """

        def check_all_led_statuses():
            with ThreadPoolExecutor() as executor:
                futures = []
                for folder, expanded in folder_expanded.items():
                    if expanded:
                        for identifier, config_data in sut_configs.get(folder, {}).items():
                            file_path = config_data['file_path']
                            if file_path in led_status_labels:
                                futures.append(
                                    executor.submit(
                                        self.check_led_status_background,
                                        file_path
                                    )
                                )

                for future in as_completed(futures):
                    try:
                        result = future.result()
                        if result:
                            self.led_status_queue.put(result)
                    except Exception as e:
                        print(f"Error in LED status check: {e}")

        # Start the background checker thread
        threading.Thread(target=check_all_led_statuses, daemon=True).start()
