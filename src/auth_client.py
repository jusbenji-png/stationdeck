import requests
import logging

logger = logging.getLogger(__name__)

AUTH_SERVER = "https://web-production-46077.up.railway.app"
TIMEOUT = 10  # seconds before giving up on a request


def register_station(station_name, password, email, phone,
                     region, location, machine_id, app_station_id=None):
    """
    Register a new station with the auth server.
    Returns: (success: bool, message: str, license_key: str or None)
    """
    try:
        response = requests.post(
            f"{AUTH_SERVER}/register",
            json={
                "station_name": station_name,
                "password":     password,
                "email":        email,
                "phone":        phone,
                "region":          region,
                "location":        location,
                "machine_id":      machine_id,
                "app_station_id":  app_station_id or station_name,
            },
            timeout=TIMEOUT
        )
        data = response.json()

        if response.status_code == 201 and data.get("success"):
            return True, data.get("message", "Registration successful."), data.get("license_key")
        else:
            return False, data.get("error", "Registration failed."), None

    except requests.exceptions.ConnectionError:
        return False, "Cannot reach StationDeck server. Check your internet connection.", None
    except requests.exceptions.Timeout:
        return False, "Server took too long to respond. Try again.", None
    except Exception as e:
        logger.error(f"Registration error: {e}")
        return False, "An unexpected error occurred.", None


def verify_station(station_name, password, machine_id):
    """
    Verify station identity on login.
    Returns: (success: bool, message: str, data: dict or None)
    """
    try:
        response = requests.post(
            f"{AUTH_SERVER}/verify",
            json={
                "station_name": station_name,
                "password":     password,
                "machine_id":   machine_id
            },
            timeout=TIMEOUT
        )
        data = response.json()

        if response.status_code == 200 and data.get("success"):
            return True, "Login successful.", data
        else:
            return False, data.get("error", "Verification failed."), None

    except requests.exceptions.ConnectionError:
        return False, "Cannot reach StationDeck server. Check your internet connection.", None
    except requests.exceptions.Timeout:
        return False, "Server took too long to respond. Try again.", None
    except Exception as e:
        logger.error(f"Verification error: {e}")
        return False, "An unexpected error occurred.", None


def recover_station(station_name, password, phone, new_machine_id):
    """
    Self-service machine recovery.
    Returns: (success: bool, message: str, license_key: str or None)
    """
    try:
        response = requests.post(
            f"{AUTH_SERVER}/recover",
            json={
                "station_name":   station_name,
                "password":       password,
                "phone":          phone,
                "new_machine_id": new_machine_id
            },
            timeout=TIMEOUT
        )
        data = response.json()

        if response.status_code == 200 and data.get("success"):
            return True, data.get("message", "Recovery successful."), data.get("license_key")
        else:
            return False, data.get("error", "Recovery failed."), None

    except requests.exceptions.ConnectionError:
        return False, "Cannot reach StationDeck server. Check your internet connection.", None
    except requests.exceptions.Timeout:
        return False, "Server took too long to respond. Try again.", None
    except Exception as e:
        logger.error(f"Recovery error: {e}")
        return False, "An unexpected error occurred.", None