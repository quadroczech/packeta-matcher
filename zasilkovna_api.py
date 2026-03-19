import re
from zeep import Client
from zeep.exceptions import Fault

WSDL_URL = "https://www.zasilkovna.cz/api/soap.wsdl"

# Sdílený SOAP klient — načte WSDL jednou při startu aplikace
_client: Client | None = None

def get_client() -> Client:
    global _client
    if _client is None:
        _client = Client(WSDL_URL)
    return _client

def _parse_packet_id(tracking_number: str) -> int | None:
    """Extrahuje numerické ID zásilky z trasovacího čísla (např. Z1234567890 → 1234567890)."""
    if not tracking_number:
        return None
    cleaned = re.sub(r'^[Zz]', '', str(tracking_number).strip())
    try:
        return int(cleaned)
    except ValueError:
        return None

def get_packet_data(api_password: str, tracking_number: str) -> dict:
    """
    Vrátí slovník s číslem objednávky a hodnotou zásilky (bez DPH).
    Používá sdílený SOAP klient (WSDL se načte jen jednou).
    Vrátí: {"order_number": str|None, "value": float|None}
    """
    packet_id = _parse_packet_id(tracking_number)
    if packet_id is None:
        return {"order_number": None, "value": None}
    try:
        result = get_client().service.packetInfo(apiPassword=api_password, packetId=packet_id)
        if not result:
            return {"order_number": None, "value": None}
        order_number = result.number if hasattr(result, 'number') else None
        raw_value = result.value if hasattr(result, 'value') else None
        try:
            value = float(raw_value) if raw_value is not None else None
        except (TypeError, ValueError):
            value = None
        return {"order_number": order_number, "value": value}
    except (Fault, Exception):
        return {"order_number": None, "value": None}
