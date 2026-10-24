import os
import uuid
from datetime import datetime
import phonenumbers


def generate_message_id():
    return "AS" + uuid.uuid4().hex[:10].upper()


def send_message(phone_number, message, country_code="IN"):
    """Validate and prepare a message. Default mode is safe local simulation."""
    if not phone_number or not phone_number.strip():
        return {"status": "failed", "error": "Phone number is required."}
    if not message or not message.strip():
        return {"status": "failed", "error": "Message cannot be empty."}

    try:
        parsed_number = phonenumbers.parse(phone_number.strip(), (country_code or "IN").strip().upper())
        if not phonenumbers.is_valid_number(parsed_number):
            return {"status": "failed", "error": "Please enter a valid phone number."}
        international_number = phonenumbers.format_number(
            parsed_number, phonenumbers.PhoneNumberFormat.E164
        )
    except phonenumbers.NumberParseException:
        return {"status": "failed", "error": "Invalid phone number format."}

    provider_mode = os.getenv("CPaaS_PROVIDER", "local").strip().lower()
    if provider_mode not in {"local", "twilio"}:
        provider_mode = "local"

    # No provider credentials or network request are used in local mode.
    # The Twilio mode is intentionally a readiness/configuration placeholder:
    # a real provider adapter must be configured and tested before live sending.
    if provider_mode == "twilio":
        required = ("TWILIO_ACCOUNT_SID", "TWILIO_AUTH_TOKEN", "TWILIO_FROM_NUMBER")
        missing = [name for name in required if not os.getenv(name)]
        if missing:
            return {
                "status": "failed",
                "error": "Twilio mode is selected but credentials are missing: " + ", ".join(missing),
            }
        return {
            "status": "failed",
            "error": "Provider credentials are present, but live sending is not enabled in this demo build. Use local mode until a provider adapter is configured.",
        }

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "message_id": generate_message_id(),
        "phone_number": international_number,
        "message": message.strip(),
        "status": "queued",
        "created_at": now,
        "updated_at": now,
        "provider": "local-simulation",
    }


def create_order_message(order_id, amount):
    return (
        f"Your order {order_id} has been received successfully. "
        f"Order amount: ₹{amount:.2f}. "
        "Thank you for shopping with us!"
    )
