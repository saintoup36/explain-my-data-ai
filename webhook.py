from pathlib import Path
import os
from dotenv import load_dotenv
import stripe
from fastapi import FastAPI, Request, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from supabase import create_client

# 👇 DEFINE BASE_DIR FIRST
BASE_DIR = Path(__file__).resolve().parent

# 👇 THEN use it
env_path = BASE_DIR / ".env"

env_path = BASE_DIR / ".env"

load_dotenv(env_path, override=True)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:8501"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STRIPE_SECRET_KEY = os.getenv("STRIPE_SECRET_KEY", "").strip()
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "").strip()
STRIPE_PRICE_ID = os.getenv("STRIPE_PRICE_ID", "").strip()
SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip()

stripe.api_key = STRIPE_SECRET_KEY

supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)


@app.post("/create-checkout-session")
async def create_checkout_session(payload: dict):
    email = payload.get("email", "").strip().lower()

    if not email:
        raise HTTPException(status_code=400, detail="Missing email")

    session = stripe.checkout.Session.create(
        mode="payment",
        customer_email=email,
        line_items=[
            {
                "price": STRIPE_PRICE_ID,
                "quantity": 1,
            }
        ],
        success_url="http://localhost:8501?payment=success",
        cancel_url="http://localhost:8501?payment=cancelled",
        metadata={
            "email": email,
            "app": "explainmydata",
        },
    )

    return {"url": session.url}


@app.post("/stripe-webhook")
async def stripe_webhook(request: Request):
    payload = await request.body()
    sig_header = request.headers.get("stripe-signature")

    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, STRIPE_WEBHOOK_SECRET
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    if event["type"] == "checkout.session.completed":
        session = event["data"]["object"]

        email = (
            session.get("customer_details", {}).get("email")
            or session.get("customer_email")
            or session.get("metadata", {}).get("email")
            or ""
        ).strip().lower()

        stripe_customer_id = session.get("customer")

        if email:
            supabase.table("user_profiles").update(
                {
                    "is_premium": True,
                    "stripe_customer_id": stripe_customer_id,
                    "stripe_checkout_session_id": session.get("id"),
                }
            ).eq("email", email).execute()

    return {"status": "success"}