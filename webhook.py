import os
import stripe
from fastapi import FastAPI, Request, Header
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

load_dotenv()

app = FastAPI()

stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

engine = create_engine(DATABASE_URL, pool_pre_ping=True, future=True) if DATABASE_URL else None


def get_engine():
    if engine is None:
        raise ValueError("DATABASE_URL is missing in .env")
    return engine


def activate_premium_for_user(user_email):
    if not user_email:
        return

    db = get_engine()
    with db.begin() as conn:
        conn.execute(
            text("""
                UPDATE subscriptions
                SET
                    plan = 'premium',
                    status = 'active',
                    updated_at = NOW()
                WHERE user_email = :user_email
            """),
            {"user_email": user_email},
        )


def grant_onetime_access(user_email):
    if not user_email:
        return

    db = get_engine()
    with db.begin() as conn:
        conn.execute(
            text("""
                UPDATE subscriptions
                SET
                    plan = 'onetime',
                    status = 'active',
                    updated_at = NOW()
                WHERE user_email = :user_email
            """),
            {"user_email": user_email},
        )


@app.get("/")
def home():
    return {"status": "webhook server running"}


@app.post("/stripe-webhook")
async def stripe_webhook(request: Request, stripe_signature: str = Header(None)):
    payload = await request.body()

    try:
        event = stripe.Webhook.construct_event(
            payload,
            stripe_signature,
            WEBHOOK_SECRET
        )
    except Exception as e:
        return {"error": str(e)}

    if event["type"] == "checkout.session.completed":
        session = event["data"]["object"]

        customer_email = session.get("customer_email")
        amount_total = session.get("amount_total", 0)

        if amount_total == 1900:
            activate_premium_for_user(customer_email)

        elif amount_total == 2500:
            grant_onetime_access(customer_email)

    return {"status": "success"}