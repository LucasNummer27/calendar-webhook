from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import requests
from datetime import datetime, timedelta
import pytz
import os

app = FastAPI()

TENANT_ID = os.environ.get("TENANT_ID", "a33b09bf-1307-49ba-9562-49a85ade260c")
CLIENT_ID = os.environ.get("CLIENT_ID", "3661e427-d345-4bde-b2c9-1611072a09ae")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "LTp8Q~~JTPjrUCsx3_INxr5i.dVayIAvMwmM9cKN")
USER_EMAIL = os.environ.get("USER_EMAIL", "lucas.nummer@mindworkerllc.com")

def get_access_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
    )
    return resp.json().get("access_token")

@app.post("/check-availability")
async def check_availability(request: Request):
    token = get_access_token()
    if not token:
        return JSONResponse({"error": "Could not get access token"})
    headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.timezone="America/Toronto"'}
    eastern = pytz.timezone("America/Toronto")
    now = datetime.now(eastern)
    available_slots = []
    for day_offset in range(1, 8):
        check_date = now + timedelta(days=day_offset)
        if check_date.weekday() >= 5:
            continue
        if len(available_slots) >= 6:
            break
        start = check_date.replace(hour=0, minute=0, second=0, microsecond=0)
        end = check_date.replace(hour=23, minute=59, second=59, microsecond=0)
        utc_offset = start.strftime("%z")
        offset_formatted = f"{utc_offset[:3]}:{utc_offset[3:]}"
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/calendarView",
            headers=headers,
            params={"startDateTime": f"{start.strftime('%Y-%m-%dT%H:%M:%S')}{offset_formatted}",
                    "endDateTime": f"{end.strftime('%Y-%m-%dT%H:%M:%S')}{offset_formatted}",
                    "$select": "subject,start,end,showAs"}
        )
        if resp.status_code != 200:
            continue
        events = resp.json().get("value", [])
        busy_times = []
        for event in events:
            if event.get("showAs", "busy") == "free":
                continue
            try:
                es = datetime.strptime(event["start"]["dateTime"][:16], "%Y-%m-%dT%H:%M")
                ee = datetime.strptime(event["end"]["dateTime"][:16], "%Y-%m-%dT%H:%M")
            except ValueError:
                continue
            sm = es.hour * 60 + es.minute
            em = ee.hour * 60 + ee.minute
            if em <= sm:
                em = 24 * 60
            busy_times.append((sm, em))
        for hour in range(9, 17):
            for minute in [0, 30]:
                ss = hour * 60 + minute
                se = ss + 30
                is_free = all(not (ss < em and se > sm) for sm, em in busy_times)
                if is_free and len(available_slots) < 6:
                    available_slots.append({
                        "day": check_date.strftime("%A"),
                        "date": check_date.strftime("%Y-%m-%d"),
                        "time": f"{hour:02d}:{minute:02d}",
                        "display": f"{check_date.strftime('%A %B %d')} at {datetime(2026,1,1,hour,minute).strftime('%I:%M %p')} Eastern"
                    })
    return JSONResponse({"available_slots": available_slots,
                         "message": f"I have {len(available_slots)} available slots. Here are some options: " +
                                    ", ".join([s["display"] for s in available_slots[:3]])})

@app.post("/book-appointment")
async def book_appointment(request: Request):
    data = await request.json()
    date = data.get("date")
    time_str = data.get("time")
    prospect_name = data.get("prospect_name", "")
    prospect_email = data.get("prospect_email", "")
    business_name = data.get("business_name", "")
    if not date or not time_str:
        return JSONResponse({"error": "date and time are required"})
    token = get_access_token()
    if not token:
        return JSONResponse({"error": "Could not get access token"})
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    hour, minute = map(int, time_str.split(":"))
    em = minute + 15
    eh = hour + (1 if em >= 60 else 0)
    em = em - 60 if em >= 60 else em
    event_body = {"subject": f"Demo Call - {prospect_name} ({business_name})",
                  "start": {"dateTime": f"{date}T{time_str}:00", "timeZone": "America/Toronto"},
                  "end": {"dateTime": f"{date}T{eh:02d}:{em:02d}:00", "timeZone": "America/Toronto"},
                  "body": {"contentType": "HTML", "content": f"<p>Demo call with {prospect_name} from {business_name}.</p><p>Email: {prospect_email}</p>"},
                  "attendees": [{"emailAddress": {"address": prospect_email, "name": prospect_name}, "type": "required"}] if prospect_email else []}
    resp = requests.post(f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/events", headers=headers, json=event_body)
    if resp.status_code in [200, 201]:
        return JSONResponse({"status": "success", "message": f"Demo call booked for {date} at {time_str} Eastern with {prospect_name}. A calendar invite has been sent to {prospect_email}."})
    return JSONResponse({"error": f"Failed to create event: {resp.text[:200]}"})

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/")
async def root():
    return {"status": "ok", "service": "We Help Any Business - Calendar Webhooks"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))
