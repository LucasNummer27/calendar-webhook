from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import requests
from datetime import datetime, timedelta
import pytz
import os

app = FastAPI()

# Microsoft Graph credentials
TENANT_ID = os.environ.get("TENANT_ID", "a33b09bf-1307-49ba-9562-49a85ade260c")
CLIENT_ID = os.environ.get("CLIENT_ID", "3661e427-d345-4bde-b2c9-1611072a09ae")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "LTp8Q~~JTPjrUCsx3_INxr5i.dVayIAvMwmM9cKN")
USER_EMAIL = os.environ.get("USER_EMAIL", "lucas.nummer@mindworkerllc.com")

def get_access_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
    )
    return resp.json().get("access_token")

@app.post("/check-availability")
async def check_availability(request: Request):
    token = get_access_token()
    if not token:
        return JSONResponse({"error": "Could not get access token"})
    
    # Use Prefer header to get times directly in Eastern timezone
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.timezone="America/Toronto"'
    }
    eastern = pytz.timezone("America/Toronto")
    now = datetime.now(eastern)
    
    available_slots = []
    
    # Check next 5 business days
    for day_offset in range(1, 8):
        check_date = now + timedelta(days=day_offset)
        if check_date.weekday() >= 5:  # Skip weekends
            continue
        if len(available_slots) >= 6:  # Return max 6 slots
            break
            
        # Use Eastern time for the query range
        start = check_date.replace(hour=0, minute=0, second=0, microsecond=0)
        end = check_date.replace(hour=23, minute=59, second=59, microsecond=0)
        
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/calendarView",
            headers=headers,
            params={
                "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%S"),
                "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%S"),
                "$select": "subject,start,end,showAs"
            }
        )
        
        if resp.status_code != 200:
            continue
            
        events = resp.json().get("value", [])
        
        # Find free slots between 9 AM and 5 PM Eastern
        # Times from Graph are now already in Eastern (thanks to Prefer header)
        busy_times = []
        for event in events:
            # Only block if showAs is busy, oof, or tentative
            show_as = event.get("showAs", "busy")
            if show_as == "free":
                continue
                
            # Parse the time directly as Eastern (no conversion needed)
            event_start_str = event["start"]["dateTime"][:16]  # "2026-05-08T10:00"
            event_end_str = event["end"]["dateTime"][:16]
            
            try:
                event_start_dt = datetime.strptime(event_start_str, "%Y-%m-%dT%H:%M")
                event_end_dt = datetime.strptime(event_end_str, "%Y-%m-%dT%H:%M")
            except ValueError:
                continue
            
            start_minutes = event_start_dt.hour * 60 + event_start_dt.minute
            end_minutes = event_end_dt.hour * 60 + event_end_dt.minute
            
            # Handle events that span midnight or are all-day
            if end_minutes <= start_minutes:
                end_minutes = 24 * 60  # treat as end of day
            
            busy_times.append((start_minutes, end_minutes))
        
        # Check each 30-min slot from 9 AM to 5 PM
        for hour in range(9, 17):
            for minute in [0, 30]:
                slot_start = hour * 60 + minute
                slot_end = slot_start + 30
                
                # Check if slot conflicts with any busy time
                is_free = True
                for busy_start, busy_end in busy_times:
                    if slot_start < busy_end and slot_end > busy_start:
                        is_free = False
                        break
                
                if is_free and len(available_slots) < 6:
                    day_name = check_date.strftime("%A")
                    date_str = check_date.strftime("%Y-%m-%d")
                    time_str = f"{hour:02d}:{minute:02d}"
                    display_time = datetime(2026, 1, 1, hour, minute).strftime("%I:%M %p")
                    available_slots.append({
                        "day": day_name,
                        "date": date_str,
                        "time": time_str,
                        "display": f"{day_name} {check_date.strftime('%B %d')} at {display_time} Eastern"
                    })
    
    return JSONResponse({
        "available_slots": available_slots,
        "message": f"I have {len(available_slots)} available slots. Here are some options: " + 
                   ", ".join([s["display"] for s in available_slots[:3]])
    })

@app.post("/book-appointment")
async def book_appointment(request: Request):
    data = await request.json()
    date = data.get("date")  # YYYY-MM-DD
    time_str = data.get("time")  # HH:MM
    prospect_name = data.get("prospect_name", "")
    prospect_email = data.get("prospect_email", "")
    business_name = data.get("business_name", "")
    
    if not date or not time_str:
        return JSONResponse({"error": "date and time are required"})
    
    token = get_access_token()
    if not token:
        return JSONResponse({"error": "Could not get access token"})
    
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # Parse date and time
    start_dt = f"{date}T{time_str}:00"
    # End time = start + 15 minutes
    hour, minute = map(int, time_str.split(":"))
    end_minute = minute + 15
    end_hour = hour
    if end_minute >= 60:
        end_minute -= 60
        end_hour += 1
    end_dt = f"{date}T{end_hour:02d}:{end_minute:02d}:00"
    
    event_body = {
        "subject": f"Demo Call - {prospect_name} ({business_name})",
        "start": {"dateTime": start_dt, "timeZone": "America/Toronto"},
        "end": {"dateTime": end_dt, "timeZone": "America/Toronto"},
        "body": {
            "contentType": "HTML",
            "content": f"<p>Demo call with {prospect_name} from {business_name}.</p><p>Email: {prospect_email}</p><p>Booked by AI Agent Alex.</p>"
        },
        "attendees": []
    }
    
    if prospect_email:
        event_body["attendees"].append({
            "emailAddress": {"address": prospect_email, "name": prospect_name},
            "type": "required"
        })
    
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/events",
        headers=headers,
        json=event_body
    )
    
    if resp.status_code in [200, 201]:
        return JSONResponse({
            "status": "success",
            "message": f"Demo call booked for {date} at {time_str} Eastern with {prospect_name}. A calendar invite has been sent to {prospect_email}."
        })
    else:
        return JSONResponse({"error": f"Failed to create event: {resp.text[:200]}"})

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/")
async def root():
    return {"status": "ok", "service": "We Help Any Business - Calendar Webhooks"}

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "10000"))
    uvicorn.run(app, host="0.0.0.0", port=port)

