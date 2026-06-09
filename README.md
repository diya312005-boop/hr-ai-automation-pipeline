# HR AI Automation Pipeline

## What this project does
An end-to-end AI-powered recruitment pipeline that automates the entire candidate screening process — from sourcing to interview booking — using Claude AI.

Instead of HR spending days manually reviewing profiles, this system does it in minutes.

## The problem it solves
Traditional hiring is slow. HR teams manually browse LinkedIn, read every profile, send individual messages, review responses, and book interviews. This takes days.

This pipeline reduces that to a few hours — mostly automated.

## How it works
1. **Source candidates** — pulls profiles from LinkedIn, Indeed, Naukri or any other platform.
2. **AI sorts by JD** — Claude reads each profile and matches it against the job description upto 70%.
3. **Auto-message sent** — shortlisted candidates get an invite to fill a screening form
4. **Candidate fills form** — structured questions relevant to the role
5. **App script reviews responses** — App script evaluates answers and decides: Shortlist / Review / Reject
6. **Nightly HR email** — HR gets a summary of all results every evening
7. **Calendar auto-booked** — shortlisted candidates get an interview slot
8. **Interview reminder** — candidate and the HR gets an automatic reminder on the day

## Tools used
- Claude Desktop (AI decision making + prompt engineering)
- Google Forms (candidate screening form)
- Gmail (automated emails)
- Google Apps Script (form evaluation + shortlist logic)
- Google Calendar (interview booking)

## Results
- Candidate screening time reduced from days to hours
- HR only needs to review the nightly summary email
- Zero manual messaging or calendar booking

## Status
Active — currently being used for real recruitment workflows.
Being expanded to include Indeed, Naukri, and more platforms.
