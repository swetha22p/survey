# Learning Format Study

A Flask study website comparing text-based and concept-map learning materials. The Flask application serves the participant interface and uses MongoDB Atlas for persistent participant records and stratified allocation counts.

## Architecture

- **Application and frontend:** Flask templates and static assets, hosted together on Render.
- **Backend:** Flask, hosted as a Render web service.
- **Database:** MongoDB Atlas M0 Free cluster.
- **Researcher dashboard/export:** protected by `RESEARCHER_KEY`.

## Run locally

Use Python 3.10+.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` with your Atlas connection string and unique secrets. Load environment variables before starting Flask. One simple approach is:

```bash
export MONGODB_URI='your-atlas-connection-string'
export MONGODB_DB='learning_format_study'
export SECRET_KEY='your-long-random-secret'
export RESEARCHER_KEY='your-private-researcher-key'
python app.py
```

Open `http://127.0.0.1:5005`.

## MongoDB Atlas M0 setup

1. Create an Atlas project and an **M0 Free** cluster.
2. Create a database user with a strong password.
3. In **Network Access**, allow the Render service to connect. For initial testing, Atlas's `0.0.0.0/0` access rule can be used, but it allows connections from any IP and should be paired with a strong database password. Prefer a narrower rule where practical.
4. Select **Connect → Drivers** and copy the Python connection string.
5. Replace the username, password, and cluster host in `MONGODB_URI`. URL-encode special characters in the password.

The application creates the `participants` and `allocations` collections and their indexes when a participant begins. Allocation counts are stored in MongoDB, not on Render's ephemeral disk.

## Deploy to Render (free web service)

1. Push this project folder to a GitHub repository. Do not commit `.env`, secrets, or your virtual environment.
2. In Render, choose **New → Blueprint** and connect the repository, or create a Web Service manually.
3. Use these settings if configuring manually:
   - Runtime: Python
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
4. Add environment variables:
   - `MONGODB_URI`: Atlas connection string
   - `MONGODB_DB`: `learning_format_study`
   - `SECRET_KEY`: long random secret
   - `RESEARCHER_KEY`: separate private key for researcher-only pages
5. Deploy, open the Render URL, and test the participant journey end-to-end.

Render's free service may sleep after inactivity, so the first visit after a period of inactivity can take some time. Free service availability and limits may change; review Render's current plan details.

## Researcher dashboard and export

Use the secret researcher key configured in Render. For example:

- Dashboard: `https://YOUR-RENDER-SERVICE.onrender.com/dash?key=YOUR_RESEARCHER_KEY`
- Excel export: `https://YOUR-RENDER-SERVICE.onrender.com/researcher/export?key=YOUR_RESEARCHER_KEY`

Keep the key private. The app returns 403 if the key is missing or incorrect.

## Notes

- Existing records in `study_responses.xlsx` are **not automatically imported** into MongoDB. Import any existing participant data separately before going live if you need it retained in the live database.
- Do not use the researcher export or dashboard key in participant-facing links.
- Review the consent text, study content, and data-retention requirements before recruiting participants.
