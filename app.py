
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file, abort
import random, os, json, io
from datetime import datetime, UTC
from threading import Lock
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-development-key-change-me")
MONGO_URI = os.environ.get("MONGODB_URI", "")
MONGO_DB_NAME = os.environ.get("MONGODB_DB", "learning_format_study")
_mongo_client = None
_mongo_lock = Lock()
allocation_lock = Lock()


def get_db():
    """Create the MongoDB client lazily so the app can import before env setup."""
    global _mongo_client
    if not MONGO_URI:
        raise RuntimeError("MONGODB_URI environment variable is not configured.")
    if _mongo_client is None:
        with _mongo_lock:
            if _mongo_client is None:
                _mongo_client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=10000)
                _mongo_client.admin.command("ping")
    return _mongo_client[MONGO_DB_NAME]


def init_db():
    db = get_db()
    db.participants.create_index("participant_code", unique=True)
    db.allocations.create_index("prior_level", unique=True)
    for level in ("no_basic", "average", "excellent"):
        db.allocations.update_one(
            {"prior_level": level},
            {"$setOnInsert": {"prior_level": level, "text_count": 0, "concept_map_count": 0}},
            upsert=True,
        )



def allocate_condition(level):
    """Balance assignment within a prior-knowledge stratum for this study."""
    with allocation_lock:
        db = get_db()

        doc = db.allocations.find_one({"prior_level": level}) or {
            "prior_level": level,
            "text_count": 0,
            "concept_map_count": 0
        }

        text_count = int(doc.get("text_count", 0))
        map_count = int(doc.get("concept_map_count", 0))

        if text_count < map_count:
            condition = "text"
        elif map_count < text_count:
            condition = "concept_map"
        else:
            condition = random.choice(["text", "concept_map"])

        field = "text_count" if condition == "text" else "concept_map_count"

        # Increment only the selected counter.
        # init_db() already creates the allocation documents.
        db.allocations.update_one(
            {"prior_level": level},
            {"$inc": {field: 1}},
            upsert=True
        )

        return condition


# Replace placeholder quiz items and learning materials with reviewed study content.
PRIOR_KNOWLEDGE_ITEMS = [
    {
        "id": "pk1",
        "q": "How familiar are you with computer networks?",
        "type": "self",
        "options": [
            "I have no idea",
            "I have a basic idea",
            "I know it fairly well",
            "I know it very well"
        ]
    },
    {
        "id": "pk2",
        "q": "A college wants to connect all the computers of the computer lab with a printer. Which type of network would be most suitable?",
        "type": "mc",
        "options": [
            "PAN",
            "LAN",
            "MAN",
            "WAN"
        ],
        "answer": 1
    },
    {
        "id": "pk3",
        "q": "A company has offices in different countries and wants to connect their local networks. Which type of network is most appropriate?",
        "type": "mc",
        "options": [
            "PAN",
            "LAN",
            "MAN",
            "WAN"
        ],
        "answer": 3
    },
    {
        "id": "pk4",
        "q": "In a network, data should be sent only to the device whose address matches the destination. Which device is designed to forward data to selected devices within a LAN?",
        "type": "mc",
        "options": [
            "Hub",
            "Switch",
            "Repeater",
            "Modem"
        ],
        "answer": 1
    },
    {
        "id": "pk5",
        "q": "A network should continue transmitting data between other devices even if one device fails. Which topology does supports this requirement?",
        "type": "mc",
        "options": [
            "Bus",
            "Ring",
            "Mesh",
            "Star"
        ],
        "answer": 2
    },
    {
        "id": "pk6",
        "q": "A computer is moved from one network to another. Which address may change as a result of this movement?",
        "type": "mc",
        "options": [
            "MAC address",
            "IP address",
            "Manufacturer ID",
            "NIC hardware address"
        ],
        "answer": 1
    },
    {
        "id": "pk7",
        "q": "A user enters a website's domain name in a browser. Which service helps the browser find the corresponding IP address?",
        "type": "mc",
        "options": [
            "DNS",
            "NIC",
            "Ethernet",
            "HTTP"
        ],
        "answer": 0
    },
    {
        "id": "pk8",
        "q": "A user wants to access linked web pages, while the Internet provides the underlying global network of connected devices. Which statement best distinguishes the Web from the Internet?",
        "type": "mc",
        "options": [
            "The Web is the global network, and the Internet is a collection of web pages",
            "The Internet is the global network, and the Web is a collection of linked resources accessed through it",
            "The Web and the Internet are exactly the same",
            "The Internet is used only for email, while the Web is used only for browsing"
        ],
        "answer": 1
    }
]

QUIZ1_ITEMS = [
    {
        "id": "q1_1",
        "q": "A computer needs to send a message to one particular computer on a network. Which communication approach best fits this situation?",
        "options": [
            "Sending the data to a single intended recipient",
            "Sending the data to every device on the network",
            "Sending the data without identifying a destination",
            "Sending the data only to an intermediate router"
        ],
        "answer": 0
    },
    {
        "id": "q1_2",
        "q": "A packet is created to travel from one device to another. Which information helps identify the devices involved in this communication?",
        "options": [
            "The source and destination IP addresses",
            "The number of links in the network",
            "The names of unrelated devices",
            "The size of the network map"
        ],
        "answer": 0
    },
    {
        "id": "q1_3",
        "q": "A device sends a packet to another device located on a different network. What must happen for the packet to move toward its destination?",
        "options": [
            "A routing decision must determine where the packet should be forwarded",
            "The source device must change into the destination device",
            "The packet must lose its destination IP address",
            "The destination device must create a new source address"
        ],
        "answer": 0
    },
    {
        "id": "q1_4",
        "q": "A sender transmits a packet, but the receiving device does not get it because of network congestion. Which outcome is most consistent with this situation?",
        "options": [
            "The packet may be lost or fail to reach its destination",
            "The packet is guaranteed to reach the destination",
            "The destination IP address automatically changes",
            "The communication becomes a broadcast"
        ],
        "answer": 0
    },
    {
        "id": "q1_5",
        "q": "A receiving device gets a packet and needs to determine whether it came from the expected sender. Which packet information is most relevant?",
        "options": [
            "The source IP address",
            "The number of routers in the network",
            "The name of the network diagram",
            "The number of unrelated connected devices"
        ],
        "answer": 0
    }
]
QUIZ2_ITEMS = [
 {"id":"q2_1","q":"A computer sends information through an analog network. Why is a modem suitable here, rather than relying on a router’s packet-routing role?","options":["It converts signals through modulation and demodulation","It chooses a packet route between nodes","It provides only a power supply","It performs e-waste recycling during transmission"],"answer":0},
 {"id":"q2_2","q":"A modem is connected correctly, but communication quality varies as the analog signal changes. Which explanation best fits the supplied material?","options":["Its operation may depend on analog signal quality","Its related entities determine signal quality","Its vendor birth process routes packets","The network cable performs demodulation alone"],"answer":0},
 {"id":"q2_3","q":"Another nearby device interferes with a modem. Which conclusion reflects the material without assuming every modem will fail?","options":["Interference is a possible side effect that can affect operation","Interference is the modem’s primary purpose","The modem must immediately be disposed of","The modem can no longer contain a modulator"],"answer":0},
 {"id":"q2_4","q":"A service plan aims to reduce avoidable modem issues over time. Which pair of actions best matches the maintenance information?","options":["Software updates and hardware checks","Modulation and packet routing","Vendor configuration and e-waste disposal","Changing source and destination IP addresses"],"answer":0},
 {"id":"q2_5","q":"A home has both a modem and a router. Which explanation best distinguishes their roles based on the supplied materials?","options":["The modem modulates/demodulates signals; the router routes packets between networks","The modem routes packets while the router demodulates signals","Both devices are power supplies","The modem’s main function is e-waste recycling"],"answer":0}
]

PARTICIPANT_HEADERS = ["participant_code", "consent", "familiarity", "prior_score", "prior_level", "prior_answers", "condition_name", "unicast_score", "modem_score", "post_score", "post_answers", "feedback_clarity", "feedback_difficulty", "feedback_comment", "started_at", "completed_at"]
ALLOCATION_HEADERS = ["prior_level", "text_count", "concept_map_count"]

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/consent", methods=["POST"])
def consent():
    if request.form.get("consent") != "yes":
        return render_template("message.html", title="Consent required", message="You chose not to participate. No study response has been saved.")
    try:
        init_db()
    except PyMongoError:
        app.logger.exception("MongoDB initialization failed")
        return render_template("message.html", title="Database unavailable", message="The study database is temporarily unavailable. Please try again later."), 503
    session.clear()
    session["participant_code"] = "P" + datetime.utcnow().strftime("%y%m%d%H%M%S") + str(random.randint(100,999))
    session["started_at"] = datetime.now(UTC).isoformat()
    return redirect(url_for("prior"))

@app.route("/prior", methods=["GET", "POST"])
def prior():
    if "participant_code" not in session:
        return redirect(url_for("index"))
    if request.method == "GET":
        return render_template("questionnaire.html", items=PRIOR_KNOWLEDGE_ITEMS, title="Step 1")
                            #   subtitle="First, tell us what you already know. You will not see the learning material until you submit this page.")
    answers = {}
    correct = 0
    for item in PRIOR_KNOWLEDGE_ITEMS:
        val = request.form.get(item["id"], "")
        answers[item["id"]] = val
        if item["type"] == "mc" and val.isdigit() and int(val) == item["answer"]:
            correct += 1
    # Objective score uses knowledge questions only; self-report familiarity is saved separately.
    n_obj = sum(1 for i in PRIOR_KNOWLEDGE_ITEMS if i["type"] == "mc")
    score = round(correct / n_obj * 100, 2)
    familiarity = answers.get("pk1", "")
    # Three categories for allocation; thresholds are pilot defaults and should be pre-registered/validated.
    if score <= 33.33:
        level = "no_basic"
    elif score <= 66.67:
        level = "average"
    else:
        level = "excellent"
    session["prior_answers"] = answers
    session["prior_score"] = score
    session["prior_level"] = level
    session["familiarity"] = familiarity
    # Balanced randomization within each prior-knowledge stratum, tracked in Excel.
    condition = allocate_condition(level)
    session["condition"] = condition
    session["stage"] = 1
    return redirect(url_for("material"))

@app.route("/material", methods=["GET"])
def material():
    if "participant_code" not in session:
        return redirect(url_for("index"))
    stage = int(session.get("stage", 1))
    condition = session.get("condition", "concept_map")
    return render_template("material.html", condition=condition, stage=stage)

@app.route("/assessment", methods=["GET", "POST"])
def assessment():
    if "participant_code" not in session:
        return redirect(url_for("index"))
    stage = int(session.get("stage", 1))
    items = QUIZ1_ITEMS if stage == 1 else QUIZ2_ITEMS
    if request.method == "GET":
        return render_template("questionnaire.html", items=items,
            title=("Unicast comprehension test" if stage == 1 else "Modem comprehension test"),
           )
    answers, correct = {}, 0
    for item in items:
        val = request.form.get(item["id"], "")
        answers[item["id"]] = val
        if val.isdigit() and int(val) == item["answer"]:
            correct += 1
    score = round(correct / len(items) * 100, 2)
    session[f"quiz{stage}_answers"] = answers
    session[f"quiz{stage}_score"] = score
    if stage == 1:
        session["stage"] = 2
        return redirect(url_for("material"))
    session["post_score"] = round((session.get("quiz1_score",0) + score) / 2, 2)
    session["post_answers"] = {**session.get("quiz1_answers",{}), **answers}
    return redirect(url_for("feedback"))

@app.route("/feedback", methods=["GET", "POST"])
def feedback():
    if "participant_code" not in session:
        return redirect(url_for("index"))
    if request.method == "GET":
        return render_template("feedback.html")
    code = session["participant_code"]
    record = {
        "participant_code": code, "consent": 1, "familiarity": session.get("familiarity", ""),
        "prior_score": session.get("prior_score", 0), "prior_level": session.get("prior_level", ""),
        "prior_answers": json.dumps(session.get("prior_answers", {}), ensure_ascii=False),
        "condition_name": session.get("condition", ""), "unicast_score": session.get("quiz1_score", 0),
        "modem_score": session.get("quiz2_score", 0), "post_score": session.get("post_score", 0),
        "post_answers": json.dumps(session.get("post_answers", {}), ensure_ascii=False),
        "feedback_clarity": request.form.get("clarity"), "feedback_difficulty": request.form.get("difficulty"),
        "feedback_comment": request.form.get("comment", ""), "started_at": session.get("started_at", ""),
        "completed_at": datetime.now(UTC).isoformat()
    }
    try:
        get_db().participants.insert_one(record)
    except PyMongoError:
        app.logger.exception("Unable to save participant response to MongoDB")
        return render_template("message.html", title="Unable to save", message="We could not save your response right now. Please contact the researcher."), 503
    return render_template("message.html", title="Thank you!", message="Your response has been recorded. You may now close this page.")


def require_researcher_key():
    """Protect researcher-only pages using a key configured in Render."""
    expected = os.environ.get("RESEARCHER_KEY", "")
    supplied = request.args.get("key", "") or request.headers.get("X-Researcher-Key", "")
    if not expected or supplied != expected:
        abort(403)


@app.route("/researcher/export")
def export():
    require_researcher_key()
    records = list(get_db().participants.find({}, {"_id": 0}))
    wb = Workbook()
    ws = wb.active
    ws.title = "Participants"
    ws.append(PARTICIPANT_HEADERS)
    for record in records:
        ws.append([record.get(h, "") for h in PARTICIPANT_HEADERS])
    alloc = wb.create_sheet("Allocations")
    alloc.append(ALLOCATION_HEADERS)
    for row in get_db().allocations.find({}, {"_id": 0}).sort("prior_level", 1):
        alloc.append([row.get(h, 0 if h != "prior_level" else "") for h in ALLOCATION_HEADERS])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="study_responses.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/dash")
def dashboard():
    require_researcher_key()
    records = list(get_db().participants.find({}, {"_id": 0}).sort("completed_at", 1))
    total = len(records)
    complete = sum(1 for r in records if r.get("completed_at"))
    groups = {"text": 0, "concept_map": 0}
    levels = {"no_basic": {"text": 0, "concept_map": 0}, "average": {"text": 0, "concept_map": 0}, "excellent": {"text": 0, "concept_map": 0}}
    scores = {"text": [], "concept_map": []}
    for r in records:
        group = r.get("condition_name") or ""
        level = r.get("prior_level") or ""
        if group in groups:
            groups[group] += 1
            if level in levels: levels[level][group] += 1
            try: scores[group].append(float(r.get("post_score") or 0))
            except (ValueError, TypeError): pass
    averages = {g: round(sum(v)/len(v), 1) if v else 0 for g,v in scores.items()}
    recent = list(reversed(records[-15:]))
    return render_template("dashboard.html", total=total, complete=complete, groups=groups, levels=levels, averages=averages, recent=recent)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5005)), debug=True)
