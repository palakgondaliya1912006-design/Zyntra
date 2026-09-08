from dotenv import load_dotenv
import os
from pymongo import MongoClient

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

client = MongoClient(os.getenv("MONGODB_URI"))

db = client["Zyntra"]

users = db["users"]
notes = db["notes"]
tasks = db["tasks"]
files = db["files"]
history = db["analysis_history"]
events = db["events"]
