from fastapi import FastAPI, APIRouter, File, UploadFile, HTTPException, Query
from fastapi.responses import JSONResponse
from dotenv import load_dotenv
from fastapi.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional
import uuid
from datetime import datetime, timezone
import openpyxl
import pandas as pd
from io import BytesIO
from contextlib import asynccontextmanager

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# Définir le gestionnaire de durée de vie
@asynccontextmanager
async def lifespan(app: FastAPI):
    # Au démarrage
    logging.info("Application démarrage...")
    
    # Initialiser la connexion MongoDB
    mongo_url = os.environ['MONGO_URL']
    client = AsyncIOMotorClient(mongo_url)
    db = client[os.environ['DB_NAME']]
    
    # Stocker dans l'état de l'application pour y accéder ailleurs
    app.state.mongo_client = client
    app.state.db = db
    
    yield  # L'application est en cours d'exécution
    
    # À l'arrêt
    logging.info("Application arrêt...")
    if hasattr(app.state, 'mongo_client'):
        app.state.mongo_client.close()
        logging.info("Connexion MongoDB fermée")

# Créer l'application avec le lifespan handler
app = FastAPI(lifespan=lifespan)

# CORS MIDDLEWARE - UN SEUL FOIS ICI
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "https://planning-cours-leeb.vercel.app",
        "https://planning-cours-leeb-42qu09vff-arnauds-projects-0c50c94d.vercel.app",
        "http://localhost:3000",
        "http://localhost:8080", 
        "http://localhost:8000"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create a router with the /api prefix
api_router = APIRouter(prefix="/api")

# Define Models
class Cours(BaseModel):
    model_config = ConfigDict(extra="ignore")
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    classe: str
    matiere: str
    date: str  # Format: DD/MM/YYYY ou YYYY-MM-DD
    start_time: str  # Format: HH:MM
    end_time: str  # Format: HH:MM
    intervenant: str
    heures_eq_td: float
    type_cours: str = "Cours"
    salle: str = ""
    commentaire: str = ""
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class CoursCreate(BaseModel):
    classe: str
    matiere: str
    date: str
    start_time: str
    end_time: str
    intervenant: str
    heures_eq_td: float
    type_cours: str = "Cours"
    salle: str = ""
    commentaire: str = ""

class ImportMetadata(BaseModel):
    model_config = ConfigDict(extra="ignore")
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    filename: str
    import_date: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    courses_count: int

class StatsResponse(BaseModel):
    total_cours: int
    total_heures: float
    total_intervenants: int

class SynthesisData(BaseModel):
    par_intervenant: List[dict]
    par_matiere_intervenant: List[dict]

@api_router.get("/")
async def root():
    return {"message": "API Planning des Cours"}

@api_router.post("/import-excel")
async def import_excel(file: UploadFile = File(...)):
    """
    Import Excel file with courses data.
    Expected columns: Classe, Matière, Date, Start Time, End Time, Intervenant, Heures équivalent TD, Salle
    """
    try:
        contents = await file.read()
        filename = file.filename
        
        # Load workbook with data_only=True to get calculated formula values
        wb = openpyxl.load_workbook(BytesIO(contents), data_only=True)
        sheet = wb.active
        
        # Get headers from first row
        headers = []
        for cell in sheet[1]:
            if cell.value:
                headers.append(str(cell.value).strip())
        
        # Parse data
        courses_data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not any(row):  # Skip empty rows
                continue
                
            row_dict = {}
            for idx, value in enumerate(row):
                if idx < len(headers) and value is not None:
                    row_dict[headers[idx]] = value
            
            if not row_dict:
                continue
            
            # Map columns (flexible matching)
            cours_data = {}
            
            # Find classe
            for key in ['Classe', 'classe', 'Class', 'CLASSE']:
                if key in row_dict:
                    cours_data['classe'] = str(row_dict[key]).strip()
                    break
            
            # Find matiere
            for key in ['Matière', 'matiere', 'Matiere', 'Subject', 'MATIERE']:
                if key in row_dict:
                    cours_data['matiere'] = str(row_dict[key]).strip()
                    break
            
            # Find date
            for key in ['Date', 'date', 'DATE', 'Start Date', 'start_date']:
                if key in row_dict:
                    date_val = row_dict[key]
                    if isinstance(date_val, datetime):
                        cours_data['date'] = date_val.strftime('%d/%m/%Y')
                    else:
                        cours_data['date'] = str(date_val).strip()
                    break
            
            # Find start_time - ALWAYS use Start Time column
            for key in ['Start Time', 'start_time', 'START TIME']:
                if key in row_dict:
                    time_val = row_dict[key]
                    if isinstance(time_val, datetime):
                        cours_data['start_time'] = time_val.strftime('%H:%M')
                    elif time_val:
                        cours_data['start_time'] = str(time_val).strip()
                    break
            
            # Find end_time
            for key in ['End Time', 'end_time', 'END TIME']:
                if key in row_dict:
                    time_val = row_dict[key]
                    if isinstance(time_val, datetime):
                        cours_data['end_time'] = time_val.strftime('%H:%M')
                    elif time_val:
                        cours_data['end_time'] = str(time_val).strip()
                    break
            
            # Find intervenant
            for key in ['Intervenant', 'intervenant', 'Teacher', 'Professeur', 'INTERVENANT']:
                if key in row_dict:
                    cours_data['intervenant'] = str(row_dict[key]).strip()
                    break
            
            # Find heures_eq_td - ALWAYS prioritize "Nombre d'heure" (column L)
            for key in ["Nombre d'heure", "Nombre d heure", 'NOMBRE D HEURE', 'Nb heure_eq TD', 'Heures équivalent TD', 'heures_eq_td', 'Heures', 'Hours', 'HEURES']:
                if key in row_dict:
                    try:
                        val = row_dict[key]
                        # Skip if value looks like a formula
                        if isinstance(val, str) and val.startswith('='):
                            continue
                        cours_data['heures_eq_td'] = float(val)
                        break
                    except:
                        continue
            
            # Default if not found
            if 'heures_eq_td' not in cours_data:
                cours_data['heures_eq_td'] = 1.0
            
            # Find salle
            for key in ['Salle', 'salle', 'Location', 'Room', 'SALLE']:
                if key in row_dict:
                    salle_val = row_dict[key]
                    if salle_val:
                        cours_data['salle'] = str(salle_val).strip()
                    break
            
            # Optional: type_cours
            for key in ['Type', 'type', 'Type cours', 'TYPE', 'CM/TD/TP']:
                if key in row_dict:
                    type_val = row_dict[key]
                    if type_val:
                        cours_data['type_cours'] = str(type_val).strip()
                    break
            
            # Optional: commentaire
            for key in ['Commentaire', 'commentaire', 'Comment', 'COMMENTAIRE']:
                if key in row_dict:
                    comment_val = row_dict[key]
                    if comment_val:
                        cours_data['commentaire'] = str(comment_val).strip()
                    break
            
            # Validate required fields
            required_fields = ['classe', 'matiere', 'date', 'start_time', 'end_time', 'intervenant']
            if all(field in cours_data for field in required_fields):
                # Set defaults for optional fields
                if 'heures_eq_td' not in cours_data:
                    cours_data['heures_eq_td'] = 1.0
                if 'type_cours' not in cours_data:
                    cours_data['type_cours'] = 'Cours'
                if 'salle' not in cours_data:
                    cours_data['salle'] = ''
                if 'commentaire' not in cours_data:
                    cours_data['commentaire'] = ''
                
                courses_data.append(cours_data)
        
        if not courses_data:
            raise HTTPException(status_code=400, detail="Aucune donnée valide trouvée dans le fichier Excel")
        
        # Clear existing courses and metadata
        await app.state.db.courses.delete_many({})
        await app.state.db.import_metadata.delete_many({})
        
        # Insert courses
        cours_objects = []
        for data in courses_data:
            cours = Cours(**data)
            doc = cours.model_dump()
            doc['created_at'] = doc['created_at'].isoformat()
            cours_objects.append(doc)
        
        if cours_objects:
            await app.state.db.courses.insert_many(cours_objects)
        
        # Save import metadata
        metadata = ImportMetadata(
            filename=filename,
            courses_count=len(cours_objects)
        )
        metadata_doc = metadata.model_dump()
        metadata_doc['import_date'] = metadata_doc['import_date'].isoformat()
        await app.state.db.import_metadata.insert_one(metadata_doc)
        
        return JSONResponse(content={
            "message": f"{len(cours_objects)} cours importés avec succès",
            "count": len(cours_objects),
            "filename": filename
        })
    
    except Exception as e:
        logging.error(f"Error importing Excel: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Erreur lors de l'import: {str(e)}")

@api_router.get("/import-metadata")
async def get_import_metadata():
    """
    Get last import metadata (filename and date).
    """
    metadata = await app.state.db.import_metadata.find_one({}, {"_id": 0}, sort=[("import_date", -1)])
    if metadata:
        if isinstance(metadata.get('import_date'), str):
            metadata['import_date'] = datetime.fromisoformat(metadata['import_date'])
        return metadata
    return None

@api_router.get("/courses", response_model=List[Cours])
async def get_courses(
    classe: Optional[str] = Query(None),
    intervenant: Optional[str] = Query(None),
    matiere: Optional[str] = Query(None)
):
    """
    Get all courses with optional filters.
    """
    query = {}
    
    if classe:
        query['classe'] = classe
    if intervenant:
        query['intervenant'] = intervenant
    if matiere:
        query['matiere'] = matiere
    
    courses = await app.state.db.courses.find(query, {"_id": 0}).to_list(10000)
    
    # Convert ISO string timestamps back to datetime objects
    for course in courses:
        if isinstance(course['created_at'], str):
            course['created_at'] = datetime.fromisoformat(course['created_at'])
    
    return courses

@api_router.get("/stats", response_model=StatsResponse)
async def get_stats():
    """
    Get statistics: total courses, total hours, and intervenants.
    """
    courses = await app.state.db.courses.find({}, {"_id": 0, "heures_eq_td": 1}).to_list(10000)
    
    total_cours = len(courses)
    total_heures = sum(course.get('heures_eq_td', 0) for course in courses)
    
    intervenants = await app.state.db.courses.distinct('intervenant')
    
    return StatsResponse(
        total_cours=total_cours,
        total_heures=round(total_heures, 2),
        total_intervenants=len(intervenants)
    )

@api_router.get("/classes")
async def get_classes():
    """
    Get list of unique classes.
    """
    classes = await app.state.db.courses.distinct('classe')
    return sorted(classes)

@api_router.get("/intervenants")
async def get_intervenants():
    """
    Get list of unique intervenants.
    """
    intervenants = await app.state.db.courses.distinct('intervenant')
    return sorted(intervenants)

@api_router.get("/matieres")
async def get_matieres():
    """
    Get list of unique matieres.
    """
    matieres = await app.state.db.courses.distinct('matiere')
    return sorted(matieres)

@api_router.get("/synthesis")
async def get_synthesis(
    classe: Optional[str] = Query(None),
    intervenant: Optional[str] = Query(None)
):
    """
    Get synthesis data for charts:
    - Hours per intervenant
    - Hours per matiere for each intervenant
    """
    query = {}
    if classe:
        query['classe'] = classe
    if intervenant:
        query['intervenant'] = intervenant
    
    courses = await app.state.db.courses.find(query, {"_id": 0}).to_list(10000)
    
    # Calculate hours and courses count per intervenant
    intervenant_hours = {}
    intervenant_count = {}
    matiere_intervenant_hours = {}
    
    for course in courses:
        inter = course['intervenant']
        mat = course['matiere']
        hours = course.get('heures_eq_td', 1.0)
        
        # Sum hours by intervenant
        if inter not in intervenant_hours:
            intervenant_hours[inter] = 0
            intervenant_count[inter] = 0
        intervenant_hours[inter] += hours
        intervenant_count[inter] += 1
        
        # Sum by matiere and intervenant
        key = f"{inter}|{mat}"
        if key not in matiere_intervenant_hours:
            matiere_intervenant_hours[key] = {'intervenant': inter, 'matiere': mat, 'heures': 0}
        matiere_intervenant_hours[key]['heures'] += hours
    
    # Format for frontend
    par_intervenant = [{'name': k, 'value': round(v, 2)} for k, v in intervenant_hours.items()]
    par_intervenant_count = [{'name': k, 'value': v} for k, v in intervenant_count.items()]
    par_matiere_intervenant = list(matiere_intervenant_hours.values())
    
    # Round heures
    for item in par_matiere_intervenant:
        item['heures'] = round(item['heures'], 2)
    
    return {
        'par_intervenant': par_intervenant,
        'par_intervenant_count': par_intervenant_count,
        'par_matiere_intervenant': par_matiere_intervenant
    }

# Include the router in the main app
app.include_router(api_router)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)