from collections import defaultdict
import calendar
import time
import datetime
from datetime import datetime, timedelta, date, timezone
from typing import Optional, List
import io
import csv
import json
import urllib.parse

from fastapi import (
    FastAPI, 
    Depends, 
    HTTPException, 
    Request, 
    Form, 
    status
)
from fastapi.responses import (
    RedirectResponse, 
    HTMLResponse, 
    Response, 
    StreamingResponse, 
    JSONResponse
)
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.datastructures import URL

from sqlalchemy.orm import Session, joinedload, relationship
from sqlalchemy import or_, inspect, Column, Integer, String, Float, DateTime, ForeignKey, Boolean, func, text

from pydantic import BaseModel

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

# استيراد قاعدة البيانات (engine و SessionLocal) من ملف database.py
from database import engine, SessionLocal, Base

# استيراد النماذج والجداول من ملف models.py
from models import (
    Company,
    Client,
    Segment,
    Location,  
    Vehicle,
    Trailer,
    Driver,
    DriverLocationLog,  
    DriverLeave,  
    Route,
    RouteViaPoint,  
    LoadType,
    User,
    Setting,
    FleetAssignment,  
    Journey,
    JourneyFollowUp,
    JourneyMaintenance,  
    OffDutyAssignment,
    JourneyStop,  
)
from jinja2 import Environment, FileSystemLoader
from fastapi.middleware.cors import CORSMiddleware
app = FastAPI(title="Journey Management System", redirect_slashes=False)
templates = Jinja2Templates(directory="templates")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # السماح لكل النطاقات والمصادر
    allow_credentials=True,
    allow_methods=["*"],  # السماح بكل الطرق (GET, POST, etc.)
    allow_headers=["*"],  # السماح بكل الهيدرز
)
jinja_env = Environment(loader=FileSystemLoader("templates"))

Base.metadata.create_all(bind=engine)


# ملفات الـ Static
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# دالة مساعدة لجلب المستخدم الحالي النشط من الـ Cookie بناءً على البريد الإلكتروني (email)
def get_current_user(request: Request, db: Session = Depends(get_db)) -> Optional[User]:
    user_email = request.cookies.get("user_email")
    if not user_email:
        return None
    return db.query(User).filter(User.email == user_email).first()

# ----------------------------------------------------
# AUTHENTICATION DEPENDENCY
# ----------------------------------------------------
def verify_login(request: Request):
    username = request.cookies.get("username")
    if not username:
        raise HTTPException(
            status_code=status.HTTP_303_SEE_OTHER,
            headers={"Location": "/login"}
        )
    return username

def require_admin_or_developer(request: Request, db: Session = Depends(get_db)):
    # تم تعديل الاعتماد على الـ username لتجنب مشاكل الـ email الفارغ
    username = request.cookies.get("username")
    if not username:
        raise HTTPException(status_code=status.HTTP_303_SEE_OTHER, headers={"Location": "/login"})
    
    user = db.query(User).filter(User.username == username).first()
    if not user:
        raise HTTPException(status_code=status.HTTP_303_SEE_OTHER, headers={"Location": "/login"})
    
    user_role = getattr(user, 'role', '').lower()
    if user_role not in ['admin', 'developer', 'مشرف']:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="عذراً، لا تمتلك الصلاحيات الكافية للقيام بهذا الإجراء."
        )
    return user

def require_developer(user: User):
    return True

# 2. تطبيق الحماية كـ middleware مع استثناء الـ static بالكامل
@app.middleware("http")
async def check_user_auth(request: Request, call_next):
    path = request.url.path
    
    # السماح لمسارات تسجيل الدخول، الـ static، ومسارات الـ API الخاصة بتتبع السائق بالمرور فوراً
    if path in ["/login", "/logout"] or path.startswith("/static/") or path.startswith("/api/driver/"):
        return await call_next(request)
    
    # الاعتماد الأساسي على الـ username لباقي صفحات الموقع
    username = request.cookies.get("username")
    if not username:
        return RedirectResponse(url="/login", status_code=303)

    response = await call_next(request)
    return response


# ----------------------------------------------------
# AUTHENTICATION (Login / Logout)
# ----------------------------------------------------
@app.post("/login", response_class=HTMLResponse)
def login_action(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db)
):
    user = db.query(User).filter(User.username == username, User.password == password).first()
    
    if not user:
        return templates.TemplateResponse(
            request, 
            "login.html", 
            {"error": "اسم المستخدم أو كلمة المرور غير صحيحة"}
        )
    
    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(key="username", value=user.username)
    # حماية ضد القيم الفارغة للـ email لتجنب أخطاء الكوكيز
    response.set_cookie(key="user_email", value=user.email if user.email else "")
    response.set_cookie(key="user_role", value=user.role if hasattr(user, 'role') and user.role else "User")
    return response

@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(key="username")
    response.delete_cookie(key="user_email")
    response.delete_cookie(key="user_role")
    return response

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse(request, "login.html", {"request": request})


# ----------------------------------------------------
# DASHBOARD
# ----------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def dashboard(
    request: Request, 
    db: Session = Depends(get_db),
    current_user = Depends(get_current_user)
):
    try:
        # 1. التحقق هل هو مشرف عام حقيقي (Superuser أو Full Access)
        is_superuser = getattr(current_user, 'is_superuser', False)
        
        allowed_company_ids = []
        allowed_client_ids = []
        
        if current_user:
            # جمع شركات النقل المتاحة للمستخدم
            for rel_name in ['companies', 'transport_companies']:
                if hasattr(current_user, rel_name):
                    rel_data = getattr(current_user, rel_name)
                    if rel_data:
                        for c in rel_data:
                            comp_id = getattr(c, 'id', None) or (c[0] if isinstance(c, (list, tuple)) else None)
                            if comp_id and comp_id not in allowed_company_ids:
                                allowed_company_ids.append(comp_id)
            
            # جمع العُملاء المتاحين من علاقات المستخدم المباشرة
            for rel_name in ['clients', 'client_companies']:
                if hasattr(current_user, rel_name):
                    rel_data = getattr(current_user, rel_name)
                    if rel_data:
                        for cl in rel_data:
                            client_id = getattr(cl, 'id', None) or (cl[0] if isinstance(cl, (list, tuple)) else None)
                            if client_id and client_id not in allowed_client_ids:
                                allowed_client_ids.append(client_id)
            
            if hasattr(current_user, 'company_id') and current_user.company_id:
                if current_user.company_id not in allowed_company_ids:
                    allowed_company_ids.append(current_user.company_id)
                    
            if hasattr(current_user, 'client_company_id') and current_user.client_company_id:
                if current_user.client_company_id not in allowed_client_ids:
                    allowed_client_ids.append(current_user.client_company_id)

            # فحص الجداول الوسيطة للشركات
            for table_name in ["user_companies", "user_company", "user_transport_companies"]:
                try:
                    result = db.execute(
                        text(f"SELECT company_id FROM {table_name} WHERE user_id = :uid"), 
                        {"uid": current_user.id}
                    ).fetchall()
                    for row in result:
                        if row[0] not in allowed_company_ids:
                            allowed_company_ids.append(row[0])
                except Exception:
                    db.rollback()

            # فحص الجدول الوسيط للعُملاء user_clients
            for table_name in ["user_clients", "user_client"]:
                try:
                    result = db.execute(
                        text(f"SELECT client_id FROM {table_name} WHERE user_id = :uid"), 
                        {"uid": current_user.id}
                    ).fetchall()
                    for row in result:
                        if row[0] not in allowed_client_ids:
                            allowed_client_ids.append(row[0])
                except Exception:
                    db.rollback()

        print(f"--- DEBUG DASHBOARD ---")
        print(f"User Email: {getattr(current_user, 'email', 'N/A')}")
        print(f"Is Superuser?: {is_superuser}")
        print(f"Allowed Company IDs: {allowed_company_ids}")
        print(f"Allowed Client IDs: {allowed_client_ids}")
        print(f"-----------------------")

        # 3. بناء الاستعلام الأساسي مع تطبيق المنطق المطلوب بدقة
        base_query = db.query(Journey)
        if not is_superuser:
            # التعديل الهام: إذا وُجد عملاء محددين للمستخدم، تكون الصلاحية مقصورة على رحلات هؤلاء العملاء فقط حصراً
            if allowed_client_ids:
                base_query = base_query.filter(Journey.client_company_id.in_(allowed_client_ids))
            elif allowed_company_ids:
                # إذا لم يُحدد عملاء، ولكن قُدّرت شركة نقل، تظهر رحلات الشركة فقط
                base_query = base_query.outerjoin(Vehicle, Journey.vehicle_id == Vehicle.id).filter(
                    or_(
                        Vehicle.company_id.in_(allowed_company_ids),
                        Journey.company_id.in_(allowed_company_ids)
                    )
                )
            else:
                # إذا لم يكن لديه لا شركات ولا عملاء مرتبطين، لا تظهر أي رحلات
                base_query = base_query.filter(Journey.id == -1)

        total_journeys = base_query.count()
        active_journeys_count = base_query.filter(Journey.status == "Active").count()
        
        # جلب الإعدادات العامة (الافتراضية) كاحتياطي
        default_setting = db.query(Setting).filter(
            Setting.company_id == None,
            Setting.client_company_id == None
        ).first()
        
        default_follow_up = float(default_setting.follow_up_interval) if default_setting and default_setting.follow_up_interval else 2.0
        default_max_driving = float(default_setting.max_driving_hours) if default_setting and default_setting.max_driving_hours else 10.0
        default_sleep_time = float(default_setting.sleep_time) if default_setting and default_setting.sleep_time else 8.0

        company_settings_cache = {}

        def get_company_settings(company_id):
            if not company_id:
                return default_follow_up, default_max_driving, default_sleep_time
            if company_id in company_settings_cache:
                return company_settings_cache[company_id]
            
            comp_setting = db.query(Setting).filter(Setting.company_id == company_id).first()
            if not comp_setting:
                comp_setting = default_setting
            
            fu = float(comp_setting.follow_up_interval) if comp_setting and comp_setting.follow_up_interval else default_follow_up
            md = float(comp_setting.max_driving_hours) if comp_setting and comp_setting.max_driving_hours else default_max_driving
            st = float(comp_setting.sleep_time) if comp_setting and comp_setting.sleep_time else default_sleep_time
            
            company_settings_cache[company_id] = (fu, md, st)
            return fu, md, st

        active_journeys = base_query.filter(Journey.status.in_(["Active", "Planned"])).all()
        
        needs_followup_journeys = []
        driving_alerts = []

        for journey in active_journeys:
            comp_id = None
            if journey.vehicle and hasattr(journey.vehicle, 'company_id'):
                comp_id = journey.vehicle.company_id
            elif hasattr(journey, 'company_id'):
                comp_id = journey.company_id

            follow_up_interval, max_driving_hours, sleep_time_setting = get_company_settings(comp_id)

            interval_ago = datetime.now() - timedelta(hours=follow_up_interval)

            latest_followup = db.query(JourneyFollowUp).filter(
                JourneyFollowUp.journey_id == journey.id,
                JourneyFollowUp.is_deleted == False
            ).order_by(JourneyFollowUp.id.desc()).first()
            
            if not latest_followup:
                needs_followup_journeys.append(journey)
            else:
                followup_time = getattr(latest_followup, 'created_at', None)
                if followup_time and followup_time < interval_ago:
                    needs_followup_journeys.append(journey)

            if journey.status == "Active":
                start_time = getattr(journey, 'manual_start_time', None)
                if start_time:
                    total_hours = (datetime.now() - start_time).total_seconds() / 3600.0
                    days_passed = int(total_hours // 24)
                    net_driving_hours = total_hours - (days_passed * sleep_time_setting)

                    if net_driving_hours > max_driving_hours:
                        vehicle_plate = journey.vehicle.plate_number if journey.vehicle else "N/A"
                        driver_name = journey.driver.name if journey.driver else "N/A"
                        driving_alerts.append({
                            "journey_id": journey.id,
                            "vehicle": vehicle_plate,
                            "driver": driver_name,
                            "driving_hours": round(net_driving_hours, 1),
                            "max_allowed": max_driving_hours
                        })

        segmented_needs_followup = {}
        for journey in needs_followup_journeys:
            segment_name = "General / Unassigned"
            seg_obj = None
            if hasattr(journey, 'segment') and journey.segment:
                seg_obj = journey.segment
            elif journey.route:
                if hasattr(journey.route, 'segment') and journey.route.segment:
                    seg_obj = journey.route.segment
                elif hasattr(journey.route, 'segment_name') and journey.route.segment_name:
                    seg_obj = journey.route.segment_name
            
            if seg_obj:
                if isinstance(seg_obj, str):
                    segment_name = seg_obj
                else:
                    for attr in ['name', 'title', 'segment_name', 'label']:
                        if hasattr(seg_obj, attr) and getattr(seg_obj, attr):
                            segment_name = str(getattr(seg_obj, attr))
                            break
            
            if segment_name not in segmented_needs_followup:
                segmented_needs_followup[segment_name] = []
            segmented_needs_followup[segment_name].append(journey)

        return templates.TemplateResponse(
            request=request,
            name="dashboard.html",
            context={
                "total_journeys": total_journeys,
                "active_journeys_count": active_journeys_count,
                "needs_followup_count": len(needs_followup_journeys),
                "segmented_needs_followup": segmented_needs_followup,
                "driving_alerts": driving_alerts
            }
        )
    except Exception as e:
        print(f"Error in Dashboard: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    
# ----------------------------------------------------
# VEHICLES
# ----------------------------------------------------
@app.get("/vehicles")
def list_vehicles(
    request: Request, 
    q: Optional[str] = None, 
    company_id: Optional[str] = None, 
    client_id: Optional[str] = None,   
    segment_id: Optional[str] = None,  
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    # تحويل القيم القادمة من الفلتر بأمان إلى أعداد صحيحة أو None
    c_id = int(company_id) if company_id and company_id.isdigit() and int(company_id) != 0 else None
    cl_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
    s_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None

    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "role", "") in ["admin", "developer"]
    )

    if is_admin_or_dev:
        companies = db.query(Company).all()
        clients = db.query(Client).all()
        segments = db.query(Segment).all()
        query = db.query(Vehicle)
    else:
        # 1. جمع الشركات المربوطة بالمستخدم فقط
        allowed_company_ids = [c.id for c in getattr(current_user, "companies", [])]
        if hasattr(current_user, "company_id") and current_user.company_id:
            if current_user.company_id not in allowed_company_ids:
                allowed_company_ids.append(current_user.company_id)
                
        # فحص الجدول الوسيط للشركات إن وجد
        for table_name in ["user_companies", "user_company", "user_transport_companies"]:
            try:
                result = db.execute(
                    text(f"SELECT company_id FROM {table_name} WHERE user_id = :uid"), 
                    {"uid": current_user.id}
                ).fetchall()
                for row in result:
                    if row[0] not in allowed_company_ids:
                        allowed_company_ids.append(row[0])
            except Exception:
                db.rollback()

        # قصر قائمة الشركات في الفلتر على المربوطة بالمستخدم فقط
        companies = db.query(Company).filter(Company.id.in_(allowed_company_ids)).all() if allowed_company_ids else []

        # 2. جمع العملاء المربوطين بالمستخدم فقط (حصراً)
        allowed_client_ids = [c.id for c in getattr(current_user, "clients", [])]
        if hasattr(current_user, "client_company_id") and current_user.client_company_id:
            if current_user.client_company_id not in allowed_client_ids:
                allowed_client_ids.append(current_user.client_company_id)

        # فحص الجدول الوسيط للعملاء إن وجد
        for table_name in ["user_clients", "user_client"]:
            try:
                result = db.execute(
                    text(f"SELECT client_id FROM {table_name} WHERE user_id = :uid"), 
                    {"uid": current_user.id}
                ).fetchall()
                for row in result:
                    if row[0] not in allowed_client_ids:
                        allowed_client_ids.append(row[0])
            except Exception:
                db.rollback()

        # قصر قائمة العملاء في الفلتر على المربوطة بالمستخدم فقط (وليس كل عملاء الشركة)
        clients = db.query(Client).filter(Client.id.in_(allowed_client_ids)).all() if allowed_client_ids else []
        
        valid_client_ids = [cl.id for cl in clients]
        segments = db.query(Segment).filter(Segment.client_id.in_(valid_client_ids)).all() if valid_client_ids else []

        # 3. فلترة المركبات بناءً على العملاء المحددين للمستخدم أولاً ثم الشركات
        query = db.query(Vehicle)
        if allowed_client_ids:
            query = query.filter(Vehicle.client_id.in_(allowed_client_ids))
        elif allowed_company_ids:
            query = query.filter(Vehicle.company_id.in_(allowed_company_ids))
        else:
            query = query.filter(False) # لا توجد صلاحيات

    # تطبيق الفلاتر الإضافية (إن وجدت في الـ Request)
    if c_id:
        query = query.filter(Vehicle.company_id == c_id)
    if cl_id:
        query = query.filter(Vehicle.client_id == cl_id)
    if s_id:
        query = query.filter(Vehicle.segment_id == s_id)
    if q:
        query = query.filter(
            (Vehicle.plate_number.like(f"%{q}%")) | 
            (Vehicle.code.like(f"%{q}%")) |
            (Vehicle.phone_number.like(f"%{q}%"))
        )
        
    vehicles = query.all()
    
    return templates.TemplateResponse(
        request=request, 
        name="vehicles.html", 
        context={
            "vehicles": vehicles, 
            "companies": companies,
            "clients": clients,
            "segments": segments,
            "selected_company_id": c_id,
            "selected_client_id": cl_id,
            "selected_segment_id": s_id,
            "search_query": q or "",
            "current_user": current_user
        }
    )

@app.post("/vehicles/add")
def add_vehicle(
    request: Request,
    plate_number: str = Form(...), 
    phone_number: Optional[str] = Form(None),
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    vehicle_type: Optional[str] = Form(None),  
    company_id: int = Form(...),  
    client_id: Optional[str] = Form(None),
    segment_id: Optional[str] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    last = db.query(Vehicle.id).order_by(Vehicle.id.desc()).first()
    
    cl_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
    s_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None

    db.add(Vehicle(
        code=f"TRK-{(last[0] + 1) if last else 1:04d}", 
        plate_number=plate_number, 
        phone_number=phone_number,
        license_expiry=license_expiry if license_expiry else None, 
        status=status,
        vehicle_type=vehicle_type,  
        company_id=company_id,
        client_id=cl_id,
        segment_id=s_id
    ))
    db.commit()
    return RedirectResponse(url="/vehicles", status_code=303)

@app.post("/vehicles/update/{id}")
def update_vehicle(
    request: Request,
    id: int, 
    plate_number: str = Form(...), 
    phone_number: Optional[str] = Form(None),
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    vehicle_type: Optional[str] = Form(None),  
    company_id: int = Form(...),  
    client_id: Optional[str] = Form(None),
    segment_id: Optional[str] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    v = db.query(Vehicle).filter(Vehicle.id == id).first()
    if v:
        v.plate_number = plate_number
        v.phone_number = phone_number
        v.license_expiry = license_expiry if license_expiry else None
        v.status = status
        v.vehicle_type = vehicle_type  
        v.company_id = company_id
        v.client_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
        v.segment_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None
        db.commit()
    return RedirectResponse(url="/vehicles", status_code=303)

@app.post("/vehicles/delete/{id}")
def delete_vehicle(
    request: Request,
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    v = db.query(Vehicle).filter(Vehicle.id == id).first()
    if v:
        db.delete(v)
        db.commit()
    return RedirectResponse(url="/vehicles", status_code=303)

# ----------------------------------------------------
# TRAILERS
# ----------------------------------------------------
@app.get("/trailers")
def list_trailers(
    request: Request, 
    q: Optional[str] = None, 
    company_id: Optional[str] = None, 
    client_id: Optional[str] = None,   
    segment_id: Optional[str] = None,  
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    # تحويل القيم القادمة من الفلتر بأمان إلى أعداد صحيحة أو None
    c_id = int(company_id) if company_id and company_id.isdigit() and int(company_id) != 0 else None
    cl_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
    s_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None

    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "role", "") in ["admin", "developer"]
    )

    if is_admin_or_dev:
        companies = db.query(Company).all()
        clients = db.query(Client).all()
        segments = db.query(Segment).all()
        query = db.query(Trailer)
    else:
        allowed_company_ids = [c.id for c in getattr(current_user, "companies", [])]
        if hasattr(current_user, "company_id") and current_user.company_id:
            allowed_company_ids.append(current_user.company_id)
            
        # تقييد الشركات الظاهرة في الفلتر للشركة المحددة للمستخدم فقط
        companies = db.query(Company).filter(Company.id.in_(allowed_company_ids)).all() if allowed_company_ids else []
        
        allowed_client_ids = [c.id for c in getattr(current_user, "clients", [])]
        
        # تقييد العملاء الظاهرين في الفلتر للعملاء المربوطين بالمستخدم فقط
        clients = db.query(Client).filter(
            Client.id.in_(allowed_client_ids)
        ).all() if allowed_client_ids else []
        
        valid_client_ids = [cl.id for cl in clients]
        segments = db.query(Segment).filter(Segment.client_id.in_(valid_client_ids)).all() if valid_client_ids else []

        # قصر ظهور المقاطورات حصرياً على العملاء المربوطين بالـ user فقط
        if allowed_client_ids:
            query = db.query(Trailer).filter(
                (Trailer.client_id.in_(allowed_client_ids)) | 
                ((Trailer.company_id == None) & (Trailer.client_id == None))
            )
        else:
            query = db.query(Trailer).filter(
                (Trailer.company_id == None) & (Trailer.client_id == None)
            )

    # تطبيق الفلاتر باستخدام المتغيرات المحولة (c_id, cl_id, s_id)
    if c_id:
        query = query.filter(Trailer.company_id == c_id)
    if cl_id:
        query = query.filter(Trailer.client_id == cl_id)
    if s_id:
        query = query.filter(Trailer.segment_id == s_id)
    if q:
        query = query.filter(
            (Trailer.trailer_number.like(f"%{q}%")) | 
            (Trailer.code.like(f"%{q}%"))
        )
        
    trailers = query.all()
    
    return templates.TemplateResponse(
        request=request, 
        name="trailers.html", 
        context={
            "trailers": trailers, 
            "companies": companies,
            "clients": clients,
            "segments": segments,
            "selected_company_id": c_id,
            "selected_client_id": cl_id,
            "selected_segment_id": s_id,
            "search_query": q or "",
            "current_user": current_user
        }
    )

@app.post("/trailers/add")
def add_trailer(
    request: Request,
    trailer_number: str = Form(...), 
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    company_id: int = Form(...),  
    client_id: Optional[str] = Form(None),
    segment_id: Optional[str] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    last = db.query(Trailer.id).order_by(Trailer.id.desc()).first()
    
    cl_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
    s_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None

    db.add(Trailer(
        code=f"TLR-{(last[0] + 1) if last else 1:04d}", 
        trailer_number=trailer_number, 
        license_expiry=license_expiry if license_expiry else None, 
        status=status,
        company_id=company_id,  
        client_id=cl_id,
        segment_id=s_id
    ))
    db.commit()
    return RedirectResponse(url="/trailers", status_code=303)


@app.post("/trailers/update/{id}")
def update_trailer(
    request: Request,
    id: int, 
    trailer_number: str = Form(...), 
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    company_id: int = Form(...),  
    client_id: Optional[str] = Form(None),
    segment_id: Optional[str] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    t = db.query(Trailer).filter(Trailer.id == id).first()
    if t:
        t.trailer_number = trailer_number
        t.license_expiry = license_expiry if license_expiry else None
        t.status = status
        t.company_id = company_id
        t.client_id = int(client_id) if client_id and client_id.isdigit() and int(client_id) != 0 else None
        t.segment_id = int(segment_id) if segment_id and segment_id.isdigit() and int(segment_id) != 0 else None
        db.commit()
    return RedirectResponse(url="/trailers", status_code=303)

@app.post("/trailers/delete/{id}")
def delete_trailer(
    request: Request,
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    t = db.query(Trailer).filter(Trailer.id == id).first()
    if t:
        db.delete(t)
        db.commit()
    return RedirectResponse(url="/trailers", status_code=303)


# ----------------------------------------------------
# DRIVERS
# ----------------------------------------------------
@app.get("/drivers")
def list_drivers(
    request: Request, 
    q: Optional[str] = None, 
    company_id: Optional[int] = None,
    client_id: Optional[int] = None,
    segment_id: Optional[int] = None,
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "role", "") in ["admin", "developer"]
    )

    if is_admin_or_dev:
        companies = db.query(Company).all()
        clients = db.query(Client).all()
        segments = db.query(Segment).all()
        query = db.query(Driver)
    else:
        allowed_company_ids = [c.id for c in getattr(current_user, "companies", [])]
        if hasattr(current_user, "company_id") and current_user.company_id:
            allowed_company_ids.append(current_user.company_id)
            
        companies = db.query(Company).filter(Company.id.in_(allowed_company_ids)).all() if allowed_company_ids else []
        
        # العملاء المخصصون للمستخدم حصراً
        allowed_client_ids = [c.id for c in getattr(current_user, "clients", [])]
        
        if allowed_client_ids:
            clients = db.query(Client).filter(Client.id.in_(allowed_client_ids)).all()
        else:
            clients = []
            
        valid_client_ids = [cl.id for cl in clients]

        if valid_client_ids:
            segments = db.query(Segment).filter(Segment.client_id.in_(valid_client_ids)).all()
        else:
            segments = []

        company_filter = Driver.company_id.in_(allowed_company_ids) if allowed_company_ids else (Driver.company_id == None)
        
        if allowed_client_ids:
            client_filter = (Driver.client_id.in_(allowed_client_ids)) | (Driver.client_id == None)
        else:
            client_filter = (Driver.client_id == None)

        query = db.query(Driver).filter(
            company_filter,
            client_filter
        )

    # تطبيق الفلاتر والبحث الإضافية
    if company_id:
        query = query.filter(Driver.company_id == company_id)
    if client_id:
        query = query.filter(Driver.client_id == client_id)
    if segment_id:
        query = query.filter(Driver.segment_id == segment_id)
    if q:
        query = query.filter(
            (Driver.name.like(f"%{q}%")) | 
            (Driver.phone.like(f"%{q}%")) | 
            (Driver.code.like(f"%{q}%")) |
            (Driver.username.like(f"%{q}%"))
        )
        
    drivers = query.all()
    
    return templates.TemplateResponse(
        request=request, 
        name="drivers.html", 
        context={
            "drivers": drivers, 
            "companies": companies,
            "clients": clients,
            "segments": segments,
            "selected_company_id": company_id,
            "selected_client_id": client_id,
            "selected_segment_id": segment_id,
            "search_query": q or "",
            "current_user": current_user
        }
    )

@app.post("/drivers/add")
def add_driver(
    request: Request,
    name: str = Form(...), 
    phone: Optional[str] = Form(None), 
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    username: Optional[str] = Form(None),
    password: Optional[str] = Form(None),
    company_id: Optional[int] = Form(None),
    client_id: Optional[int] = Form(None),
    segment_id: Optional[int] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    last = db.query(Driver.id).order_by(Driver.id.desc()).first()
    c_id = company_id if company_id and company_id != 0 else None
    cl_id = client_id if client_id and client_id != 0 else None
    s_id = segment_id if segment_id and segment_id != 0 else None

    db.add(Driver(
        code=f"DRV-{(last[0] + 1) if last else 1:04d}", 
        name=name, 
        phone=phone, 
        license_expiry=license_expiry, 
        status=status,
        username=username if username else None,
        password=password if password else None,
        company_id=c_id,
        client_id=cl_id,
        segment_id=s_id
    ))
    db.commit()
    return RedirectResponse(url="/drivers", status_code=303)

@app.post("/drivers/update/{id}")
def update_driver(
    request: Request,
    id: int, 
    name: str = Form(...), 
    phone: Optional[str] = Form(None), 
    license_expiry: Optional[str] = Form(None), 
    status: str = Form("Active"),
    username: Optional[str] = Form(None),
    password: Optional[str] = Form(None),
    company_id: Optional[int] = Form(None),
    client_id: Optional[int] = Form(None),
    segment_id: Optional[int] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    d = db.query(Driver).filter(Driver.id == id).first()
    if d:
        d.name = name
        d.phone = phone
        d.license_expiry = license_expiry
        d.status = status
        d.username = username if username else None
        # تحديث كلمة المرور فقط إذا تم إدخال قيمة جديدة
        if password:
            d.password = password
        d.company_id = company_id if company_id and company_id != 0 else None
        d.client_id = client_id if client_id and client_id != 0 else None
        d.segment_id = segment_id if segment_id and segment_id != 0 else None
        db.commit()
    return RedirectResponse(url="/drivers", status_code=303)

@app.post("/drivers/delete/{id}")
def delete_driver(
    request: Request,
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    d = db.query(Driver).filter(Driver.id == id).first()
    if d:
        db.delete(d)
        db.commit()
    return RedirectResponse(url="/drivers", status_code=303)

from pydantic import BaseModel

class DriverLoginSchema(BaseModel):
    username: str
    password: str

@app.post("/api/driver/login")
def driver_login(data: DriverLoginSchema, db: Session = Depends(get_db)):
    driver = db.query(Driver).filter(Driver.username == data.username, Driver.password == data.password).first()
    
    if not driver:
        return {"status": "error", "message": "اسم المستخدم أو كلمة المرور غير صحيحة"}
    
    return {
        "status": "success",
        "driver_id": driver.id,
        "driver_name": driver.name if driver.name else driver.username,
        "message": "تم تسجيل الدخول بنجاح"
    }

import io
import urllib.parse
from datetime import datetime, timezone
from typing import Optional
from fastapi import Request, Depends, HTTPException, Form, status
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from sqlalchemy.orm import Session
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ----------------------------------------------------
# 1. fleet overview page (GET)
# ----------------------------------------------------

@app.get("/fleet-overview", response_class=HTMLResponse)
def fleet_overview_page(
    request: Request,
    company_id: Optional[str] = None,
    client_id: Optional[str] = None,
    segment_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

        # لو الـ user مش متحددله شركة أو عميل ميبانش أي بيانات
        if not user_company_ids or not user_client_ids:
            return templates.TemplateResponse(
                request=request,
                name="fleet_overview.html",
                context={
                    "current_user": current_user,
                    "fleet_data": [],
                    "companies": [],
                    "clients": [],
                    "segments": [],
                    "vehicles": [],
                    "drivers": [],
                    "trailers": [],
                    "selected_company": None,
                    "selected_client": None,
                    "selected_segment": None
                }
            )

        c_id = int(company_id) if company_id and company_id.isdigit() else None
        cl_id = int(client_id) if client_id and client_id.isdigit() else None
        s_id = int(segment_id) if segment_id and segment_id.isdigit() else None

        if c_id and c_id not in user_company_ids:
            c_id = None
        if cl_id and cl_id not in user_client_ids:
            cl_id = None

        # 1. فلترة الشركات المسموحة
        companies = db.query(Company).filter(Company.id.in_(user_company_ids)).all()

        # 2. فلترة العملاء المسموحين بناءً على الشركة المختارة
        clients_query = db.query(Client).filter(Client.id.in_(user_client_ids))
        if c_id:
            clients_query = clients_query.filter(Client.company_id == c_id)
        elif len(user_company_ids) == 1:
            clients_query = clients_query.filter(Client.company_id == user_company_ids[0])
        clients = clients_query.all()

        effective_company_id = c_id if c_id else (user_company_ids[0] if len(user_company_ids) == 1 else None)
        effective_client_id = cl_id if cl_id else (user_client_ids[0] if len(user_client_ids) == 1 else None)

        # 3. جلب الـ Segments المربوطة بالعميل المحدد (Effective Client)
        segments_query = db.query(Segment).filter(Segment.client_id.in_(user_client_ids))
        if effective_client_id:
            segments_query = segments_query.filter(Segment.client_id == effective_client_id)
        segments = segments_query.all()
        
        # استخراج أيدي الـ Segments التابعة لهذا العميل للفلترة الدقيقة
        valid_segment_ids = [s.id for s in segments]

        # 4. جلب السيارات، السائقين والمقطورات المربوطة بالعميل أو بقطاعاته (Segments) المحددة
        vehicles_query = db.query(Vehicle).filter(Vehicle.client_id.in_(user_client_ids))
        drivers_query = db.query(Driver).filter(Driver.client_id.in_(user_client_ids))
        trailers_query = db.query(Trailer).filter(Trailer.client_id.in_(user_client_ids))

        if effective_client_id:
            vehicles_query = vehicles_query.filter(Vehicle.client_id == effective_client_id)
            drivers_query = drivers_query.filter(Driver.client_id == effective_client_id)
            trailers_query = trailers_query.filter(Trailer.client_id == effective_client_id)

        # إذا تم اختيار Segment معين، نقوم بتصفيته بدقة، وإلا تظهر كل سيارات وسائقين العميل
        if s_id and s_id in valid_segment_ids:
            vehicles_query = vehicles_query.filter(Vehicle.segment_id == s_id)
            drivers_query = drivers_query.filter(Driver.segment_id == s_id)
            trailers_query = trailers_query.filter(Trailer.segment_id == s_id)

        # تحويل النتائج إلى قواميس آمنة لتفادي أخطاء الـ JSON وعرضها بالقوائم المنسدلة
        vehicles = [{"id": v.id, "plate_number": v.plate_number, "company_id": v.company_id, "client_id": v.client_id, "segment_id": v.segment_id} for v in vehicles_query.all()]
        drivers = [{"id": d.id, "name": d.name, "company_id": d.company_id, "client_id": d.client_id, "segment_id": d.segment_id} for d in drivers_query.all()]
        trailers = [{"id": t.id, "trailer_number": t.trailer_number, "company_id": t.company_id, "client_id": t.client_id, "segment_id": t.segment_id} for t in trailers_query.all()]

        # 5. جلب التسكينات الأساسية للجدول مع دعم الفلترة الهرمية
        query = db.query(FleetAssignment).filter(
            FleetAssignment.company_id.in_(user_company_ids),
            FleetAssignment.client_id.in_(user_client_ids)
        )

        if effective_company_id:
            query = query.filter(FleetAssignment.company_id == effective_company_id)
        if effective_client_id:
            query = query.filter(FleetAssignment.client_id == effective_client_id)
        if s_id and s_id in valid_segment_ids:
            query = query.filter(FleetAssignment.segment_id == s_id)
            
        assignments = query.all()
        
        fleet_data = []
        for a in assignments:
            if a.company_id not in user_company_ids or a.client_id not in user_client_ids:
                continue

            v = db.query(Vehicle).filter(Vehicle.id == a.vehicle_id).first()
            d = db.query(Driver).filter(Driver.id == a.driver_id).first()
            t = db.query(Trailer).filter(Trailer.id == a.trailer_id).first() if a.trailer_id else None
            comp = db.query(Company).filter(Company.id == a.company_id).first()
            client = db.query(Client).filter(Client.id == a.client_id).first()
            seg = db.query(Segment).filter(Segment.id == a.segment_id).first()

            fleet_data.append({
                "id": a.id,
                "vehicle_plate": v.plate_number if v else "N/A",
                "vehicle_type": getattr(v, 'vehicle_type', 'Truck') if v else "Truck",
                "vehicle_phone": getattr(v, 'phone_number', 'N/A') if v else "N/A",
                "vehicle_license": getattr(v, 'license_expiry', 'N/A') if v else "N/A",
                "vehicle_id": a.vehicle_id,
                "driver_name": d.name if d else "N/A",
                "driver_phone": getattr(d, 'phone', 'N/A') if d else "N/A",
                "driver_license": getattr(d, 'license_expiry', 'N/A') if d else "N/A",
                "driver_id": a.driver_id,
                "trailer_number": t.trailer_number if t else None,
                "trailer_license": getattr(t, 'license_expiry', 'N/A') if t else None,
                "trailer_id": a.trailer_id,
                "company_name": comp.name if comp else "N/A",
                "company_id": a.company_id,
                "client_name": client.name if client else "N/A",
                "client_id": a.client_id,
                "segment_name": seg.name if seg else "N/A",
                "segment_id": a.segment_id
            })

        return templates.TemplateResponse(
            request=request,
            name="fleet_overview.html",
            context={
                "current_user": current_user,
                "fleet_data": fleet_data,
                "companies": companies,
                "clients": clients,
                "segments": segments,
                "vehicles": vehicles,
                "drivers": drivers,
                "trailers": trailers,
                "selected_company": effective_company_id,
                "selected_client": effective_client_id,
                "selected_segment": s_id
            }
        )
    except Exception as e:
        print(f"Error in /fleet-overview: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

# ----------------------------------------------------
# 2. مسار حفظ أو تحديث التسكين اليدوي (POST)
# ----------------------------------------------------

@app.post("/fleet-overview/save")
def save_fleet_assignment(
    request: Request,
    assignment_id: Optional[int] = Form(None),
    company_id: int = Form(...),
    client_id: Optional[int] = Form(None),
    segment_id: Optional[int] = Form(None),
    vehicle_id: int = Form(...),
    driver_id: int = Form(...),
    trailer_id: Optional[int] = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

        if not user_company_ids or not user_client_ids:
            err_msg = "المستخدم غير مسموح له بإدارة أي شركات أو عملاء!"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        if company_id not in user_company_ids or (client_id and client_id not in user_client_ids):
            err_msg = "ليس لديك صلاحية لاختيار هذه الشركة أو العميل!"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        v_obj = db.query(Vehicle).filter(Vehicle.id == vehicle_id, Vehicle.company_id.in_(user_company_ids), Vehicle.client_id.in_(user_client_ids)).first()
        d_obj = db.query(Driver).filter(Driver.id == driver_id, Driver.company_id.in_(user_company_ids), Driver.client_id.in_(user_client_ids)).first()
        
        if not v_obj:
            err_msg = "السيارة المحددة لا تتبع صلاحياتك المعتمدة!"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)
            
        if not d_obj:
            err_msg = "السائق المحدد لا يتبع صلاحياتك المعتمدة!"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        if trailer_id:
            t_obj = db.query(Trailer).filter(Trailer.id == trailer_id, Trailer.company_id.in_(user_company_ids), Trailer.client_id.in_(user_client_ids)).first()
            if not t_obj:
                err_msg = "المقطورة المحددة لا تتبع صلاحياتك المعتمدة!"
                return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        def get_conflict_details(assignment):
            v = db.query(Vehicle).filter(Vehicle.id == assignment.vehicle_id).first()
            d = db.query(Driver).filter(Driver.id == assignment.driver_id).first()
            comp = db.query(Company).filter(Company.id == assignment.company_id).first()
            client = db.query(Client).filter(Client.id == assignment.client_id).first() if assignment.client_id else None
            
            v_plate = v.plate_number if v else "سيارة غير معروفة"
            d_name = d.name if d else "سائق غير معروف"
            c_name = comp.name if comp else "شركة غير معروفة"
            cl_name = f" - العميل: {client.name}" if client else ""
            return f"[الشركة: {c_name}{cl_name} | السيارة: {v_plate} | السائق: {d_name}]"

        existing_vehicle = db.query(FleetAssignment).filter(
            FleetAssignment.vehicle_id == vehicle_id,
            FleetAssignment.id != assignment_id if assignment_id else True
        ).first()
        
        if existing_vehicle:
            plate = v_obj.plate_number if v_obj else "السيارة"
            conflict_info = get_conflict_details(existing_vehicle)
            err_msg = f"هذه السيارة ({plate}) مسجلة بالفعل في تسكين آخر: {conflict_info}"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        existing_driver = db.query(FleetAssignment).filter(
            FleetAssignment.driver_id == driver_id,
            FleetAssignment.id != assignment_id if assignment_id else True
        ).first()
        
        if existing_driver:
            d_name = d_obj.name if d_obj else "السائق"
            conflict_info = get_conflict_details(existing_driver)
            err_msg = f"هذا السائق ({d_name}) مسجل بالفعل في تسكين آخر: {conflict_info}"
            return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        if trailer_id:
            existing_trailer = db.query(FleetAssignment).filter(
                FleetAssignment.trailer_id == trailer_id,
                FleetAssignment.id != assignment_id if assignment_id else True
            ).first()
            
            if existing_trailer:
                t_obj = db.query(Trailer).filter(Trailer.id == trailer_id).first()
                t_num = t_obj.trailer_number if t_obj else "المقطورة"
                conflict_info = get_conflict_details(existing_trailer)
                err_msg = f"هذه المقطورة ({t_num}) مسجلة بالفعل في تسكين آخر: {conflict_info}"
                return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)

        if assignment_id:
            assignment = db.query(FleetAssignment).filter(
                FleetAssignment.id == assignment_id,
                FleetAssignment.company_id.in_(user_company_ids),
                FleetAssignment.client_id.in_(user_client_ids)
            ).first()
            if assignment:
                assignment.company_id = company_id
                assignment.client_id = client_id if client_id else None
                assignment.segment_id = segment_id if segment_id else None
                assignment.vehicle_id = vehicle_id
                assignment.driver_id = driver_id
                assignment.trailer_id = trailer_id if trailer_id else None
        else:
            new_assignment = FleetAssignment(
                company_id=company_id,
                client_id=client_id if client_id else None,
                segment_id=segment_id if segment_id else None,
                vehicle_id=vehicle_id,
                driver_id=driver_id,
                trailer_id=trailer_id if trailer_id else None
            )
            db.add(new_assignment)

        db.commit()
        return RedirectResponse(url="/fleet-overview", status_code=303)

    except Exception as e:
        db.rollback()
        err_msg = str(e)
        return RedirectResponse(url=f"/fleet-overview?error={urllib.parse.quote(err_msg)}", status_code=303)


# ----------------------------------------------------
# 3. مسار حذف التسكين (GET)
# ----------------------------------------------------

@app.get("/fleet-overview/delete/{assignment_id}")
def delete_fleet_assignment(
    assignment_id: int, 
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

    assignment = db.query(FleetAssignment).filter(
        FleetAssignment.id == assignment_id,
        FleetAssignment.company_id.in_(user_company_ids),
        FleetAssignment.client_id.in_(user_client_ids)
    ).first()
    
    if assignment:
        db.delete(assignment)
        db.commit()
    return RedirectResponse(url="/fleet-overview", status_code=status.HTTP_303_SEE_OTHER)


# ----------------------------------------------------
# 4. مسار تصدير البيانات إلى Excel (GET)
# ----------------------------------------------------

@app.get("/fleet-overview/export")
def export_fleet_excel(
    company_id: Optional[str] = None,
    client_id: Optional[str] = None,
    segment_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

    if not user_company_ids or not user_client_ids:
        raise HTTPException(status_code=403, detail="Unauthorized")

    query = db.query(FleetAssignment).filter(
        FleetAssignment.company_id.in_(user_company_ids),
        FleetAssignment.client_id.in_(user_client_ids)
    )
    
    if company_id and company_id.isdigit():
        c_id = int(company_id)
        if c_id in user_company_ids:
            query = query.filter(FleetAssignment.company_id == c_id)
    if client_id and client_id.isdigit():
        cl_id = int(client_id)
        if cl_id in user_client_ids:
            query = query.filter(FleetAssignment.client_id == cl_id)
    if segment_id and segment_id.isdigit():
        query = query.filter(FleetAssignment.segment_id == int(segment_id))
        
    assignments = query.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Fleet Assignments"
    ws.views.sheetView[0].rightToLeft = True

    header_fill = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1B263B")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="555555")
    
    align_center = Alignment(horizontal="center", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    ws.merge_cells("B1:M1")
    ws["B1"] = "Fleet Assignments Overview Report"
    ws["B1"].font = title_font
    ws["B1"].alignment = align_center

    ws.merge_cells("B2:M2")
    ws["B2"] = f"Exported At: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M')}"
    ws["B2"].font = subtitle_font
    ws["B2"].alignment = align_center

    ws.append([])

    headers = [
        "#", "Vehicle Plate", "Vehicle Type", "Vehicle Phone", "Vehicle License", 
        "Driver Name", "Driver Phone", "Driver License", 
        "Trailer Number", "Trailer License", 
        "Transport Company", "Client Company", "Segment / Branch"
    ]
    ws.append(headers)

    header_row_idx = 4
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row_idx, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    for idx, a in enumerate(assignments, 1):
        v = db.query(Vehicle).filter(Vehicle.id == a.vehicle_id).first()
        d = db.query(Driver).filter(Driver.id == a.driver_id).first()
        t = db.query(Trailer).filter(Trailer.id == a.trailer_id).first() if a.trailer_id else None
        comp = db.query(Company).filter(Company.id == a.company_id).first()
        client = db.query(Client).filter(Client.id == a.client_id).first()
        seg = db.query(Segment).filter(Segment.id == a.segment_id).first()

        row_data = [
            idx,
            v.plate_number if v else "N/A",
            getattr(v, 'vehicle_type', 'Truck') if v else "Truck",
            str(getattr(v, 'phone_number', 'N/A')) if v else "N/A",
            getattr(v, 'license_expiry', 'N/A') if v else "N/A",
            d.name if d else "N/A",
            str(getattr(d, 'phone', 'N/A')) if d else "N/A",
            getattr(d, 'license_expiry', 'N/A') if v else "N/A",
            t.trailer_number if t else "No Trailer",
            getattr(t, 'license_expiry', 'N/A') if t else "N/A",
            comp.name if comp else "N/A",
            client.name if client else "N/A",
            seg.name if seg else "N/A"
        ]
        ws.append(row_data)

        current_row = ws.max_row
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.border = thin_border
            cell.alignment = align_center

    for col in ws.columns:
        max_len = 0
        col_letter = None
        for cell in col:
            if cell.row > 2 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
            if not col_letter:
                from openpyxl.utils import get_column_letter
                col_letter = get_column_letter(cell.column)
        if col_letter:
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=fleet_assignments_report.xlsx"}
    )

# --- مسارات إجازات السائقين (Driver Leaves) ---
@app.get("/driver-leaves", response_class=HTMLResponse)
def driver_leaves_page(
    request: Request,
    company_id: Optional[str] = None,
    client_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

        if not user_company_ids:
            return templates.TemplateResponse(
                request=request, name="driver_leaves.html",
                context={"current_user": current_user, "leaves": [], "companies": [], "clients": [], "drivers": [], "active_leave_ids": [], "selected_company": None, "selected_client": None}
            )

        c_id = int(company_id) if company_id and company_id.isdigit() else None
        cl_id = int(client_id) if client_id and client_id.isdigit() else None

        if c_id and c_id not in user_company_ids: c_id = None
        if cl_id and cl_id not in user_client_ids: cl_id = None

        companies = db.query(Company).filter(Company.id.in_(user_company_ids)).all()

        # جلب العملاء المسموحين للمستخدم فقط والمحددين في صفحة الـ Users
        clients_query = db.query(Client)
        if user_client_ids:
            clients_query = clients_query.filter(Client.id.in_(user_client_ids))
        if c_id: 
            clients_query = clients_query.filter(Client.company_id == c_id)
        elif len(user_company_ids) == 1: 
            clients_query = clients_query.filter(Client.company_id == user_company_ids[0])
        clients = clients_query.all()

        effective_company_id = c_id if c_id else (user_company_ids[0] if len(user_company_ids) == 1 else None)
        effective_client_id = cl_id if cl_id else None

        # استخراج أسماء العملاء المسموحين للمستخدم لتقييد البيانات بهم
        permitted_client_names = [cl.name for cl in clients]

        # بناء استعلام الإجازات بناءً على صلاحيات المستخدم والعملاء المخصصين له
        leaves_query = db.query(DriverLeave).filter(DriverLeave.company_id.in_(user_company_ids))
        if effective_company_id: 
            leaves_query = leaves_query.filter(DriverLeave.company_id == effective_company_id)
        
        if effective_client_id:
            selected_client_obj = db.query(Client).filter(Client.id == effective_client_id).first()
            if selected_client_obj:
                leaves_query = leaves_query.filter(DriverLeave.last_client == selected_client_obj.name)
        elif user_client_ids:
            # إذا لم يتم اختيار فلتر معين، يتم قصر البيانات على العملاء المربوطين بالمستخدم حصراً
            leaves_query = leaves_query.filter(DriverLeave.last_client.in_(permitted_client_names))

        leaves = leaves_query.all()

        active_leave_driver_ids = [l.driver_id for l in leaves if not l.end_date or l.status == 'Active']

        # تقييد قائمة السائقين بنفس الصلاحيات والعملاء المخصصين
        drivers_query = db.query(Driver).filter(Driver.company_id.in_(user_company_ids))
        if effective_company_id: drivers_query = drivers_query.filter(Driver.company_id == effective_company_id)
        if effective_client_id: 
            drivers_query = drivers_query.filter(Driver.client_id == effective_client_id)
        elif user_client_ids:
            drivers_query = drivers_query.filter(Driver.client_id.in_(user_client_ids))
            
        drivers = drivers_query.all()

        return templates.TemplateResponse(
            request=request, name="driver_leaves.html",
            context={
                "current_user": current_user, "leaves": leaves, "companies": companies,
                "clients": clients, "drivers": drivers, "active_leave_ids": active_leave_driver_ids,
                "selected_company": effective_company_id, "selected_client": effective_client_id
            }
        )
    except Exception as e:
        print(f"Error in /driver-leaves: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/driver-leaves/save")
def save_driver_leave(
    driver_id: int = Form(...),
    company_id: int = Form(...),
    start_date: str = Form(...),
    end_date: Optional[str] = Form(None),   
    notes: Optional[str] = Form(None),         
    last_client: Optional[str] = Form(None),   
    db: Session = Depends(get_db)
):
    try:
        days_count = None
        if end_date and end_date.strip():
            d1 = datetime.strptime(start_date, "%Y-%m-%d")
            d2 = datetime.strptime(end_date, "%Y-%m-%d")
            days_count = (d2 - d1).days + 1
            if days_count <= 0:
                raise HTTPException(status_code=400, detail="تاريخ العودة يجب أن يكون بعد أو يوافق تاريخ البداية!")

        # التحقق إذا كان السائق في إجازة نشطة مسجلة مسبقاً
        existing_leave = db.query(DriverLeave).filter(
            DriverLeave.driver_id == driver_id,
            (DriverLeave.end_date == None) | (DriverLeave.status == 'Active')
        ).first()
        
        if existing_leave:
            raise HTTPException(status_code=400, detail="هذا السائق في إجازة مسجلة مسبقاً ولم يعود بعد!")

        new_leave = DriverLeave(
            driver_id=driver_id,
            company_id=company_id,
            start_date=start_date,
            end_date=end_date if end_date and end_date.strip() else None,
            days_count=days_count,
            notes=notes,
            last_client=last_client,
            status="Active"
        )
        db.add(new_leave)
        db.commit()
        
        return RedirectResponse(url="/driver-leaves", status_code=303)
    except Exception as e:
        db.rollback()
        print(f"Error saving driver leave: {str(e)}") # لمعرفة سبب الخطأ بالتفصيل في الـ Terminal
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/driver-leaves/return/{leave_id}")
def return_driver_leave(
    leave_id: int,
    end_date: str = Form(...),
    db: Session = Depends(get_db)
):
    try:
        leave = db.query(DriverLeave).filter(DriverLeave.id == leave_id).first()
        if not leave:
            raise HTTPException(status_code=404, detail="الإجازة غير موجودة!")

        d1 = datetime.strptime(leave.start_date, "%Y-%m-%d")
        d2 = datetime.strptime(end_date, "%Y-%m-%d")
        days_count = (d2 - d1).days + 1

        if days_count <= 0:
            raise HTTPException(status_code=400, detail="تاريخ العودة يجب أن يكون بعد أو يوافق تاريخ البداية!")

        # تحديث بيانات العودة والحالة
        leave.end_date = end_date
        leave.days_count = days_count
        leave.status = "Returned"  # تم العودة من الإجازة
        db.commit()

        return RedirectResponse(url="/driver-leaves", status_code=303)
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/driver-leaves/delete/{leave_id}")
def delete_driver_leave(leave_id: int, db: Session = Depends(get_db)):
    try:
        leave = db.query(DriverLeave).filter(DriverLeave.id == leave_id).first()
        if leave:
            db.delete(leave)
            db.commit()
        return RedirectResponse(url="/driver-leaves", status_code=303)
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))

# ====================================================
# DRIVER LEAGUE
# ====================================================
@app.get("/drivers/league", response_class=HTMLResponse)
def driver_league(
    request: Request, 
    start_date: Optional[str] = None, 
    end_date: Optional[str] = None, 
    driver_name: Optional[str] = None, 
    company_id: Optional[str] = None,
    client_id: Optional[str] = None,
    segment_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        user_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        user_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

        default_start = "2026-08-01T00:00"
        default_end = datetime.now().strftime("%Y-%m-%dT23:59")

        if not user_company_ids:
            return templates.TemplateResponse(
                request=request,
                name="driver_league.html",
                context={
                    "drivers": [], "companies": [], "clients": [], "segments": [],
                    "start_date": start_date or default_start, "end_date": end_date or default_end,
                    "driver_name": driver_name or "", "selected_company": None, "selected_client": None, "selected_segment": None
                }
            )

        if not start_date:
            start_date = default_start
        if not end_date:
            end_date = default_end

        c_id = int(company_id) if company_id and company_id.isdigit() else None
        cl_id = int(client_id) if client_id and client_id.isdigit() else None
        seg_id = int(segment_id) if segment_id and segment_id.isdigit() else None

        if c_id and c_id not in user_company_ids: c_id = None
        if cl_id and cl_id not in user_client_ids: cl_id = None

        companies = db.query(Company).filter(Company.id.in_(user_company_ids)).all()
        
        clients_query = db.query(Client)
        if user_client_ids:
            clients_query = clients_query.filter(Client.id.in_(user_client_ids))
        if c_id:
            clients_query = clients_query.filter(Client.company_id == c_id)
        clients = clients_query.all()

        segments = []
        if cl_id:
            segments_query = db.query(Segment)
            if hasattr(Segment, 'client_id'):
                segments_query = segments_query.filter(Segment.client_id == cl_id)
            segments = segments_query.all()

        drivers_query = db.query(Driver).filter(Driver.company_id.in_(user_company_ids))
        if user_client_ids:
            drivers_query = drivers_query.filter(Driver.client_id.in_(user_client_ids))
        if c_id:
            drivers_query = drivers_query.filter(Driver.company_id == c_id)
        if cl_id:
            drivers_query = drivers_query.filter(Driver.client_id == cl_id)
        if seg_id:
            drivers_query = drivers_query.filter(Driver.segment_id == seg_id)

        all_drivers = drivers_query.all()
        allowed_driver_ids = [d.id for d in all_drivers]

        query = db.query(Journey).filter(Journey.status != "Cancelled")
        if allowed_driver_ids:
            query = query.filter(Journey.driver_id.in_(allowed_driver_ids))
        else:
            query = query.filter(Journey.driver_id.in_([-1]))

        # تصفية الرحلات بالتاريخ والوقت البدء (From Date & Time)
        if start_date:
            try:
                s_val = start_date.replace('T', ' ')
                if len(s_val) == 10:
                    s_val += " 00:00:00"
                dt_from = datetime.fromisoformat(s_val)
                query = query.filter(Journey.manual_start_time >= dt_from)
            except ValueError:
                pass

        # تصفية الرحلات بالتاريخ والوقت النهاية (To Date & Time)
        if end_date:
            try:
                e_val = end_date.replace('T', ' ')
                if len(e_val) == 10:
                    e_val += " 23:59:59"
                dt_to = datetime.fromisoformat(e_val)
                query = query.filter(Journey.manual_start_time <= dt_to)
            except ValueError:
                pass

        journeys = query.all()
        drivers_stats = {}

        for d in all_drivers:
            segment_name = d.segment.name if hasattr(d, 'segment') and d.segment else "-"
            drivers_stats[d.id] = {
                "id": d.id,
                "name": d.name,
                "code": getattr(d, 'code', f"DRV-{d.id:04d}"),
                "segment_name": segment_name,
                "km": 0.0,
                "total_diff_seconds": 0.0,      
                "abs_time_diff_seconds": 0.0,   
                "completed_journeys": 0,
                "loaded_journeys_count": 0,
                "first_journey_time": None,
                "last_journey_time": None
            }

        setting = db.query(Setting).filter(
            Setting.company_id == None,
            Setting.client_company_id == None
        ).first()
        speed = float(setting.speed) if setting and hasattr(setting, 'speed') and setting.speed else 40.0
        rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time else 0.5
        sleep_hours_setting = float(setting.sleep_time) if setting and hasattr(setting, 'sleep_time') and setting.sleep_time else 8.0

        for j in journeys:
            if not j.driver_id or j.driver_id not in drivers_stats:
                continue

            trip_km = float(j.route.distance_km) if j.route and j.route.distance_km else 0.0
            drivers_stats[j.driver_id]["km"] += trip_km
            drivers_stats[j.driver_id]["completed_journeys"] += 1

            # تحديد أول وأخر رحلة للسائق ضمن النطاق الزمني المختار
            journey_time = j.manual_start_time
            if journey_time:
                stats = drivers_stats[j.driver_id]
                if stats["first_journey_time"] is None or journey_time < stats["first_journey_time"]:
                    stats["first_journey_time"] = journey_time
                if stats["last_journey_time"] is None or journey_time > stats["last_journey_time"]:
                    stats["last_journey_time"] = journey_time

            is_empty = False
            if hasattr(j, 'load_type') and j.load_type:
                load_name = getattr(j.load_type, 'name', '') or str(j.load_type)
                if 'فارغ' in load_name.lower() or 'empty' in load_name.lower():
                    is_empty = True
            
            if not is_empty:
                drivers_stats[j.driver_id]["loaded_journeys_count"] += 1

            if j.route and j.manual_start_time:
                route_obj = j.route
                driving_hours_val = getattr(route_obj, 'driving_hours', None)
                if not driving_hours_val and route_obj.distance_km:
                    driving_hours_val = route_obj.distance_km / speed if speed > 0 else 0.0
                driving_hours_val = driving_hours_val or 0.0

                rest_time_val = getattr(route_obj, 'rest_time', None)
                if rest_time_val is None and route_obj.distance_km:
                    rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
                    rest_time_val = rest_stops * rest_duration
                rest_time_val = rest_time_val or 0.0

                total_time_val = getattr(route_obj, 'total_time', None)
                if not total_time_val:
                    total_time_val = driving_hours_val + rest_time_val

                total_planned_driving = total_time_val
                if getattr(j, 'include_sleep', 'N') == "Y":
                    total_planned_driving += sleep_hours_setting

                extra_off_duty = 24.0 if getattr(j, 'is_off_duty', 'N') == "Y" else 0.0
                planned_end = j.manual_start_time + timedelta(hours=total_planned_driving + extra_off_duty)
            else:
                planned_end = j.planned_end_time

            is_off_duty_val = getattr(j, 'is_off_duty', 'N') == "Y"
            ref_end_time = j.manual_end_time if is_off_duty_val else j.manual_arrived_time

            if j.manual_start_time and ref_end_time and planned_end:
                diff_delay = ref_end_time - planned_end
                sec_val = float(diff_delay.total_seconds())
                
                drivers_stats[j.driver_id]["total_diff_seconds"] += sec_val
                drivers_stats[j.driver_id]["abs_time_diff_seconds"] += abs(sec_val)

        processed_drivers = []
        for d_id, stats in drivers_stats.items():
            km_points = stats["km"] * 0.02              
            journey_points = stats["loaded_journeys_count"] * 5.0 
            
            total_hours_algebraic = stats["total_diff_seconds"] / 3600.0
            
            if total_hours_algebraic > 0:
                time_penalty = total_hours_algebraic * 2.0
                time_bonus = 0.0
            else:
                time_penalty = 0.0
                time_bonus = abs(total_hours_algebraic) * 1.5
            
            final_score = round(km_points + journey_points + time_bonus - time_penalty, 1)
            if final_score < 0:
                final_score = 0.0

            total_sec = stats["total_diff_seconds"]
            sign_prefix = "-" if total_sec < 0 else "+"
            abs_total_sec = int(round(abs(total_sec)))
            
            h_val = abs_total_sec // 3600
            m_val = (abs_total_sec % 3600) // 60
            
            formatted_time_diff = f"{sign_prefix}{h_val:02d}:{m_val:02d}"

            processed_drivers.append({
                "id": stats["id"],
                "name": stats["name"],
                "code": stats["code"],
                "segment_name": stats["segment_name"],
                "km": int(stats["km"]),
                "time_diff_hours": formatted_time_diff,
                "journey_count": stats["loaded_journeys_count"],
                "first_journey": stats["first_journey_time"].strftime("%Y-%m-%d %H:%M") if stats["first_journey_time"] else "-",
                "last_journey": stats["last_journey_time"].strftime("%Y-%m-%d %H:%M") if stats["last_journey_time"] else "-",
                "score": final_score
            })

        sorted_drivers = sorted(processed_drivers, key=lambda x: x["score"], reverse=True)
        for index, driver_data in enumerate(sorted_drivers, start=1):
            driver_data["rank"] = index

        filtered_drivers = sorted_drivers
        if driver_name and driver_name.strip():
            search_term = driver_name.strip().lower()
            filtered_drivers = [d for d in sorted_drivers if search_term in d["name"].lower()]

        return templates.TemplateResponse(
            request=request,
            name="driver_league.html",
            context={
                "drivers": filtered_drivers,
                "companies": companies,
                "clients": clients,
                "segments": segments,
                "start_date": start_date,
                "end_date": end_date,
                "driver_name": driver_name or "",
                "selected_company": c_id,
                "selected_client": cl_id,
                "selected_segment": seg_id
            }
        )
    except Exception as e:
        print(f"Error in /drivers/league: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


# ----------------------------------------------------
# EMPLOYEES LEAGUE (Filtered strictly by User Permissions & Company/Client)
# ----------------------------------------------------
@app.get("/employees/league")
def employees_league(
    request: Request,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    employee_name: Optional[str] = None,
    company_id: Optional[str] = None,
    client_company_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    comp_id = int(company_id) if company_id and company_id.isdigit() else None
    client_comp_id = int(client_company_id) if client_company_id and client_company_id.isdigit() else None

    today = date.today()
    if not from_date:
        from_date = today.replace(day=1).strftime("%Y-%m-%d")
    if not to_date:
        to_date = today.strftime("%Y-%m-%d")

    try:
        start_dt = datetime.strptime(from_date, "%Y-%m-%d")
        end_dt = datetime.strptime(to_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    except ValueError:
        start_dt = datetime(today.year, today.month, 1)
        end_dt = datetime(today.year, today.month, today.day, 23, 59, 59)

    from sqlalchemy import inspect
    inspector = inspect(db.bind)
    table_names = inspector.get_table_names()

    current_user_id = getattr(current_user, "id", None) if current_user else None

    allowed_company_ids = []
    allowed_client_ids = []

    if current_user_id:
        if "user_companies" in table_names:
            uc_allowed = db.execute(text("SELECT company_id FROM user_companies WHERE user_id = :uid"), {"uid": current_user_id}).fetchall()
            allowed_company_ids = [r[0] for r in uc_allowed]
        
        if "user_clients" in table_names:
            ucl_allowed = db.execute(text("SELECT client_id FROM user_clients WHERE user_id = :uid"), {"uid": current_user_id}).fetchall()
            allowed_client_ids = [r[0] for r in ucl_allowed]

    # جلب المستخدمين مع تقييدهم بصلاحيات المستخدم الحالي مباشرة من قاعدة البيانات أو الفلترة اللاحقة
    user_query = db.query(User)
    if hasattr(User, 'role'):
        user_query = user_query.filter(
            or_(
                User.role.ilike("%operation%"),
                User.role.ilike("%movement%"),
                User.role.ilike("%supervisor%")
            )
        )

    users = user_query.all()
    if not users:
        users = db.query(User).all()

    users_performance = []

    for user in users:
        u_name = getattr(user, 'username', None) or str(user)
        u_email = getattr(user, 'email', '')
        u_id = getattr(user, 'id', None)

        user_comp_ids = []
        user_client_ids_list = []
        u_company_names_list = []
        u_client_names_list = []

        if u_id:
            if "user_companies" in table_names:
                uc_rows = db.execute(text("SELECT company_id FROM user_companies WHERE user_id = :uid"), {"uid": u_id}).fetchall()
                user_comp_ids = [r[0] for r in uc_rows]
                if user_comp_ids and "companies" in table_names:
                    placeholders = ", ".join([f":c_{i}" for i in range(len(user_comp_ids))])
                    params = {f"c_{i}": cid for i, cid in enumerate(user_comp_ids)}
                    comps = db.execute(text(f"SELECT id, name FROM companies WHERE id IN ({placeholders})"), params).fetchall()
                    u_company_names_list = [c[1] for c in comps]

            if "user_clients" in table_names:
                ucl_rows = db.execute(text("SELECT client_id FROM user_clients WHERE user_id = :uid"), {"uid": u_id}).fetchall()
                user_client_ids_list = [r[0] for r in ucl_rows]
                if user_client_ids_list and "clients" in table_names:
                    placeholders_cl = ", ".join([f":cl_{i}" for i in range(len(user_client_ids_list))])
                    params_cl = {f"cl_{i}": clid for i, clid in enumerate(user_client_ids_list)}
                    cls = db.execute(text(f"SELECT id, name FROM clients WHERE id IN ({placeholders_cl})"), params_cl).fetchall()
                    u_client_names_list = [cl[1] for cl in cls]

        # 1. استبعاد المستخدم إذا كانت شركاته لا تقع ضمن شركات المستخدم الحالي (لو وُجدت قيود)
        if allowed_company_ids and not any(cid in allowed_company_ids for cid in user_comp_ids):
            continue

        # 2. استبعاد المستخدم إذا لم يكن مشتركاً مع المستخدم الحالي في أي من العملاء (Clients) المسموح له بهم
        if allowed_client_ids and not any(clid in allowed_client_ids for clid in user_client_ids_list):
            continue

        # 3. الفلترة بالاختيارات المحددة في الفلتر العلوي للصفحة
        if comp_id and comp_id not in user_comp_ids:
            continue
        if client_comp_id and client_comp_id not in user_client_ids_list:
            continue

        u_company_name = ", ".join(u_company_names_list) if u_company_names_list else "لا يوجد"
        u_client_name = ", ".join(u_client_names_list) if u_client_names_list else "لا يوجد"

        journey_query = db.query(Journey)
        j_conditions = []
        if hasattr(Journey, 'created_by'):
            if u_email:
                j_conditions.append(Journey.created_by.ilike(u_email))
            if u_name:
                j_conditions.append(Journey.created_by.ilike(u_name))
            if u_id is not None:
                j_conditions.append(Journey.created_by == str(u_id))
        
        if j_conditions:
            journey_query = journey_query.filter(or_(*j_conditions))
            
        all_user_journeys = journey_query.all()

        followup_query = db.query(JourneyFollowUp)
        if hasattr(JourneyFollowUp, 'is_deleted'):
            followup_query = followup_query.filter(JourneyFollowUp.is_deleted == False)

        f_conditions = []
        if hasattr(JourneyFollowUp, 'employee_name'):
            if u_email:
                f_conditions.append(JourneyFollowUp.employee_name.ilike(u_email))
            if u_name:
                f_conditions.append(JourneyFollowUp.employee_name.ilike(u_name))
            if u_id is not None:
                f_conditions.append(JourneyFollowUp.employee_name == str(u_id))
        
        if f_conditions:
            followup_query = followup_query.filter(or_(*f_conditions))

        all_user_followups = followup_query.all()

        all_activity_dates = []
        for j in all_user_journeys:
            dt_val = getattr(j, 'created_at', None)
            if isinstance(dt_val, datetime):
                all_activity_dates.append(dt_val)
        for f in all_user_followups:
            dt_val = getattr(f, 'created_at', None)
            if isinstance(dt_val, datetime):
                all_activity_dates.append(dt_val)

        all_activity_dates = sorted(list(set(all_activity_dates)))

        first_activity_date = all_activity_dates[0] if all_activity_dates else None
        last_activity_date = all_activity_dates[-1] if all_activity_dates else None

        filtered_journeys = [j for j in all_user_journeys if getattr(j, 'created_at', None) and start_dt <= j.created_at <= end_dt]
        handled_journeys_count = len(filtered_journeys)

        filtered_followups = [f for f in all_user_followups if getattr(f, 'created_at', None) and start_dt <= f.created_at <= end_dt]
        user_followups = sorted(filtered_followups, key=lambda x: x.created_at)
        followups_count = len(user_followups)

        avg_diff_str = "غير متوفر"
        if followups_count > 1:
            total_seconds = 0
            intervals_count = 0
            for i in range(1, followups_count):
                prev_time = user_followups[i - 1].created_at
                curr_time = user_followups[i].created_at
                if isinstance(prev_time, datetime) and isinstance(curr_time, datetime):
                    diff = (curr_time - prev_time).total_seconds()
                    if diff > 0:
                        total_seconds += diff
                        intervals_count += 1
            
            if intervals_count > 0:
                avg_seconds = total_seconds / intervals_count
                avg_minutes = int(avg_seconds // 60)
                avg_hours = int(avg_minutes // 60)
                if avg_hours > 0:
                    rem_mins = avg_minutes % 60
                    avg_diff_str = f"كل {avg_hours}س و {rem_mins}د"
                else:
                    avg_diff_str = f"كل {max(1, avg_minutes)} دقيقة"
            else:
                avg_diff_str = "منتظم"
        elif followups_count == 1:
            avg_diff_str = "متابعة واحدة"
        else:
            avg_diff_str = "لا توجد متابعات"

        score = round((handled_journeys_count * 2.0) + (followups_count * 1.5), 1)

        users_performance.append({
            "name": u_name,
            "email": u_email,
            "company_id": u_company_name,
            "client_company_id": u_client_name,
            "handled_journeys": handled_journeys_count,
            "follow_ups": followups_count,
            "avg_response": avg_diff_str,
            "first_journey": first_activity_date.strftime("%Y-%m-%d | %I:%M %p") if first_activity_date else "لا يوجد",
            "last_journey": last_activity_date.strftime("%Y-%m-%d | %I:%M %p") if last_activity_date else "لا يوجد",
            "score": score
        })

    sorted_users = sorted(users_performance, key=lambda x: (str(x["company_id"]), str(x["client_company_id"]), x["score"]), reverse=True)
    
    for index, user_data in enumerate(sorted_users, start=1):
        user_data["rank"] = index

    final_users = sorted_users
    if employee_name and employee_name.strip():
        search_term = employee_name.strip().lower()
        final_users = [
            u for u in sorted_users 
            if search_term in u["name"].lower() or search_term in u["email"].lower()
        ]

    companies = []
    clients = []

    if allowed_company_ids and "companies" in table_names:
        placeholders_comp = ", ".join([f":ac_{i}" for i in range(len(allowed_company_ids))])
        params_comp = {f"ac_{i}": acid for i, acid in enumerate(allowed_company_ids)}
        companies = db.execute(text(f"SELECT id, name FROM companies WHERE id IN ({placeholders_comp})"), params_comp).fetchall()
    elif not allowed_company_ids and "companies" in table_names:
        companies = db.execute(text("SELECT id, name FROM companies")).fetchall()

    if allowed_client_ids and "clients" in table_names:
        placeholders_cl_all = ", ".join([f":acl_{i}" for i in range(len(allowed_client_ids))])
        params_cl_all = {f"acl_{i}": aclid for i, aclid in enumerate(allowed_client_ids)}
        
        if comp_id:
            clients = db.execute(
                text(f"SELECT cl.id, cl.name FROM clients cl JOIN user_clients ucl ON cl.id = ucl.client_id WHERE ucl.user_id = :uid AND cl.company_id = :cid AND cl.id IN ({placeholders_cl_all})"),
                {"uid": current_user_id, "cid": comp_id, **params_cl_all}
            ).fetchall()
        else:
            clients = db.execute(
                text(f"SELECT id, name FROM clients WHERE id IN ({placeholders_cl_all})"),
                params_cl_all
            ).fetchall()
    else:
        if "clients" in table_names:
            if comp_id:
                clients = db.execute(
                    text("SELECT id, name FROM clients WHERE company_id = :cid"), 
                    {"cid": comp_id}
                ).fetchall()
            else:
                clients = db.execute(text("SELECT id, name FROM clients")).fetchall()

    return templates.TemplateResponse(
        request=request,
        name="employee_league.html",
        context={
            "users_performance": final_users,
            "from_date": from_date,
            "to_date": to_date,
            "employee_name": employee_name or "",
            "company_id": company_id or "",
            "client_company_id": client_company_id or "",
            "companies": companies,
            "clients": clients
        }
    )

# ----------------------------------------------------
# COMPANIES, SEGMENTS, ROUTES, LOAD TYPES & SETTINGS
# ----------------------------------------------------
# ----------------------------------------------------
# COMPANIES, CLIENTS & SEGMENTS MANAGEMENT
# ----------------------------------------------------

@app.get("/companies")
def list_companies(request: Request, db: Session = Depends(get_db)):
    logged_in_user = get_current_user(request, db) if "get_current_user" in globals() else None
    
    transport_companies = db.query(Company).all()
    client_companies = db.query(Client).all()
    
    return templates.TemplateResponse(
        request=request, 
        name="companies.html", 
        context={
            "transport_companies": transport_companies,
            "client_companies": client_companies,
            "current_user": logged_in_user
        }
    )

# --- 1. إضافة شركة نقل رئيسية (Transport Company) ---
@app.post("/companies/add")
def add_company(
    name: str = Form(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    new_company = Company(name=name)
    db.add(new_company)
    db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/companies/update/{id}")
def update_company(
    id: int,
    name: str = Form(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    company = db.query(Company).filter(Company.id == id).first()
    if company:
        company.name = name
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/companies/delete/{id}")
def delete_company(
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(require_admin_or_developer)
):
    company = db.query(Company).filter(Company.id == id).first()
    if company:
        db.delete(company)
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)


# --- 2. إضافة شركة عميل (Client) تابعة لشركة نقل ---
@app.post("/clients/add")
def add_client(
    company_id: int = Form(...),
    name: str = Form(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    new_client = Client(name=name, company_id=company_id)
    db.add(new_client)
    db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/clients/update/{id}")
def update_client(
    id: int,
    company_id: int = Form(...),
    name: str = Form(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    client = db.query(Client).filter(Client.id == id).first()
    if client:
        client.name = name
        client.company_id = company_id
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/clients/delete/{id}")
def delete_client(
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(require_admin_or_developer)
):
    client = db.query(Client).filter(Client.id == id).first()
    if client:
        db.delete(client)
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)


# --- 3. إضافة قطاع/قسم (Segment) يتبع للعميل ---
@app.post("/segments/add")
def add_segment(
    client_id: int = Form(...), 
    name: str = Form(...), 
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    new_segment = Segment(client_id=client_id, name=name)
    db.add(new_segment)
    db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/segments/update/{id}")
def update_segment(
    id: int, 
    name: str = Form(...), 
    client_id: int = Form(...), 
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer)
):
    segment = db.query(Segment).filter(Segment.id == id).first()
    if segment:
        segment.name = name
        segment.client_id = client_id
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)

@app.post("/segments/delete/{id}")
def delete_segment(
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(require_admin_or_developer)
):
    segment = db.query(Segment).filter(Segment.id == id).first()
    if segment:
        db.delete(segment)
        db.commit()
    return RedirectResponse(url="/companies", status_code=303)

# --- ROUTES & LOCATIONS ---

# 1. صفحة إدارة الأماكن (مع تقييد الشركات والعملاء حصرياً حسب صلاحيات الـ User وإلغاء الـ General)
@app.get("/locations")
def list_locations(
    request: Request,
    company_id: Optional[str] = None,
    client_id: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    if not current_user:
        return RedirectResponse(url="/login", status_code=303)
        
    try:
        company_id = int(company_id) if company_id and str(company_id).isdigit() else None
    except ValueError:
        company_id = None

    try:
        client_id = int(client_id) if client_id and str(client_id).isdigit() else None
    except ValueError:
        client_id = None

    # استخراج الشركات والعملاء المربوطين بالمستخدم الحالي حصرياً لكل المستخدمين
    companies = current_user.companies if hasattr(current_user, 'companies') and current_user.companies else []
    clients = current_user.clients if hasattr(current_user, 'clients') and current_user.clients else []

    allowed_company_ids = [c.id for c in companies]
    allowed_client_ids = [cl.id for cl in clients]

    # إذا لم تكن هناك شركات أو عملاء مصرح بها للمستخدم، تكون القوائم فارغة تماماً
    if not allowed_company_ids or not allowed_client_ids:
        companies = []
        clients = []
        company_id = None
        client_id = None
        locations = []
        locations_json = json.dumps([])
    else:
        if company_id not in allowed_company_ids:
            company_id = companies[0].id if companies else None

        if client_id not in allowed_client_ids:
            client_id = clients[0].id if clients else None

        # جلب الأماكن مع تطبيق قيود الشركات والعملاء الخاصة بالمستخدم حصرياً
        query = db.query(Location)
        query = query.filter(Location.company_id.in_(allowed_company_ids))
        query = query.filter(Location.client_id.in_(allowed_client_ids))

        if company_id:
            query = query.filter(Location.company_id == company_id)
        if client_id:
            query = query.filter(Location.client_id == client_id)
            
        locations = query.all()

        locations_list = [
            {
                "id": loc.id,
                "name": loc.name,
                "location_type": getattr(loc, 'location_type', 'main'),
                "lat": float(loc.latitude) if loc.latitude is not None else None,
                "lng": float(loc.longitude) if loc.longitude is not None else None
            } for loc in locations
        ]
        locations_json = json.dumps(locations_list)

    return templates.TemplateResponse(
        request=request,
        name="locations.html",
        context={
            "locations": locations,
            "companies": companies,
            "clients": clients,
            "locations_json": locations_json,
            "selected_company_id": company_id,
            "selected_client_id": client_id
        }
    )

# 2. إضافة مكان جديد من الخريطة
@app.post("/api/locations/add")
def add_location_api(
    name: str = Form(...),
    location_type: Optional[str] = Form('main'),
    latitude: float = Form(...),
    longitude: float = Form(...),
    company_id: int = Form(...),
    client_id: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []
    
    c_id = int(company_id) if company_id else None
    cl_id = int(client_id) if client_id and str(client_id).isdigit() else None

    if c_id not in allowed_company_ids or (cl_id and cl_id not in allowed_client_ids):
        raise HTTPException(status_code=403, detail="غير مسموح لك بإضافة مكان لهذه الشركة أو العميل")

    final_type = location_type or 'main'

    new_location = Location(
        name=name,
        location_type=final_type,
        latitude=latitude,
        longitude=longitude,
        company_id=c_id,
        client_id=cl_id
    )

    db.add(new_location)
    db.commit()
    
    redirect_url = f"/locations?company_id={c_id}"
    if cl_id:
        redirect_url += f"&client_id={cl_id}"
        
    return RedirectResponse(url=redirect_url, status_code=303)

# 3. حذف مكان من صفحة الأماكن
@app.post("/locations/delete/{id}")
def delete_location(
    id: int,
    company_id: Optional[str] = Form(None),
    client_id: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    location = db.query(Location).filter(Location.id == id).first()
    if location:
        allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []
        
        if location.company_id in allowed_company_ids and (not location.client_id or location.client_id in allowed_client_ids):
            db.delete(location)
            db.commit()
        else:
            raise HTTPException(status_code=403, detail="غير مسموح لك بحذف هذا المكان")

    redirect_url = "/locations"
    params = []
    if company_id and str(company_id).isdigit():
        params.append(f"company_id={company_id}")
    if client_id and str(client_id).isdigit():
        params.append(f"client_id={client_id}")
        
    if params:
        redirect_url += "?" + "&".join(params)

    return RedirectResponse(url=redirect_url, status_code=303)

# 4. تعديل مكان
@app.post("/locations/update/{location_id}")
def update_location(
    location_id: int,
    name: str = Form(...),
    client_id: Optional[str] = Form(None),
    location_type: str = Form(...),
    latitude: float = Form(...),
    longitude: float = Form(...),
    company_id: Optional[str] = Form(None),
    current_client_id_filter: Optional[str] = Form(None, alias="client_id_filter"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    location = db.query(Location).filter(Location.id == location_id).first()
    if not location:
        raise HTTPException(status_code=404, detail="المكان غير موجود")

    allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

    if location.company_id not in allowed_company_ids:
        raise HTTPException(status_code=403, detail="غير مسموح لك بتعديل هذا المكان")

    c_id = int(company_id) if company_id and str(company_id).isdigit() else None
    cl_id = int(client_id) if client_id and str(client_id).isdigit() else None

    if c_id and c_id not in allowed_company_ids:
        raise HTTPException(status_code=403, detail="غير مسموح نقل المكان لهذه الشركة")
    if cl_id and cl_id not in allowed_client_ids:
        raise HTTPException(status_code=403, detail="غير مسموح نقل المكان لهذا العميل")

    location.name = name
    location.client_id = cl_id
    location.location_type = location_type
    location.latitude = latitude
    location.longitude = longitude
    
    db.commit()

    redirect_url = "/locations"
    params = []
    
    comp_val = c_id if c_id else location.company_id
    if comp_val:
        params.append(f"company_id={comp_val}")
        
    if cl_id:
        params.append(f"client_id={cl_id}")
        
    if params:
        redirect_url += "?" + "&".join(params)

    return RedirectResponse(url=redirect_url, status_code=303)


# 5. عرض خطوط السير (Routes) مع ربط بيانات المواقع والعملاء باليوزر حصرياً
@app.get("/routes")
def list_routes(
    request: Request, 
    company_id: Optional[str] = None, 
    client_company_id: Optional[str] = None, 
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    if not current_user:
        return RedirectResponse(url="/login", status_code=303)
    
    companies = current_user.companies if hasattr(current_user, 'companies') and current_user.companies else []
    clients = current_user.clients if hasattr(current_user, 'clients') and current_user.clients else []

    allowed_company_ids = [c.id for c in companies]
    allowed_client_ids = [cl.id for cl in clients]

    try:
        company_id = int(company_id) if company_id and str(company_id).isdigit() else None
    except ValueError:
        company_id = None

    try:
        client_company_id = int(client_company_id) if client_company_id and str(client_company_id).isdigit() else None
    except ValueError:
        client_company_id = None

    # دالة مساعدة لتحويل الأرقام العشرية إلى صيغة hh:mm
    def format_to_hhmm(hours_float):
        if not hours_float or hours_float < 0:
            return "00:00"
        h = int(hours_float)
        m = int(round((hours_float - h) * 60))
        if m == 60:
            h += 1
            m = 0
        return f"{h:02d}:{m:02d}"

    if not allowed_company_ids or not allowed_client_ids:
        companies = []
        clients = []
        company_id = None
        client_company_id = None
        routes = []
        available_locations = []
        locations_json = json.dumps([])
        routes_data = []
        speed = 40.0
    else:
        if company_id not in allowed_company_ids:
            company_id = companies[0].id if companies else None
        if client_company_id not in allowed_client_ids:
            client_company_id = clients[0].id if clients else None

        query = db.query(Route)
        query = query.filter(Route.company_id.in_(allowed_company_ids))
        query = query.filter(Route.client_company_id.in_(allowed_client_ids))

        if company_id:
            query = query.filter(Route.company_id == company_id)
        if client_company_id:
            query = query.filter(Route.client_company_id == client_company_id)

        routes = query.all()
        
        loc_query = db.query(Location)
        loc_query = loc_query.filter(Location.company_id.in_(allowed_company_ids))
        loc_query = loc_query.filter(Location.client_id.in_(allowed_client_ids))

        if company_id:
            loc_query = loc_query.filter(Location.company_id == company_id)
        if client_company_id:
            loc_query = loc_query.filter(Location.client_id == client_company_id)
        available_locations = loc_query.all()

        locations_list = [
            {
                "id": loc.id,
                "name": loc.name,
                "location_type": getattr(loc, 'location_type', 'main'),
                "lat": float(loc.latitude) if loc.latitude is not None else None,
                "lng": float(loc.longitude) if loc.longitude is not None else None
            } for loc in available_locations
        ]
        locations_json = json.dumps(locations_list)

        setting = None
        if company_id and client_company_id:
            setting = db.query(Setting).filter(
                Setting.company_id == company_id,
                Setting.client_company_id == client_company_id
            ).first()
            
        if not setting and company_id:
            setting = db.query(Setting).filter(
                Setting.company_id == company_id,
                Setting.client_company_id == None
            ).first()
            
        if not setting:
            setting = db.query(Setting).filter(
                Setting.company_id == None,
                Setting.client_company_id == None
            ).first()
            
        speed = float(setting.default_speed) if setting and setting.default_speed else 40.0
        rest_duration = float(setting.rest_time) if setting and setting.rest_time else 0.5
        
        routes_data = []
        for r in routes:
            # الاعتماد على الحقول الموجودة في المسار أو حسابها بناء على السرعة
            d_hours = getattr(r, 'driving_hours', None)
            if d_hours is None:
                d_hours = (r.distance_km / speed) if (r.distance_km and speed > 0) else 0.0
            
            r_time = getattr(r, 'rest_time', None)
            if r_time is None:
                rest_stops = int(d_hours // 2) if d_hours >= 2 else 0
                r_time = rest_stops * rest_duration
                
            t_time = getattr(r, 'total_time', None)
            if t_time is None:
                t_time = d_hours + r_time

            routes_data.append({
                "route": r, 
                "driving_hours": format_to_hhmm(d_hours),
                "total_rest_time": format_to_hhmm(r_time),
                "calculated_hours": format_to_hhmm(t_time)
            })
        
    return templates.TemplateResponse(
        request=request, 
        name="routes.html", 
        context={
            "routes_data": routes_data, 
            "companies": companies,
            "clients": clients,
            "locations": available_locations,
            "locations_json": locations_json,
            "selected_company_id": company_id,
            "selected_client_id": client_company_id,
            "speed": speed
        }
    )

# 6. إضافة خط السير الجديد
@app.post("/routes/add")
def add_route(
    origin_id: int = Form(...), 
    destination_id: int = Form(...), 
    distance_km: float = Form(...), 
    route_type_name: Optional[str] = Form(None),
    company_id: int = Form(...),
    client_company_id: int = Form(...), 
    via_point_ids: Optional[list[int]] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []

    c_id = int(company_id)
    cc_id = int(client_company_id) if client_company_id and str(client_company_id).isdigit() else None

    if c_id not in allowed_company_ids or (cc_id and cc_id not in allowed_client_ids):
        raise HTTPException(status_code=403, detail="غير مسموح لك بإضافة خط سير لهذه الشركة أو العميل")

    last = db.query(Route.id).order_by(Route.id.desc()).first()
    new_route = Route(
        code=f"RT-{(last[0] + 1) if last else 1:04d}", 
        route_type_name=route_type_name,
        origin_id=origin_id, 
        destination_id=destination_id, 
        distance_km=distance_km,
        company_id=c_id,
        client_company_id=cc_id
    )
    
    db.add(new_route)
    db.flush()

    if via_point_ids:
        for index, v_id in enumerate(via_point_ids):
            via_point_record = RouteViaPoint(
                route_id=new_route.id,
                location_id=v_id,
                sequence_order=index
            )
            db.add(via_point_record)

    db.commit()
    
    redirect_url = f"/routes?company_id={c_id}"
    if cc_id:
        redirect_url += f"&client_company_id={cc_id}"
        
    return RedirectResponse(url=redirect_url, status_code=303)

# 7. حذف خط السير
@app.post("/routes/delete/{id}")
def delete_route(
    id: int, 
    company_id: Optional[str] = Form(None),
    client_company_id: Optional[str] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    route = db.query(Route).filter(Route.id == id).first()
    if route:
        allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
        allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []
        
        if route.company_id in allowed_company_ids and (not route.client_company_id or route.client_company_id in allowed_client_ids):
            db.delete(route)
            db.commit()
        else:
            raise HTTPException(status_code=403, detail="غير مسموح لك بحذف خط السير هذا")
        
    redirect_url = f"/routes"
    params = []
    try:
        c_id = int(company_id) if company_id and str(company_id).isdigit() else None
        if c_id:
            params.append(f"company_id={c_id}")
    except ValueError:
        pass

    try:
        cc_id = int(client_company_id) if client_company_id and str(client_company_id).isdigit() else None
        if cc_id:
            params.append(f"client_company_id={cc_id}")
    except ValueError:
        pass

    if params:
        redirect_url += "?" + "&".join(params)
        
    return RedirectResponse(url=redirect_url, status_code=303)

# 8. تعديل خط السير
@app.post("/routes/update/{id}")
def update_route(
    id: int,
    origin_id: int = Form(...), 
    destination_id: int = Form(...), 
    distance_km: float = Form(...),
    route_type_name: Optional[str] = Form(None),
    company_id: Optional[str] = Form(None),
    client_company_id: Optional[str] = Form(None),
    via_point_ids: Optional[list[int]] = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    route = db.query(Route).filter(Route.id == id).first()
    if not route:
        raise HTTPException(status_code=404, detail="Route not found")
        
    allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    allowed_client_ids = [cl.id for cl in current_user.clients] if hasattr(current_user, 'clients') and current_user.clients else []
    
    if route.company_id not in allowed_company_ids:
        raise HTTPException(status_code=403, detail="غير مسموح لك بتعديل خط السير هذا")

    c_id = int(company_id) if company_id and str(company_id).isdigit() else None
    cc_id = int(client_company_id) if client_company_id and str(client_company_id).isdigit() else None

    if c_id and c_id not in allowed_company_ids:
        raise HTTPException(status_code=403, detail="غير مسموح نقل خط السير لهذه الشركة")
    if cc_id and cc_id not in allowed_client_ids:
        raise HTTPException(status_code=403, detail="غير مسموح نقل خط السير لهذا العميل")

    route.origin_id = origin_id
    route.destination_id = destination_id
    route.distance_km = distance_km
    if c_id:
        route.company_id = c_id
    if cc_id:
        route.client_company_id = cc_id
    if route_type_name is not None:
        route.route_type_name = route_type_name
        
    db.query(RouteViaPoint).filter(RouteViaPoint.route_id == route.id).delete()
    
    if via_point_ids:
        for index, v_id in enumerate(via_point_ids):
            via_point_record = RouteViaPoint(
                route_id=route.id,
                location_id=v_id,
                sequence_order=index
            )
            db.add(via_point_record)

    db.commit()

    redirect_url = f"/routes"
    params = []
    if c_id:
        params.append(f"company_id={c_id}")
    if cc_id:
        params.append(f"client_company_id={cc_id}")

    if params:
        redirect_url += "?" + "&".join(params)
        
    return RedirectResponse(url=redirect_url, status_code=303)

# 9. جلب تفاصيل خط السير والنقاط الوسيطة
@app.get("/api/routes/{route_id}/details")
def get_route_details(route_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if not current_user:
        return {"error": "Unauthorized"}
        
    route = db.query(Route).filter(Route.id == route_id).first()
    if not route:
        return {"error": "Route not found"}
        
    allowed_company_ids = [c.id for c in current_user.companies] if hasattr(current_user, 'companies') and current_user.companies else []
    
    if route.company_id not in allowed_company_ids:
        return {"error": "غير مسموح لك بمعاينة خط السير هذا"}
    
    via_points_list = []
    if hasattr(route, 'via_points') and route.via_points:
        for vp in route.via_points:
            loc = vp.location if hasattr(vp, 'location') else None
            if loc:
                via_points_list.append({
                    "id": loc.id,
                    "name": loc.name,
                    "location_type": getattr(loc, 'location_type', 'via'),
                    "lat": float(loc.latitude) if loc.latitude is not None else None,
                    "lng": float(loc.longitude) if loc.longitude is not None else None
                })

    return {
        "id": route.id,
        "route_type_name": route.route_type_name,
        "distance_km": route.distance_km,
        "origin": {
            "id": route.origin_location.id if route.origin_location else None,
            "name": route.origin_location.name if route.origin_location else "",
            "location_type": getattr(route.origin_location, 'location_type', 'main') if route.origin_location else 'main',
            "lat": float(route.origin_location.latitude) if route.origin_location and route.origin_location.latitude is not None else None,
            "lng": float(route.origin_location.longitude) if route.origin_location and route.origin_location.longitude is not None else None,
        },
        "destination": {
            "id": route.destination_location.id if route.destination_location else None,
            "name": route.destination_location.name if route.destination_location else "",
            "location_type": getattr(route.destination_location, 'location_type', 'main') if route.destination_location else 'main',
            "lat": float(route.destination_location.latitude) if route.destination_location and route.destination_location.latitude is not None else None,
            "lng": float(route.destination_location.longitude) if route.destination_location and route.destination_location.longitude is not None else None,
        },
        "via_points": via_points_list
    }

# --- SETTINGS ---
# --- دالة مساعدة لتحويل HH:MM إلى ساعات عشرية للحسابات ---
def parse_hh_mm_to_float(val_str: str) -> float:
    if not val_str or ":" not in val_str:
        try:
            return float(val_str or 0.0)
        except ValueError:
            return 0.0
    try:
        h, m = val_str.split(":")
        return float(h) + float(m) / 60.0
    except ValueError:
        return 0.0

# --- دالة مساعدة لتحويل الساعات العشرية إلى صيغة HH:MM للواجهة ---
def float_to_hh_mm(val_float: float) -> str:
    if val_float is None:
        return "00:00"
    hours = int(val_float)
    minutes = round((val_float - hours) * 60)
    if minutes == 60:
        hours += 1
        minutes = 0
    return f"{hours:02d}:{minutes:02d}"

# --- دالة مساعدة للتحقق من صلاحيات الشركات للمستخدم الحالي ---
def get_user_allowed_companies(current_user: User, db: Session):
    # إذا كان المشرف أو المطور، يمكنه رؤية كل الشركات
    if getattr(current_user, "is_admin", False) or getattr(current_user, "is_developer", False) or getattr(current_user, "role", "") in ["admin", "developer"]:
        return db.query(Company).all(), db.query(Client).all()
    
    # وإلا يتم جلب الشركات والعملاء المربوطين بالمستخدم فقط (حسب حقول نظامك)
    allowed_company_ids = [c.id for c in getattr(current_user, "companies", [])]
    # دعم لو كان المستخدم مربوط بشركة واحدة مفردة مثلاً: current_user.company_id
    if hasattr(current_user, "company_id") and current_user.company_id:
        allowed_company_ids.append(current_user.company_id)
        
    companies = db.query(Company).filter(Company.id.in_(allowed_company_ids)).all() if allowed_company_ids else []
    
    allowed_client_ids = [cl.id for cl in getattr(current_user, "clients", [])]
    clients = db.query(Client).filter(Client.id.in_(allowed_client_ids)).all() if allowed_client_ids else []
    
    return companies, clients

# --- SETTINGS: عرض صفحة الإعدادات ---
@app.get("/settings", response_class=HTMLResponse)
def get_settings(request: Request, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    # التحقق مما إذا كان المستخدم مطوراً أو أدمن
    is_admin_or_developer = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    
    # إذا لم يكن أدمن أو مطور، يتم توجيهه لصفحة الخطأ 403
    if not is_admin_or_developer:
        return templates.TemplateResponse(
            request=request, 
            name="403.html", 
            status_code=403
        )

    # جلب الشركات والعملاء المصرح لهم فقط للمستخدم الحالي
    companies, clients = get_user_allowed_companies(current_user, db)
    
    return templates.TemplateResponse(
        request=request, 
        name="settings.html", 
        context={
            "companies": companies,
            "clients": clients,
            "current_user": current_user
        }
    )

# --- API: جلب الإعدادات المخصصة لشركة أو عميل معين بصيغة HH:MM ---
@app.get("/api/settings")
def get_api_settings(
    company_id: int = None, 
    client_company_id: int = None, 
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    # تنقية المتغيرات لضمان التعامل السليم مع القيم الفارغة
    c_id = company_id if company_id and company_id != 0 else None
    client_id = client_company_id if client_company_id and client_company_id != 0 else None

    # التحقق من صلاحية المستخدم على الشركة المطلوبة إن لم يكن مشرفاً عاماً
    if not (getattr(current_user, "is_admin", False) or getattr(current_user, "is_developer", False) or getattr(current_user, "role", "") in ["admin", "developer"]):
        companies, _ = get_user_allowed_companies(current_user, db)
        allowed_c_ids = [c.id for c in companies]
        if c_id and c_id not in allowed_c_ids:
            return {"error": "Unauthorized company access"}, 403

    # البحث عن الإعدادات المخصصة بناءً على الشركة والعميل
    query = db.query(Setting)
    
    if c_id:
        query = query.filter(Setting.company_id == c_id)
    else:
        query = query.filter(Setting.company_id == None)
        
    if client_id:
        query = query.filter(Setting.client_company_id == client_id)
    else:
        query = query.filter(Setting.client_company_id == None)
        
    setting = query.first()
    
    # إذا لم تكن موجودة خاصة بالشركة، نبحث عن الإعدادات العامة (الافتراضية) للنظام
    if not setting:
        setting = db.query(Setting).filter(
            Setting.company_id == None,
            Setting.client_company_id == None
        ).first()

    if not setting:
        # إرجاع قيم افتراضية آمنة في حال عدم وجود أي سجل
        return {
            "default_speed": 40.0,
            "rest_time": "00:30",
            "sleep_time": "08:00",
            "max_driving_hours": "10:00",
            "follow_up_interval": "02:00"
        }
        
    return {
        "default_speed": setting.default_speed,
        "rest_time": float_to_hh_mm(setting.rest_time),
        "sleep_time": float_to_hh_mm(setting.sleep_time),
        "max_driving_hours": float_to_hh_mm(setting.max_driving_hours),
        "follow_up_interval": float_to_hh_mm(setting.follow_up_interval)
    }

# --- SAVE SETTINGS: حفظ أو تحديث الإعدادات المخصصة لكل شركة وعميل ---
@app.post("/settings/save")
def save_settings(
    company_id: int = Form(None),
    client_company_id: int = Form(None),
    default_speed: float = Form(...),
    rest_time: str = Form(...),
    sleep_time: str = Form(...),
    max_driving_hours: str = Form(...),
    follow_up_interval: str = Form(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    # التحقق من صلاحيات التعديل
    if not (getattr(current_user, "is_admin", False) or getattr(current_user, "is_developer", False) or getattr(current_user, "role", "") in ["admin", "developer"]):
        companies, _ = get_user_allowed_companies(current_user, db)
        allowed_c_ids = [c.id for c in companies]
        if company_id and company_id != 0 and company_id not in allowed_c_ids:
            return {"status": "error", "message": "Unauthorized action for this company"}

    # تحويل الأوقات الواردة بصيغة hh:mm إلى قيم عشرية لتخزينها بدقة في قاعدة البيانات للحسابات
    rest_float = parse_hh_mm_to_float(rest_time)
    sleep_float = parse_hh_mm_to_float(sleep_time)
    driving_float = parse_hh_mm_to_float(max_driving_hours)
    follow_float = parse_hh_mm_to_float(follow_up_interval)
    
    # تنقية المتغيرات للبحث بدقة
    c_id = company_id if company_id and company_id != 0 else None
    client_id = client_company_id if client_company_id and client_company_id != 0 else None

    # البحث عما إذا كان السجل موجوداً مسبقاً لنفس الشركة والعميل
    query = db.query(Setting)
    
    if c_id:
        query = query.filter(Setting.company_id == c_id)
    else:
        query = query.filter(Setting.company_id == None)
        
    if client_id:
        query = query.filter(Setting.client_company_id == client_id)
    else:
        query = query.filter(Setting.client_company_id == None)
        
    setting = query.first()
    
    if setting:
        # تحديث القيم الحالية للسجل الموجود لهذه الشركة فقط دون المساس بالبقية
        setting.default_speed = default_speed
        setting.rest_time = rest_float
        setting.sleep_time = sleep_float
        setting.max_driving_hours = driving_float
        setting.follow_up_interval = follow_float
    else:
        # إنشاء سجل جديد مخصص بالكامل لهذه الشركة والعميل
        setting = Setting(
            company_id=c_id,
            client_company_id=client_id,
            default_speed=default_speed,
            rest_time=rest_float,
            sleep_time=sleep_float,
            max_driving_hours=driving_float,
            follow_up_interval=follow_float
        )
        db.add(setting)
        
    db.commit()
    return {"status": "success", "message": "Settings updated successfully"}


# --- LOAD TYPES ---
@app.get("/load-types", response_class=HTMLResponse)
def list_load_types(
    request: Request, 
    q: Optional[str] = None, 
    company_id: Optional[int] = None,
    client_id: Optional[int] = None,
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    # التحقق من أن المستخدم Admin أو Developer فقط
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )

    if not is_admin_or_dev:
        return templates.TemplateResponse(
            request=request, 
            name="403.html", 
            status_code=403
        )

    companies = db.query(Company).all()
    clients = db.query(Client).all()
    query = db.query(LoadType)

    # تطبيق فلاتر البحث المخصصة إن وجدت
    if company_id:
        query = query.filter(LoadType.company_id == company_id)
    if client_id:
        query = query.filter(LoadType.client_id == client_id)
    if q:
        query = query.filter(LoadType.name.like(f"%{q}%"))
        
    load_types = query.all()
    
    return templates.TemplateResponse(
        request=request, 
        name="load_types.html", 
        context={
            "load_types": load_types, 
            "companies": companies,
            "clients": clients,
            "selected_company_id": company_id,
            "selected_client_id": client_id,
            "search_query": q or "",
            "current_user": current_user
        }
    )

@app.post("/load-types/add")
def add_load_type(
    request: Request,
    name: str = Form(...), 
    company_id: Optional[int] = Form(None), 
    client_id: Optional[int] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    c_id = company_id if company_id and company_id != 0 else None
    cl_id = client_id if client_id and client_id != 0 else None
    
    db.add(LoadType(name=name, company_id=c_id, client_id=cl_id))
    db.commit()
    return RedirectResponse(url="/load-types", status_code=303)

@app.post("/load-types/update/{id}")
def update_load_type(
    request: Request,
    id: int, 
    name: str = Form(...), 
    company_id: Optional[int] = Form(None), 
    client_id: Optional[int] = Form(None),
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        getattr(current_user, "is_superuser", False) or
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )
    if not is_admin_or_dev:
        return templates.TemplateResponse(request=request, name="403.html", status_code=403)

    c_id = company_id if company_id and company_id != 0 else None
    cl_id = client_id if client_id and client_id != 0 else None
    
    load_type = db.query(LoadType).filter(LoadType.id == id).first()
    if load_type:
        load_type.name = name
        load_type.company_id = c_id
        load_type.client_id = cl_id
        db.commit()
    return RedirectResponse(url="/load-types", status_code=303)

@app.post("/load-types/delete/{id}")
def delete_load_type(
    id: int, 
    db: Session = Depends(get_db), 
    current_user: User = Depends(get_current_user)
):
    load_type = db.query(LoadType).filter(LoadType.id == id).first()
    if load_type:
        db.delete(load_type)
        db.commit()
    return RedirectResponse(url="/load-types", status_code=303)


# ----------------------------------------------------
# USERS
# ----------------------------------------------------
@app.get("/users")
def list_users(
    request: Request, 
    q: Optional[str] = None, 
    filter_transport_id: Optional[str] = None, 
    filter_client_id: Optional[str] = None,    
    db: Session = Depends(get_db)
):
    logged_in_user = get_current_user(request, db) if "get_current_user" in globals() else None
    
    # تحويل القيم النصية القادمة من الـ Query إلى أرقام صحيحة بأمان
    t_id = int(filter_transport_id) if filter_transport_id and filter_transport_id.isdigit() else None
    c_id = int(filter_client_id) if filter_client_id and filter_client_id.isdigit() else None

    # استخدام joinedload لتحسين الأداء وجلب العلاقات مع المستخدم بquery واحد
    query = db.query(User).options(joinedload(User.companies), joinedload(User.clients))
    
    if q:
        query = query.filter(
            or_(User.username.like(f"%{q}%"), User.email.like(f"%{q}%"))
        )
    
    if t_id and t_id != 0:
        query = query.filter(User.companies.any(Company.id == t_id))
        
    if c_id and c_id != 0:
        query = query.filter(User.clients.any(Client.id == c_id))

    users = query.all()
    
    transport_companies = db.query(Company).all()
    client_companies = db.query(Client).all()
    companies = db.query(Company).all()
    
    current_user_data = {
        "username": logged_in_user.username if logged_in_user else "Admin", 
        "email": logged_in_user.email if logged_in_user and hasattr(logged_in_user, 'email') else "admin@system.com",
        "role": logged_in_user.role if logged_in_user and hasattr(logged_in_user, 'role') else "Admin"
    }
    
    # لو الطلب جاى عن طريق AJAX (لتحديث الجدول فقط بدون إعادة تحميل الصفحة بالكامل)
    if request.headers.get("x-requested-with") == "XMLHttpRequest":
        return templates.TemplateResponse(
            request=request, 
            name="users_table_partial.html", 
            context={"users": users}
        )

    return templates.TemplateResponse(
        request=request, 
        name="users.html", 
        context={
            "users": users, 
            "companies": companies,
            "transport_companies": transport_companies,
            "client_companies": client_companies,
            "search_query": q or "",
            "selected_transport_filter": str(t_id) if t_id else "",
            "selected_client_filter": str(c_id) if c_id else "",
            "current_user": current_user_data
        }
    )

@app.post("/users/add")
def add_user(
    username: Optional[str] = Form(None),
    full_name: Optional[str] = Form(None),
    email: Optional[str] = Form(None),
    password: str = Form(...),
    role: str = Form(...),
    permissions: Optional[str] = Form(None),
    company_ids: List[int] = Form(default=[]),  # استخدام default=[] لتجنب الـ 422
    client_ids: List[int] = Form(default=[]),   # استخدام default=[] لتجنب الـ 422
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin_or_developer),
):
  try:
    actual_name = full_name or username

    # إنشاء المستخدم الجديد
    new_user = User(
        username=actual_name,
        email=email,
        password=password,
        role=role,
        permissions=permissions,
    )

    # التعامل مع شركات النقل (Many-to-Many)
    if company_ids and hasattr(new_user, "companies"):
        selected_companies = db.query(Company).filter(Company.id.in_(company_ids)).all()
        new_user.companies = selected_companies

    # التعامل مع العملاء
    if client_ids and hasattr(new_user, "clients"):
        selected_clients = db.query(Client).filter(
            Client.id.in_(client_ids),
            Client.company_id.in_(company_ids) if company_ids else True
        ).all()
        new_user.clients = selected_clients

    db.add(new_user)
    db.commit()

    return RedirectResponse(url="/users", status_code=303)

  except Exception as e:
    db.rollback()
    import traceback
    print(traceback.format_exc())
    raise HTTPException(
        status_code=500, detail=f"Internal Server Error: {str(e)}"
    )

@app.post("/users/update/{user_id}")
async def update_user(
    user_id: int,
    username: str = Form(...),
    email: str = Form(...),
    password: Optional[str] = Form(None),
    role: str = Form(...),
    permissions: Optional[str] = Form(None),
    company_ids: List[int] = Form(default=[]),  # استخدام default=[] لتجنب الـ 422
    client_ids: List[int] = Form(default=[]),   # استخدام default=[] لتجنب الـ 422
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        user = db.query(User).filter(User.id == user_id).first()
        if not user:
            raise HTTPException(status_code=404, detail="User not found")
        
        user.username = username
        user.email = email
        if password and password.strip() != "":
            user.password = password  
        user.role = role
        if permissions is not None:
            user.permissions = permissions
        
        # تحديث العلاقة المتعددة لشركات النقل
        if company_ids and hasattr(user, "companies"):
            transport_comps = db.query(Company).filter(Company.id.in_(company_ids)).all()
            user.companies = transport_comps
        elif hasattr(user, "companies"):
            user.companies = []

        # تحديث العلاقة المتعددة للعملاء
        if client_ids and hasattr(user, "clients"):
            client_comps = db.query(Client).filter(
                Client.id.in_(client_ids),
                Client.company_id.in_(company_ids) if company_ids else True
            ).all()
            user.clients = client_comps
        elif hasattr(user, "clients"):
            user.clients = []
             
        db.commit()
        
        return RedirectResponse(url="/users", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print(traceback.format_exc())
        raise HTTPException(
            status_code=500, detail=f"Internal Server Error: {str(e)}"
        )

@app.post("/users/delete/{id}")
def delete_user(id: int, db: Session = Depends(get_db), current_user: User = Depends(require_admin_or_developer)):
    user = db.query(User).filter(User.id == id).first()
    if not user:
        raise HTTPException(status_code=404, detail="المستخدم غير موجود")

    db.delete(user)
    db.commit()
    return RedirectResponse(url="/users", status_code=303)

# ====================================================
# JOURNEYS & JOURNEY LIST
# ====================================================

@app.get("/journeys/new", response_class=HTMLResponse)
def new_journey_page(
    request: Request, 
    vehicle_id: Optional[int] = None,
    driver_id: Optional[int] = None,
    route_id: Optional[int] = None,
    load_type_id: Optional[int] = None,
    trailer_id: Optional[int] = None,
    segment_id: Optional[int] = None,
    client_company_id: Optional[int] = None,
    db: Session = Depends(get_db)
):
    try:
        logged_in_user = get_current_user(request, db)
        
        # 1. التحقق من صلاحيات المستخدم وهل هو Admin / Developer لعرض كل الشركات أو المقيدة له فقط
        is_admin_or_dev = (
            getattr(logged_in_user, "is_admin", False) or 
            getattr(logged_in_user, "is_developer", False) or 
            getattr(logged_in_user, "role", "") in ["admin", "developer"]
        )

        # 2. جلب الشركات/العملاء المرتبطين بالمستخدم حصرياً من جدول الصلاحيات الخاص به في users
        user_clients = []
        if is_admin_or_dev:
            user_clients = db.query(Client).all()
        else:
            if hasattr(logged_in_user, "clients") and logged_in_user.clients:
                user_clients = logged_in_user.clients
            elif hasattr(logged_in_user, "companies") and logged_in_user.companies:
                comp_ids = [c.id for c in logged_in_user.companies]
                user_clients = db.query(Client).filter(Client.company_id.in_(comp_ids)).all()
            else:
                user_clients = []

        clients = user_clients

        # 3. جعل التوابع الأولية فارغة تماماً عند فتح الصفحة لحين اختيار الشركة
        segments = []
        load_types = []
        vehicles = []
        trailers = []
        drivers = []
        routes = []

        # إذا تم اختيار شركة مسبقاً أو تمرير الـ ID يتم جلب توابعها الخاصة فقط
        selected_client_id = client_company_id
        selected_client = db.query(Client).filter(Client.id == selected_client_id).first() if selected_client_id else None
        
        if selected_client:
            target_transport_company_id = selected_client.company_id
            
            # جلب الأقسام التابعة للعميل
            segments = db.query(Segment).filter(Segment.client_id == selected_client_id).all()
            
            # جلب أنواع الأحمال المرتبطة بالعميل حصرياً
            load_types_raw = db.query(LoadType).filter(
                LoadType.client_id == selected_client_id
            ).all()
            load_types = [{"id": lt.id, "name": lt.name} for lt in load_types_raw]

            # جلب السائقين المرتبطين
            drivers_raw = db.query(Driver).filter(
                Driver.company_id == target_transport_company_id,
                or_(Driver.client_id == selected_client_id, Driver.client_id == None)
            ).all()
            drivers = [{"id": d.id, "name": d.name, "license_expiry": d.license_expiry or "N/A"} for d in drivers_raw]

            # جلب السيارات المرتبطة
            vehicles_raw = db.query(Vehicle).filter(
                Vehicle.company_id == target_transport_company_id,
                or_(Vehicle.client_id == selected_client_id, Vehicle.client_id == None)
            ).all()
            vehicles = [{"id": v.id, "plate_number": v.plate_number, "license_expiry": v.license_expiry or "N/A"} for v in vehicles_raw]

            # جلب المقطورات المرتبطة
            trailers_raw = db.query(Trailer).filter(
                or_(Trailer.company_id == target_transport_company_id, Trailer.company_id == None),
                or_(Trailer.client_id == selected_client_id, Trailer.client_id == None)
            ).all()
            trailers = [{"id": t.id, "trailer_number": t.trailer_number, "license_expiry": t.license_expiry or "N/A"} for t in trailers_raw]

            # جلب خطوط السير المرتبطة بالعميل حصرياً (عبر client_company_id)
            routes_raw = db.query(Route).filter(
                Route.client_company_id == selected_client_id
            ).all()
            
            # حساب الأوقات والمسافات
            setting = db.query(Setting).filter(
                Setting.company_id == target_transport_company_id,
                Setting.client_company_id == selected_client_id
            ).first() or db.query(Setting).filter(Setting.company_id == None, Setting.client_company_id == None).first()
            
            speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed else 40.0
            rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time else 0.5

            for r in routes_raw:
                dist = r.distance_km or 0.0
                driving_hours = dist / speed if speed > 0 else 0
                rest_stops = int(driving_hours // 2) if driving_hours >= 2 else 0
                calculated_total = round(driving_hours + (rest_stops * rest_duration), 1)

                routes.append({
                    "id": r.id,
                    "name": getattr(r, 'code', f"RT-{r.id:04d}"), # اسم أو كود خط السير
                    "route_type_name": getattr(r, 'route_type_name', ''), # نوع خط السير
                    "origin": r.origin_location.name if r.origin_location else "",
                    "destination": r.destination_location.name if r.destination_location else "",
                    "distance_km": dist,
                    "total_time": r.total_time if (hasattr(r, 'total_time') and r.total_time) else calculated_total
                })

        last_journey = db.query(Journey).order_by(Journey.id.desc()).first()
        next_id = (last_journey.id + 1) if last_journey else 1
        next_journey_serial = f"JR-{next_id:04d}"
        
        return templates.TemplateResponse(
            request=request,
            name="new_Journey.html",
            context={
                "clients": clients,
                "load_types": load_types,
                "segments": segments,
                "vehicles": vehicles,
                "trailers": trailers,
                "drivers": drivers,
                "routes": routes,
                "current_user": logged_in_user,
                "next_serial": next_journey_serial,
                "pre_vehicle_id": vehicle_id,
                "pre_driver_id": driver_id,
                "pre_route_id": route_id,
                "pre_load_type_id": load_type_id,
                "pre_trailer_id": trailer_id,
                "pre_segment_id": segment_id,
                "pre_client_company_id": selected_client_id
            }
        )
    except Exception as e:
        print(f"Error in /journeys/new: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/client-dependent-data", response_class=JSONResponse)
def get_client_dependent_data(client_id: int, db: Session = Depends(get_db)):
    try:
        client = db.query(Client).filter(Client.id == client_id).first()
        if not client:
            return {"error": "Client not found"}

        target_company_id = client.company_id

        # الأقسام (Segments)
        segments_raw = db.query(Segment).filter(Segment.client_id == client_id).all()
        segments = [{"id": s.id, "name": s.name} for s in segments_raw]

        # أنواع الأحمال (Load Types) - مرتبطة بالعميل حصرياً
        load_types_raw = db.query(LoadType).filter(
            LoadType.client_id == client_id
        ).all()
        load_types = [{"id": lt.id, "name": lt.name} for lt in load_types_raw]

        # السائقين (Drivers)
        drivers_raw = db.query(Driver).filter(
            Driver.company_id == target_company_id,
            or_(Driver.client_id == client_id, Driver.client_id == None)
        ).all()
        drivers = [{"id": d.id, "name": d.name, "license_expiry": d.license_expiry or "N/A"} for d in drivers_raw]

        # السيارات (Vehicles)
        vehicles_raw = db.query(Vehicle).filter(
            Vehicle.company_id == target_company_id,
            or_(Vehicle.client_id == client_id, Vehicle.client_id == None)
        ).all()
        vehicles = [{"id": v.id, "plate_number": v.plate_number, "license_expiry": v.license_expiry or "N/A"} for v in vehicles_raw]

        # المقطورات (Trailers)
        trailers_raw = db.query(Trailer).filter(
            or_(Trailer.company_id == target_company_id, Trailer.company_id == None),
            or_(Trailer.client_id == client_id, Trailer.client_id == None)
        ).all()
        trailers = [{"id": t.id, "trailer_number": t.trailer_number, "license_expiry": t.license_expiry or "N/A"} for t in trailers_raw]

        # خطوط السير (Routes) - مرتبطة بالعميل حصرياً عبر client_company_id والإعدادات
        setting = db.query(Setting).filter(
            Setting.company_id == target_company_id,
            Setting.client_company_id == client_id
        ).first() or db.query(Setting).filter(Setting.company_id == None, Setting.client_company_id == None).first()
        
        speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed else 40.0
        rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time else 0.5

        routes_raw = db.query(Route).filter(
            Route.client_company_id == client_id
        ).all()
        
        routes = []
        for r in routes_raw:
            dist = r.distance_km or 0.0
            driving_hours = dist / speed if speed > 0 else 0
            rest_stops = int(driving_hours // 2) if driving_hours >= 2 else 0
            calculated_total = round(driving_hours + (rest_stops * rest_duration), 1)

            routes.append({
                "id": r.id,
                "name": getattr(r, 'code', f"RT-{r.id:04d}"), # اسم أو كود خط السير
                "route_type_name": getattr(r, 'route_type_name', ''), # نوع خط السير
                "origin": r.origin_location.name if r.origin_location else "",
                "destination": r.destination_location.name if r.destination_location else "",
                "distance_km": dist,
                "total_time": r.total_time if (hasattr(r, 'total_time') and r.total_time) else calculated_total
            })

        return {
            "segments": segments,
            "load_types": load_types,
            "drivers": drivers,
            "vehicles": vehicles,
            "trailers": trailers,
            "routes": routes
        }
    except Exception as e:
        print(f"Error in client-dependent-data: {str(e)}")
        return {"error": str(e)}


@app.post("/journeys/create")
def create_journey(
    request: Request,
    client_company_id: int = Form(...),
    segment_id: Optional[int] = Form(None),
    driver_id: int = Form(...),
    vehicle_id: int = Form(...),
    trailer_id: Optional[int] = Form(None),
    route_id: int = Form(...),
    load_type_id: int = Form(...),
    departure_time: Optional[str] = Form(None),
    journey_serial: Optional[str] = Form(None),
    include_sleep: str = Form("N"),
    is_off_duty: str = Form("N"),
    remarks: Optional[str] = Form(None),
    db: Session = Depends(get_db)
):
    try:
        logged_in_user = get_current_user(request, db)
        
        client_obj = db.query(Client).filter(Client.id == client_company_id).first()
        company_id = client_obj.company_id if client_obj else None
        
        if not company_id:
            if hasattr(logged_in_user, 'companies') and logged_in_user.companies:
                company_id = logged_in_user.companies[0].id
            else:
                default_company = db.query(Company).first()
                if not default_company:
                    raise HTTPException(status_code=400, detail="لا توجد أي شركة نقل مسجلة في النظام.")
                company_id = default_company.id

        open_journey_filter = Journey.status.in_(["Active", "Planned"])
        
        conflict_driver = db.query(Journey).filter(Journey.driver_id == driver_id, open_journey_filter).first()
        if conflict_driver:
            raise HTTPException(status_code=400, detail=f"عذراً، هذا السائق لديه رحلة نشطة بالفعل (برقم: {conflict_driver.serial}).")

        conflict_vehicle = db.query(Journey).filter(Journey.vehicle_id == vehicle_id, open_journey_filter).first()
        if conflict_vehicle:
            raise HTTPException(status_code=400, detail=f"عذراً، هذه السيارة لديها رحلة نشطة بالفعل (برقم: {conflict_vehicle.serial}).")

        if trailer_id:
            conflict_trailer = db.query(Journey).filter(Journey.trailer_id == trailer_id, open_journey_filter).first()
            if conflict_trailer:
                raise HTTPException(status_code=400, detail=f"عذراً، هذه المقطورة مرتبطة برحلة نشطة حالياً (برقم: {conflict_trailer.serial}).")

        dep_time = datetime.fromisoformat(departure_time) if departure_time and departure_time.strip() != "" else None
        
        # 🔒 توليد السيريال لحظياً عند الحفظ وقفل الجدول مؤقتاً (with_for_update) لمنع التداخل بين المستخدمين
        last_journey = db.query(Journey).order_by(Journey.id.desc()).with_for_update().first()
        
        next_num = 1
        if last_journey and last_journey.serial:
            parts = last_journey.serial.split("-")
            if len(parts) > 1 and parts[-1].isdigit():
                next_num = int(parts[-1]) + 1
        serial_to_use = f"JR-{next_num:04d}"
        
        route_obj = db.query(Route).filter(Route.id == route_id).first()
        
        setting = db.query(Setting).filter(
            Setting.company_id == company_id,
            Setting.client_company_id == client_company_id
        ).first() or db.query(Setting).filter(
            Setting.company_id == None,
            Setting.client_company_id == None
        ).first()
        
        speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed else (float(setting.speed) if setting and hasattr(setting, 'speed') and setting.speed else 40.0)
        rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time else 0.5
        sleep_hours_setting = float(setting.sleep_time) if setting and hasattr(setting, 'sleep_time') and setting.sleep_time else 8.0
        
        route_total_time = 0.0
        if route_obj:
            if hasattr(route_obj, 'total_time') and route_obj.total_time:
                route_total_time = float(route_obj.total_time)
            elif hasattr(route_obj, 'estimated_hours') and route_obj.estimated_hours:
                route_total_time = float(route_obj.estimated_hours)
            elif route_obj.distance_km:
                driving_hours_val = route_obj.distance_km / speed if speed > 0 else 0.0
                rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
                route_total_time = driving_hours_val + (rest_stops * rest_duration)

        total_planned_hours = route_total_time
        if include_sleep == "Y":
            total_planned_hours += sleep_hours_setting
            
        if is_off_duty == "Y":
            total_planned_hours += 24.0

        planned_end = dep_time + timedelta(hours=total_planned_hours) if dep_time else None
        creator_name = getattr(logged_in_user, 'full_name', None) or getattr(logged_in_user, 'username', 'System')

        # حفظ الرحلة الجديدة في قاعدة البيانات بالسيريال المولد لحظياً
        new_journey = Journey(
            serial=serial_to_use,
            company_id=company_id,
            client_company_id=client_company_id,
            segment_id=segment_id,
            driver_id=driver_id,
            vehicle_id=vehicle_id,
            trailer_id=trailer_id,
            route_id=route_id,
            load_type_id=load_type_id,
            manual_start_time=dep_time,
            planned_end_time=planned_end,
            include_sleep=include_sleep,
            is_off_duty=is_off_duty,
            status="Active" if dep_time else "Planned",
            created_by=creator_name,
            remarks=remarks.strip() if remarks and remarks.strip() != "" else None
        )

        db.add(new_journey)
        db.commit()
        db.refresh(new_journey)
        
        return {
            "success": True, 
            "message": "تم إنشاء الرحلة بنجاح!",
            "redirect_url": f"/journeys/ticket/{new_journey.id}"
        }
        
    except HTTPException as he:
        db.rollback()
        raise he
    except Exception as e:
        db.rollback()
        print(f"Error in /journeys/create: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/journeys/tickets", response_class=HTMLResponse)
def list_all_journeys(
    request: Request, 
    q: Optional[str] = None, 
    status: Optional[str] = None,
    load_type_id: Optional[str] = None,
    client_id: Optional[str] = None,
    grouped: Optional[str] = "1",
    sort_by: Optional[str] = None,
    order: Optional[str] = "asc",
    date_type: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    db: Session = Depends(get_db),
    current_user = Depends(get_current_user)
):
    try:
        is_admin_or_dev = (
            getattr(current_user, "is_admin", False) or 
            getattr(current_user, "is_developer", False) or 
            getattr(current_user, "role", "") in ["admin", "developer"]
        )

        user_clients = []
        if is_admin_or_dev:
            user_clients = db.query(Client).all()
        else:
            if hasattr(current_user, "clients") and current_user.clients:
                user_clients = current_user.clients
            elif hasattr(current_user, "companies") and current_user.companies:
                comp_ids = [c.id for c in current_user.companies]
                user_clients = db.query(Client).filter(Client.company_id.in_(comp_ids)).all()

        clients = user_clients
        client_ids_list = [c.id for c in clients]

        query = db.query(Journey)
        
        need_vehicle_join = (not getattr(current_user, 'is_superuser', False) and hasattr(current_user, 'companies')) or bool(q)
        
        if need_vehicle_join:
            query = query.outerjoin(Vehicle, Journey.vehicle_id == Vehicle.id)
            if not getattr(current_user, 'is_superuser', False) and hasattr(current_user, 'companies'):
                allowed_company_ids = [c.id for c in current_user.companies]
                query = query.filter(Vehicle.company_id.in_(allowed_company_ids))
                
            if q:
                query = query.filter(
                    or_(
                        Vehicle.plate_number.ilike(f"%{q}%"),
                        Journey.serial.ilike(f"%{q}%")
                    )
                )
        elif q:
            query = query.outerjoin(Vehicle, Journey.vehicle_id == Vehicle.id).filter(
                or_(
                    Vehicle.plate_number.ilike(f"%{q}%"),
                    Journey.serial.ilike(f"%{q}%")
                )
            )

        need_route_join = bool(client_id and client_id.isdigit()) or (bool(client_ids_list) and not is_admin_or_dev)
        
        if need_route_join:
            query = query.outerjoin(Route, Journey.route_id == Route.id)
            if client_id and client_id.isdigit():
                cid = int(client_id)
                query = query.filter(
                    or_(
                        Route.client_company_id == cid,
                        Journey.segment.has(Segment.client_id == cid) if hasattr(Journey, 'segment') else False
                    )
                )
            elif client_ids_list and not is_admin_or_dev:
                query = query.filter(
                    or_(
                        Route.client_company_id.in_(client_ids_list),
                        Journey.segment.has(Segment.client_id.in_(client_ids_list)) if hasattr(Journey, 'segment') else False
                    )
                )

        if status:
            if status in ["Active", "Ended", "Planned", "Cancelled"]:
                query = query.filter(Journey.status == status, Journey.is_off_duty != "Y")
            elif status == "Active_Planned":
                query = query.filter(Journey.status.in_(["Active", "Planned"]), Journey.is_off_duty != "Y")
            elif status == "Active_Planned_OffDuty":
                query = query.filter(
                    or_(
                        Journey.status.in_(["Active", "Planned"]),
                        Journey.is_off_duty == "Y"
                    )
                )
            elif status == "Off Duty":
                query = query.filter(Journey.is_off_duty == "Y")
        
        if load_type_id and load_type_id.isdigit():
            query = query.filter(Journey.load_type_id == int(load_type_id))
            
        journeys = query.all()

        lt_query = db.query(LoadType)
        if client_id and client_id.isdigit():
            lt_query = lt_query.filter(LoadType.client_id == int(client_id))
        elif not is_admin_or_dev and client_ids_list:
            lt_query = lt_query.filter(LoadType.client_id.in_(client_ids_list))
        load_types_list = lt_query.all()

        # --- توحيد جلب الإعدادات وحساب Planned Arrival تماماً مثل التذكرة ---
        for journey in journeys:
            # استخراج معرف شركة العميل للرحلة بدقة
            client_company_id = None
            if journey.route:
                client_company_id = getattr(journey.route, 'client_company_id', None)
            if not client_company_id:
                client_company_id = getattr(journey, 'company_id', None)

            setting = None
            if client_company_id:
                setting = db.query(Setting).filter(Setting.client_company_id == client_company_id).first()
            if not setting:
                setting = db.query(Setting).filter(
                    Setting.company_id == None,
                    Setting.client_company_id == None
                ).first()

            speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed is not None else 40.0
            rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time is not None else 0.5
            sleep_hours_setting = float(setting.sleep_time) if setting and hasattr(setting, 'sleep_time') and setting.sleep_time is not None else 8.0

            if journey.route and journey.manual_start_time:
                route_obj = journey.route
                driving_hours_val = getattr(route_obj, 'driving_hours', None)
                if not driving_hours_val and route_obj.distance_km:
                    driving_hours_val = route_obj.distance_km / speed if speed > 0 else 0.0
                driving_hours_val = driving_hours_val or 0.0

                rest_time_val = getattr(route_obj, 'rest_time', None)
                if rest_time_val is None and route_obj.distance_km:
                    rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
                    rest_time_val = rest_stops * rest_duration
                rest_time_val = rest_time_val or 0.0

                total_time_val = getattr(route_obj, 'total_time', None)
                if not total_time_val:
                    total_time_val = driving_hours_val + rest_time_val

                total_planned_driving = total_time_val
                if getattr(journey, 'include_sleep', 'N') == "Y":
                    total_planned_driving += sleep_hours_setting

                extra_off_duty = 24.0 if getattr(journey, 'is_off_duty', 'N') == "Y" else 0.0
                journey.planned_end_time = journey.manual_start_time + timedelta(hours=total_planned_driving + extra_off_duty)

            is_off_duty_val = getattr(journey, 'is_off_duty', 'N') == "Y"
            ref_end_time = journey.manual_end_time if is_off_duty_val else journey.manual_arrived_time

            if journey.manual_start_time and ref_end_time and journey.planned_end_time:
                diff_delay = ref_end_time - journey.planned_end_time
                delay_seconds = diff_delay.total_seconds()
                if abs(delay_seconds) < 60:
                    journey.delayed_duration = "00:00"
                else:
                    prefix = "-" if delay_seconds < 0 else "+"
                    abs_seconds = abs(delay_seconds)
                    h = int(abs_seconds // 3600)
                    m = int((abs_seconds % 3600) // 60)
                    journey.delayed_duration = f"{prefix}{h:02d}:{m:02d}"
            else:
                journey.delayed_duration = "00:00"

        # تطبيق الفلاتر الخاصة بالتاريخ والترتيب بعد تحديث الحسابات
        if date_type and (date_from or date_to):
            target_column = None
            if date_type == "departure":
                target_column = Journey.manual_start_time
            elif date_type == "planned_arrival":
                target_column = Journey.planned_end_time
            elif date_type == "actual_arrival":
                target_column = Journey.manual_arrived_time
                
            if target_column is not None:
                if date_from:
                    try:
                        df_str = str(date_from).strip()
                        if df_str and df_str.lower() != "none":
                            dt_from = datetime.fromisoformat(df_str)
                            journeys = [j for j in journeys if getattr(j, target_column.key, None) and getattr(j, target_column.key) >= dt_from]
                    except Exception:
                        pass
                if date_to:
                    try:
                        dt_str = str(date_to).strip()
                        if dt_str and dt_str.lower() != "none":
                            dt_to = datetime.fromisoformat(dt_str)
                            journeys = [j for j in journeys if getattr(j, target_column.key, None) and getattr(j, target_column.key) <= dt_to]
                    except Exception:
                        pass

        if sort_by:
            reverse_order = (order == "desc")
            if sort_by == "departure":
                journeys.sort(key=lambda x: x.manual_start_time or datetime.min, reverse=reverse_order)
            elif sort_by == "planned_arrival":
                journeys.sort(key=lambda x: x.planned_end_time or datetime.min, reverse=reverse_order)
            elif sort_by == "actual_arrival":
                journeys.sort(key=lambda x: x.manual_arrived_time or datetime.min, reverse=reverse_order)
            elif sort_by == "delay":
                journeys.sort(key=lambda x: x.delayed_duration, reverse=reverse_order)
            elif sort_by == "remarks":
                journeys.sort(key=lambda x: x.remarks or "", reverse=reverse_order)
        else:
            journeys.sort(key=lambda x: x.id, reverse=True)
            
        segmented_journeys = {}
        if grouped == "1":
            for journey in journeys:
                segment_name = "General / Unassigned"
                seg_obj = None
                if hasattr(journey, 'segment') and journey.segment:
                    seg_obj = journey.segment
                elif journey.route:
                    if hasattr(journey.route, 'segment') and journey.route.segment:
                        seg_obj = journey.route.segment
                    elif hasattr(journey.route, 'segment_name') and journey.route.segment_name:
                        seg_obj = journey.route.segment_name
                
                if seg_obj:
                    if isinstance(seg_obj, str):
                        segment_name = seg_obj
                    else:
                        for attr in ['name', 'title', 'segment_name', 'label']:
                            if hasattr(seg_obj, attr) and getattr(seg_obj, attr):
                                segment_name = str(getattr(seg_obj, attr))
                                break
                
                if segment_name not in segmented_journeys:
                    segmented_journeys[segment_name] = []
                segmented_journeys[segment_name].append(journey)
        
        return templates.TemplateResponse(
            request=request,
            name="Journeys_list.html",
            context={
                "journeys": journeys,
                "segmented_journeys": segmented_journeys,
                "load_types_list": load_types_list,
                "clients": clients,
                "search_query": q or "",
                "current_status": status or "",
                "current_load_type_id": load_type_id or "",
                "current_client_id": client_id or "",
                "grouped": grouped,
                "current_sort": sort_by or "",
                "current_order": order or "asc",
                "date_type": date_type or "",
                "date_from": date_from or "",
                "date_to": date_to or ""
            }
        )
    except Exception as e:
        print(f"Error in /journeys/tickets: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/journeys/segment-update/{segment_name}", response_class=HTMLResponse)
def segment_last_update(segment_name: str, request: Request, db: Session = Depends(get_db)):
    try:
        active_planned_filter = Journey.status.in_(["Active", "Planned"])
        segment_obj = db.query(Segment).filter(Segment.name.ilike(f"%{segment_name}%")).first()
        journeys = []
        if segment_obj:
            journeys = db.query(Journey).options(
                joinedload(Journey.follow_ups),
                joinedload(Journey.vehicle),
                joinedload(Journey.trailer),
                joinedload(Journey.driver)
            ).filter(
                Journey.segment_id == segment_obj.id,
                active_planned_filter
            ).all()
        
        if not journeys:
            query = db.query(Journey).options(
                joinedload(Journey.follow_ups),
                joinedload(Journey.vehicle),
                joinedload(Journey.trailer),
                joinedload(Journey.driver)
            ).outerjoin(Company, Journey.company_id == Company.id)
            
            if 'Route' in globals():
                query = query.outerjoin(Route, Journey.route_id == Route.id)
                
            journeys = query.filter(
                or_(
                    Company.name.ilike(f"%{segment_name}%"),
                    Route.origin.ilike(f"%{segment_name}%"),
                    Route.destination.ilike(f"%{segment_name}%")
                ),
                active_planned_filter
            ).all()
            
        for j in journeys:
            latest_fu = db.query(JourneyFollowUp).filter(
                JourneyFollowUp.journey_id == j.id,
                JourneyFollowUp.is_deleted == False
            ).order_by(JourneyFollowUp.id.desc()).first()
            
            j.latest_comment = latest_fu.comment if latest_fu and latest_fu.comment else None
            j.latest_location = latest_fu.location if latest_fu and latest_fu.location else None
            
    except Exception as e:
        print(f"Error in segment_last_update: {e}")
        active_planned_filter = Journey.status.in_(["Active", "Planned"])
        journeys = db.query(Journey).options(
            joinedload(Journey.follow_ups),
            joinedload(Journey.vehicle),
            joinedload(Journey.trailer),
            joinedload(Journey.driver)
        ).filter(active_planned_filter).all()
        
        for j in journeys:
            latest_fu = db.query(JourneyFollowUp).filter(
                JourneyFollowUp.journey_id == j.id,
                JourneyFollowUp.is_deleted == False
            ).order_by(JourneyFollowUp.id.desc()).first()
            
            j.latest_comment = latest_fu.comment if latest_fu and latest_fu.comment else None
            j.latest_location = latest_fu.location if latest_fu and latest_fu.location else None
    
    return templates.TemplateResponse(request, "segment_update.html", {
        "request": request,
        "segment_name": segment_name,
        "journeys": journeys
    })


@app.get("/journeys/ticket/{journey_id}", response_class=HTMLResponse)
def journey_ticket_page(journey_id: int, request: Request, response: Response, db: Session = Depends(get_db)):
    try:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"

        logged_in_user = get_current_user(request, db)
        if not logged_in_user:
            return RedirectResponse(url="/login", status_code=303)
            
        current_user_id = getattr(logged_in_user, 'id', None)
        current_user_name = logged_in_user.email if hasattr(logged_in_user, 'email') else "Admin"

        journey = db.query(Journey).options(
            joinedload(Journey.maintenances),
            joinedload(Journey.stops),
            joinedload(Journey.driver),
            joinedload(Journey.vehicle),
            joinedload(Journey.trailer),
            joinedload(Journey.route).joinedload(Route.origin_location),
            joinedload(Journey.route).joinedload(Route.destination_location),
            joinedload(Journey.load_type),
            joinedload(Journey.company),
            joinedload(Journey.segment)
        ).filter(Journey.id == journey_id).first()
        
        if not journey:
            raise HTTPException(status_code=404, detail="Journey not found")
            
        # --- نظام القفل المحدث (بدون مدة زمنية - بمجرد الخروج تصبح متاحة) ---
        if journey.locked_by and journey.locked_by != current_user_id:
            locking_user = db.query(User).filter(User.id == journey.locked_by).first()
            locker_name = locking_user.username if locking_user and hasattr(locking_user, 'username') else "مستخدم آخر"
            serial_val = journey.serial or f"JR-{journey.id:04d}"
            
            raise HTTPException(
                status_code=403, 
                detail=f"عذراً، هذه التذكرة ({serial_val}) مفتوحة حالياً من قبل المستخدم: {locker_name}"
            )
                
        journey.locked_by = current_user_id
        journey.lock_expires_at = None  # إلغاء الاعتماد على الوقت
        db.commit()
            
        drivers = db.query(Driver).all()
        routes = db.query(Route).all()
        vehicles = db.query(Vehicle).all()
        trailers = db.query(Trailer).all()
        
        # --- تجهيز اسم المسار (Route Name) بشكل مطور وشامل لجميع الاحتمالات ---
        route_name = "N/A"
        if journey.route:
            r = journey.route
            origin_name = None
            destination_name = None
            
            if hasattr(r, 'origin_location') and r.origin_location:
                origin_name = getattr(r.origin_location, 'name', None)
            if not origin_name:
                origin_name = getattr(r, 'origin', None) or getattr(r, 'from_city', None) or getattr(r, 'source', None)
                
            if hasattr(r, 'destination_location') and r.destination_location:
                destination_name = getattr(r.destination_location, 'name', None)
            if not destination_name:
                destination_name = getattr(r, 'destination', None) or getattr(r, 'to_city', None)
            
            if origin_name and destination_name:
                route_name = f"{origin_name} -> {destination_name}"
            elif hasattr(r, 'name') and r.name:
                route_name = r.name
            elif hasattr(r, 'code') and r.code:
                route_name = r.code
            else:
                route_name = f"RT-{r.id:04d}"

        # --- استخراج اسم العميل (Client) ومعرف شركة العميل بشكل صحيح ودقيق ---
        client_name = "N/A"
        client_company_id = None

        if journey.route:
            if hasattr(journey.route, 'client_company_id') and journey.route.client_company_id:
                client_company_id = journey.route.client_company_id
            
            if hasattr(journey.route, 'client_company') and journey.route.client_company:
                client_company_obj = journey.route.client_company
                client_name = getattr(client_company_obj, 'name', str(client_company_obj))
                if not client_company_id:
                    client_company_id = getattr(client_company_obj, 'id', None)
            elif hasattr(journey.route, 'client_name') and journey.route.client_name:
                client_name = journey.route.client_name

        if client_name == "N/A" and hasattr(journey, 'client') and journey.client:
            client_name = getattr(journey.client, 'name', str(journey.client))
            if not client_company_id:
                client_company_id = getattr(journey.client, 'id', None)

        if client_name == "N/A" and hasattr(journey, 'company_id') and journey.company_id:
            client_company_id = journey.company_id

        if client_name == "N/A" and client_company_id:
            try:
                comp_obj = db.query(Company).filter(Company.id == client_company_id).first()
                if comp_obj and hasattr(comp_obj, 'name'):
                    client_name = comp_obj.name
            except Exception:
                pass

        raw_followups = db.query(JourneyFollowUp).filter(
            JourneyFollowUp.journey_id == journey_id,
            JourneyFollowUp.is_deleted == False
        ).order_by(JourneyFollowUp.id.asc()).all()
        
        followups = []
        for index, f in enumerate(raw_followups):
            followups.append({
                "id": f.id,
                "employee_name": f.employee_name,
                "comment": f.comment,
                "location": getattr(f, 'location', None),
                "created_at": getattr(f, 'created_at', None),
                "display_index": index + 1
            })
            
        journey_maintenance = []
        if journey.maintenances:
            journey_maintenance = [m for m in journey.maintenances if not getattr(m, 'is_deleted', False)]
            
        journey_stops = []
        if journey.stops:
            journey_stops = journey.stops
        
        # --- جلب الإعدادات (Settings) الخاصة بالـ Client أو الإعدادات العامة كافتراضي ---
        setting = None
        if client_company_id:
            setting = db.query(Setting).filter(Setting.client_company_id == client_company_id).first()
            
        if not setting:
            setting = db.query(Setting).filter(
                Setting.company_id == None,
                Setting.client_company_id == None
            ).first()

        speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed is not None else 40.0
        rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time is not None else 0.5
        sleep_hours_setting = float(setting.sleep_time) if setting and hasattr(setting, 'sleep_time') and setting.sleep_time is not None else 8.0

        route_obj = journey.route
        
        driving_hours_val = getattr(route_obj, 'driving_hours', None) if route_obj else None
        if not driving_hours_val and route_obj and route_obj.distance_km:
            driving_hours_val = route_obj.distance_km / speed if speed > 0 else 0.0
        driving_hours_val = driving_hours_val or 0.0

        rest_time_val = getattr(route_obj, 'rest_time', None) if route_obj else None
        if rest_time_val is None and route_obj and route_obj.distance_km:
            rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
            rest_time_val = rest_stops * rest_duration
        rest_time_val = rest_time_val or 0.0

        total_time_val = getattr(route_obj, 'total_time', None) if route_obj else None
        if not total_time_val:
            total_time_val = driving_hours_val + rest_time_val

        def hours_to_hhmm(h_val):
            if not h_val or h_val < 0:
                h_val = 0.0
            h = int(h_val)
            m = int((h_val - h) * 60)
            return f"{h:02d}:{m:02d}"

        driving_hours = hours_to_hhmm(driving_hours_val)
        rest_time = hours_to_hhmm(rest_time_val)
        total_time = hours_to_hhmm(total_time_val)

        total_planned_driving = total_time_val
        if getattr(journey, 'include_sleep', 'N') == "Y":
            total_planned_driving += sleep_hours_setting

        if journey.manual_start_time and isinstance(journey.manual_start_time, datetime):
            extra_off_duty = 24.0 if getattr(journey, 'is_off_duty', 'N') == "Y" else 0.0
            journey.planned_end_time = journey.manual_start_time + timedelta(hours=total_planned_driving + extra_off_duty)
        else:
            journey.planned_end_time = None

        if journey.manual_start_time and journey.manual_arrived_time:
            diff_actual = journey.manual_arrived_time - journey.manual_start_time
            actual_tot_hours = diff_actual.total_seconds() / 3600
            
            # تم تعديل هذا الجزء بحيث لا تؤثر خانة include_sleep نهائياً على ساعات القيادة الفعلية (Actual Hours)
            
            total_maintenance_hours = sum(m.duration_hours for m in journey_maintenance if m.duration_hours)
            actual_tot_hours -= total_maintenance_hours
            
            total_stops_hours = sum(s.hours_count for s in journey_stops if s.hours_count)
            actual_tot_hours -= total_stops_hours
                
            if actual_tot_hours < 0:
                actual_tot_hours = 0.0
            h_act = int(actual_tot_hours)
            m_act = int((actual_tot_hours - h_act) * 60)
            journey.actual_hours = f"{h_act:02d}:{m_act:02d}"
        else:
            journey.actual_hours = "00:00"

        ref_end_time = journey.manual_end_time if getattr(journey, 'is_off_duty', 'N') == "Y" else journey.manual_arrived_time
        
        if ref_end_time and journey.planned_end_time:
            diff_delay = ref_end_time - journey.planned_end_time
            delay_seconds = diff_delay.total_seconds()
            if abs(delay_seconds) < 60:
                journey.delayed_duration = "00:00"
            else:
                prefix = "-" if delay_seconds < 0 else "+"
                abs_seconds = abs(delay_seconds)
                h = int(abs_seconds // 3600)
                m = int((abs_seconds % 3600) // 60)
                journey.delayed_duration = f"{prefix}{h:02d}:{m:02d}"
        else:
            journey.delayed_duration = "00:00"

        if journey.manual_start_time and journey.manual_end_time:
            diff_total = journey.manual_end_time - journey.manual_start_time
            tot_seconds = diff_total.total_seconds()
            if tot_seconds < 0:
                tot_seconds = 0
            tot_h = int(tot_seconds // 3600)
            tot_m = int((tot_seconds % 3600) // 60)
            journey.total_journey_hours = f"{tot_h:02d}:{tot_m:02d}"
        else:
            journey.total_journey_hours = "00:00"

        def check_license_status(expiry_val):
            if not expiry_val:
                return {'date': 'غير مسجل', 'expired': False, 'warning': False}
            try:
                if isinstance(expiry_val, str):
                    exp_date = datetime.strptime(expiry_val.split()[0], '%Y-%m-%d').date()
                elif isinstance(expiry_val, datetime):
                    exp_date = expiry_val.date()
                else:
                    exp_date = expiry_val
            except Exception:
                return {'date': str(expiry_val), 'expired': False, 'warning': False}

            today = datetime.now().date()
            days_left = (exp_date - today).days
            return {
                'date': exp_date.strftime('%Y-%m-%d'),
                'expired': days_left < 0,
                'warning': 0 <= days_left <= 30
            }

        driver_lic_status = check_license_status(getattr(journey.driver, 'license_expiry', None) if journey.driver else None)
        vehicle_lic_status = check_license_status(getattr(journey.vehicle, 'license_expiry', None) if journey.vehicle else None)
        trailer_lic_status = check_license_status(getattr(journey.trailer, 'license_expiry', None) if journey.trailer else None)

        licenses_info = {
            "driver_license": driver_lic_status['date'],
            "driver_expired": driver_lic_status['expired'],
            "driver_warning": driver_lic_status['warning'],
            "vehicle_license": vehicle_lic_status['date'],
            "vehicle_expired": vehicle_lic_status['expired'],
            "vehicle_warning": vehicle_lic_status['warning'],
            "trailer_license": trailer_lic_status['date'],
            "trailer_expired": trailer_lic_status['expired'],
            "trailer_warning": trailer_lic_status['warning']
        }

        contact_info = {
            "driver_phone": getattr(journey.driver, 'phone', None) or 'غير مسجل' if journey.driver else 'غير مسجل',
            "vehicle_phone": getattr(journey.vehicle, 'phone', None) or getattr(journey.vehicle, 'phone_number', 'غير مسجل') if journey.vehicle else 'غير مسجل'
        }

        return templates.TemplateResponse(
            request=request,
            name="Journey.html",
            context={
                "journey": journey,
                "drivers": drivers,
                "routes": routes,
                "vehicles": vehicles,
                "trailers": trailers,
                "current_user_name": current_user_name,
                "followups": followups,
                "journey_maintenance": journey_maintenance,
                "journey_stops": journey_stops,
                "planned_arrival_time": journey.planned_end_time,
                "licenses_info": licenses_info,
                "contact_info": contact_info,
                "driving_hours": driving_hours,
                "rest_time": rest_time,
                "total_time": total_time,
                "client_name": client_name,
                "route_name": route_name
            }
        )
    except HTTPException as he:
        raise he
    except Exception as e:
        print(f"Error in /journeys/ticket/{journey_id}: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/journeys/map/{journey_id}", response_class=HTMLResponse)
def journey_map_page(journey_id: int, request: Request, db: Session = Depends(get_db)):
    try:
        logged_in_user = get_current_user(request, db)
        if not logged_in_user:
            return RedirectResponse(url="/login", status_code=303)

        # جلب الرحلة مع كافة العلاقات بما فيها التوقفات والصيانة وخط السير ونقاط عبر الطريق (Via Points) مع تفاصيل الأماكن الخاصة بها
        journey = db.query(Journey).options(
            joinedload(Journey.driver),
            joinedload(Journey.vehicle),
            joinedload(Journey.trailer),
            joinedload(Journey.route).joinedload(Route.origin_location),
            joinedload(Journey.route).joinedload(Route.destination_location),
            joinedload(Journey.route).joinedload(Route.via_points).joinedload(RouteViaPoint.location), # جلب النقاط الوسيطة مع بيانات موقعها الجغرافي
            joinedload(Journey.follow_ups),
            joinedload(Journey.stops),
            joinedload(Journey.maintenances)
        ).filter(Journey.id == journey_id).first()

        if not journey:
            raise HTTPException(status_code=404, detail="Journey not found")

        origin_lat = 30.0444
        origin_lng = 31.2357
        dest_lat = 31.2001
        dest_lng = 29.9187

        if journey.route:
            r = journey.route
            if hasattr(r, 'origin_location') and r.origin_location:
                origin_lat = float(getattr(r.origin_location, 'latitude', origin_lat) or origin_lat)
                origin_lng = float(getattr(r.origin_location, 'longitude', origin_lng) or origin_lng)
            if hasattr(r, 'destination_location') and r.destination_location:
                dest_lat = float(getattr(r.destination_location, 'latitude', dest_lat) or dest_lat)
                dest_lng = float(getattr(r.destination_location, 'longitude', dest_lng) or dest_lng)

        # تجهيز النقاط الوسيطة (Via Points) بناءً على جدول RouteViaPoint المرتبط بالعلاقة via_points
        waypoints_data = []
        if journey.route and hasattr(journey.route, 'via_points') and journey.route.via_points:
            # ترتيب النقاط تصاعدياً حسب sequence_order
            sorted_via_points = sorted(journey.route.via_points, key=lambda x: getattr(x, 'sequence_order', 0))
            for vp in sorted_via_points:
                if vp.location:
                    waypoints_data.append({
                        "id": getattr(vp, 'id', None),
                        "name": getattr(vp.location, 'name', 'Waypoint'),
                        "latitude": float(vp.location.latitude),
                        "longitude": float(vp.location.longitude),
                        "order": getattr(vp, 'sequence_order', 0)
                    })

        # إعدادات السرعة ووقت الراحة
        client_company_id = getattr(journey.route, 'client_company_id', None) if journey.route else getattr(journey, 'company_id', None)
        
        # البحث عن الإعدادات بنفس تسلسل الهرمية المعتمد
        setting = None
        if journey.route and journey.route.company_id and client_company_id:
            setting = db.query(Setting).filter(
                Setting.company_id == journey.route.company_id,
                Setting.client_company_id == client_company_id
            ).first()
        if not setting and journey.route and journey.route.company_id:
            setting = db.query(Setting).filter(
                Setting.company_id == journey.route.company_id,
                Setting.client_company_id == None
            ).first()
        if not setting:
            setting = db.query(Setting).filter(Setting.company_id == None, Setting.client_company_id == None).first()

        speed = float(setting.default_speed) if setting and hasattr(setting, 'default_speed') and setting.default_speed is not None else 40.0
        rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time is not None else 0.5

        # حساب المدة بدقة طبقاً لمنطق خطوط السير (مسافة المسار / السرعة + وقت الراحة)
        duration_hours = float(getattr(journey, 'duration_hours', 0.0) or 0.0)
        if duration_hours <= 0:
            route_distance = float(getattr(journey.route, 'distance_km', 0.0) or 0.0) if journey.route else 0.0
            if route_distance > 0 and speed > 0:
                d_hours = route_distance / speed
                rest_stops = int(d_hours // 2) if d_hours >= 2 else 0
                r_time = rest_stops * rest_duration
                duration_hours = round(d_hours + r_time, 2)
            else:
                duration_hours = 5.0  # قيمة افتراضية في حال عدم توفر مسافة

        # تجهيز المتابعات
        followups_data = []
        if journey.follow_ups:
            for f in sorted(journey.follow_ups, key=lambda x: x.id):
                if not getattr(f, 'is_deleted', False):
                    followups_data.append({
                        "id": f.id,
                        "comment": f.comment,
                        "location": getattr(f, 'location', ''),
                        "employee_name": f.employee_name,
                        "created_at": f.created_at.strftime('%Y-%m-%d %H:%M') if f.created_at else ''
                    })

        # تجهيز التوقفات (Stops) مع دعم الاحتمالين لأسماء أعمدة الإحداثيات (latitude أو lat)
        stops_data = []
        if journey.stops:
            for s in journey.stops:
                s_lat = getattr(s, 'latitude', None) or getattr(s, 'lat', None)
                s_lng = getattr(s, 'longitude', None) or getattr(s, 'lng', None)
                stops_data.append({
                    "id": s.id,
                    "location": s.location,
                    "latitude": float(s_lat) if s_lat else None,
                    "longitude": float(s_lng) if s_lng else None,
                    "reason": s.reason,
                    "start_time": s.start_time.strftime('%Y-%m-%d %H:%M') if s.start_time else '',
                    "end_time": s.end_time.strftime('%Y-%m-%d %H:%M') if s.end_time else None,
                    "is_active": s.end_time is None
                })

        # تجهيز سجلات الصيانة (Maintenances) مع دعم الاحتمالين لأسماء أعمدة الإحداثيات
        maintenances_data = []
        if journey.maintenances:
            for m in journey.maintenances:
                if not getattr(m, 'is_deleted', False):
                    m_lat = getattr(m, 'latitude', None) or getattr(m, 'lat', None)
                    m_lng = getattr(m, 'longitude', None) or getattr(m, 'lng', None)
                    maintenances_data.append({
                        "id": m.id,
                        "maintenance_type": m.maintenance_type,
                        "description": m.description,
                        "location": m.location,
                        "latitude": float(m_lat) if m_lat else None,
                        "longitude": float(m_lng) if m_lng else None,
                        "start_time": m.start_time.strftime('%Y-%m-%d %H:%M') if m.start_time else '',
                        "end_time": m.end_time.strftime('%Y-%m-%d %H:%M') if m.end_time else None,
                        "is_active": m.end_time is None
                    })

        # وقت المغادرة بصيغة ISO
        raw_departure = journey.manual_start_time if journey.manual_start_time else getattr(journey, 'departure_time', None)
        if hasattr(raw_departure, 'strftime'):
            departure_time_val = raw_departure.strftime('%Y-%m-%dT%H:%M:%S')
        elif raw_departure:
            departure_time_val = str(raw_departure).replace(' ', 'T')
        else:
            departure_time_val = ""

        # التحقق من وجود توقف نشط أو صيانة جارية
        has_active_stop = any(s.end_time is None for s in (journey.stops or []))
        has_active_maintenance = any(m.end_time is None and not getattr(m, 'is_deleted', False) for m in (journey.maintenances or []))

        return templates.TemplateResponse(
            request=request,
            name="JourneyMap.html",
            context={
                "journey": journey,
                "origin_lat": origin_lat,
                "origin_lng": origin_lng,
                "dest_lat": dest_lat,
                "dest_lng": dest_lng,
                "waypoints_data": waypoints_data, # إرسال النقاط الوسيطة المجهزة للـ Template
                "speed": speed,
                "departure_time": departure_time_val,
                "duration_hours": duration_hours,
                "followups_data": followups_data,
                "stops_data": stops_data,
                "maintenances_data": maintenances_data,
                "has_active_stop": has_active_stop,
                "has_active_maintenance": has_active_maintenance
            }
        )
    except Exception as e:
        print(f"Error in journey map page: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/journeys/ticket/{journey_id}/unlock")
def unlock_journey_ticket(journey_id: int, db: Session = Depends(get_db)):
    try:
        journey = db.query(Journey).filter(Journey.id == journey_id).first()
        if journey:
            journey.locked_by = None
            journey.lock_expires_at = None
            db.commit()
        return {"status": "success"}
    except Exception as e:
        print(f"Error unlocking journey {journey_id}: {str(e)}")
        return {"status": "error"}

@app.post("/journeys/{journey_id}/suspend")
def suspend_journey(
    journey_id: int,
    location: str = Form(...),
    reason: str = Form(...),
    start_time: str = Form(...),
    end_time: Optional[str] = Form(None),      # وقت النهاية اختياري ليتم تسجيل البدء أولاً
    hours_count: Optional[float] = Form(None),  # يتم حسابها تلقائياً
    recorded_by: str = Form(...),
    db: Session = Depends(get_db)
):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey:
        raise HTTPException(status_code=404, detail="Journey not found")
    
    open_maintenance = db.query(JourneyMaintenance).filter(
        JourneyMaintenance.journey_id == journey_id,
        JourneyMaintenance.end_time.is_(None)
    ).first()
    
    if open_maintenance:
        raise HTTPException(
            status_code=400, 
            detail="لا يمكن تسجيل توقف تشغيلي حالياً، يوجد سجل صيانة مفتوح لهذه الرحلة. يجب إنهاؤه أولاً."
        )
    
    parsed_start_time = datetime.fromisoformat(start_time) if start_time else None
    parsed_end_time = datetime.fromisoformat(end_time) if end_time and end_time.strip() != "" else None
    
    calculated_hours = hours_count
    if parsed_start_time and parsed_end_time:
        duration_seconds = (parsed_end_time - parsed_start_time).total_seconds()
        calculated_hours = round(duration_seconds / 3600.0, 2)

    new_stop = JourneyStop(
        journey_id=journey_id,
        location=location,
        reason=reason,
        start_time=parsed_start_time,
        end_time=parsed_end_time,
        hours_count=calculated_hours,
        recorded_by=recorded_by
    )
    db.add(new_stop)
    journey.status = "Suspended"
        
    db.commit()
    
    return RedirectResponse(url=f"/journeys/ticket/{journey_id}", status_code=303)


@app.post("/journeys/stop/update/{stop_id}")
def update_journey_stop(
    stop_id: int,
    location: Optional[str] = Form(None),
    reason: Optional[str] = Form(None),
    start_time: Optional[str] = Form(None),
    end_time: Optional[str] = Form(None),
    hours_count: Optional[float] = Form(None),
    db: Session = Depends(get_db)
):
    stop = db.query(JourneyStop).filter(JourneyStop.id == stop_id).first()
    if not stop:
        raise HTTPException(status_code=404, detail="Stop record not found")
    
    if location:
        stop.location = location
    if reason:
        stop.reason = reason
    if start_time and start_time.strip() != "":
        stop.start_time = datetime.fromisoformat(start_time)
    if end_time is not None:
        stop.end_time = datetime.fromisoformat(end_time) if end_time.strip() != "" else None
        
    if stop.start_time and stop.end_time:
        duration_seconds = (stop.end_time - stop.start_time).total_seconds()
        stop.hours_count = round(duration_seconds / 3600.0, 2)
    elif hours_count is not None:
        stop.hours_count = hours_count
    elif end_time is not None and end_time.strip() == "":
        stop.hours_count = None
        
    db.commit()
    return RedirectResponse(url=f"/journeys/ticket/{stop.journey_id}", status_code=303)


@app.get("/journeys/stop/delete/{stop_id}")
def delete_journey_stop(
    stop_id: int,
    db: Session = Depends(get_db)
):
    stop = db.query(JourneyStop).filter(JourneyStop.id == stop_id).first()
    if not stop:
        raise HTTPException(status_code=404, detail="Stop record not found")
    
    journey_id = stop.journey_id
    db.delete(stop)
    db.commit()
    
    return RedirectResponse(url=f"/journeys/ticket/{journey_id}", status_code=303)


@app.post("/journeys/ticket/cancel/{journey_id}")
def cancel_journey_ticket(journey_id: int, db: Session = Depends(get_db)):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey:
        raise HTTPException(status_code=404, detail="Journey not found")
    
    journey.status = "Cancelled"
    journey.locked_by = None
    journey.lock_expires_at = None
    db.commit()
    return RedirectResponse(url=f"/journeys/ticket/{journey.id}", status_code=303)


@app.post("/journeys/ticket/update/{journey_id}")
def update_journey_ticket(
    journey_id: int,
    request: Request,
    driver_id: int = Form(...),
    route_id: int = Form(...),
    include_sleep: str = Form("N"),
    is_off_duty: str = Form("N"),
    manual_start_time: Optional[str] = Form(None),
    manual_arrived_time: Optional[str] = Form(None),
    manual_end_time: Optional[str] = Form(None),
    status: Optional[str] = Form("Planned"),
    remarks: Optional[str] = Form(None),
    db: Session = Depends(get_db)
):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey:
        raise HTTPException(status_code=404, detail="Journey not found")

    logged_in_user = get_current_user(request, db)
    current_user_name = "System"
    if logged_in_user:
        current_user_name = getattr(logged_in_user, 'email', None) or getattr(logged_in_user, 'username', None) or "System"

    # Unicode Left-to-Right Mark to fix text rendering order for English names in Arabic sentences
    lrm = "\u200e"

    # 1. Check and log driver change
    old_driver_id = journey.driver_id
    if driver_id and driver_id != old_driver_id:
        old_driver = db.query(Driver).filter(Driver.id == old_driver_id).first() if old_driver_id else None
        new_driver = db.query(Driver).filter(Driver.id == driver_id).first()
        
        old_driver_name = getattr(old_driver, 'name', None) or getattr(old_driver, 'full_name', None) or (f"Driver #{old_driver_id}" if old_driver_id else "None")
        new_driver_name = getattr(new_driver, 'name', None) or getattr(new_driver, 'full_name', None) or f"Driver #{driver_id}"
        
        followup_comment = f"Driver changed from ({lrm}{old_driver_name}{lrm}) to ({lrm}{new_driver_name}{lrm})"
        driver_followup = JourneyFollowUp(
            journey_id=journey_id,
            employee_name=current_user_name,
            comment=followup_comment,
            location="Driver Change"
        )
        db.add(driver_followup)

    def get_route_str(r):
        if not r:
            return "None"
        origin_name = None
        destination_name = None
        if hasattr(r, 'origin_location') and r.origin_location:
            origin_name = getattr(r.origin_location, 'name', None)
        if not origin_name:
            origin_name = getattr(r, 'origin', None) or getattr(r, 'from_city', None) or getattr(r, 'source', None)
            
        if hasattr(r, 'destination_location') and r.destination_location:
            destination_name = getattr(r.destination_location, 'name', None)
        if not destination_name:
            destination_name = getattr(r, 'destination', None) or getattr(r, 'to_city', None)
        
        # إضافة كود المسار أو رقم الـ ID للتمييز بين الذهاب والعودة
        route_code = getattr(r, 'code', None) or f"RT-{r.id:04d}"
        
        if origin_name and destination_name:
            return f"{origin_name} -> {destination_name} ({route_code})"
        elif hasattr(r, 'name') and r.name:
            return f"{r.name} ({route_code})"
        else:
            return f"Route #{r.id} ({route_code})"

    # 2. Check and log route change
    old_route_id = journey.route_id
    if route_id and route_id != old_route_id:
        old_route = db.query(Route).filter(Route.id == old_route_id).first() if old_route_id else None
        new_route = db.query(Route).filter(Route.id == route_id).first()
        
        old_route_name = get_route_str(old_route)
        new_route_name = get_route_str(new_route)
        
        route_followup_comment = f"Route changed from ({lrm}{old_route_name}{lrm}) to ({lrm}{new_route_name}{lrm})"
        route_followup = JourneyFollowUp(
            journey_id=journey_id,
            employee_name=current_user_name,
            comment=route_followup_comment,
            location="Route Change"
        )
        db.add(route_followup)

    if driver_id:
        driver = db.query(Driver).filter(Driver.id == driver_id).first()
        if not driver or driver.client_id != journey.client_company_id:
            raise HTTPException(
                status_code=400, 
                detail="Sorry, the selected driver does not belong to this journey's client company."
            )

        active_journey = db.query(Journey).filter(
            Journey.driver_id == driver_id,
            Journey.id != journey_id,
            Journey.status.in_(["Active", "Planned"])
        ).first()
        
        if active_journey:
            raise HTTPException(
                status_code=400, 
                detail=f"Sorry, this driver is already busy in another active or planned journey (JR-{active_journey.id:04d})."
            )

    if route_id:
        route = db.query(Route).filter(Route.id == route_id).first()
        if not route or route.client_company_id != journey.client_company_id:
            raise HTTPException(
                status_code=400, 
                detail="Sorry, the selected route does not belong to this journey's client company."
            )

    new_start_time = datetime.fromisoformat(manual_start_time) if manual_start_time and manual_start_time.strip() != "" else journey.manual_start_time

    if status == "Cancelled":
        journey.status = "Cancelled"
    elif new_start_time and status != "Ended":
        journey.status = "Active"
    else:
        journey.status = status if status else journey.status

    if journey.status == "Ended" and not manual_arrived_time and not journey.manual_arrived_time:
        raise HTTPException(
            status_code=400, 
            detail="Sorry, you must record the actual arrival home time before ending the journey."
        )

    journey.driver_id = driver_id
    journey.route_id = route_id
    journey.include_sleep = include_sleep
    journey.is_off_duty = is_off_duty
    journey.remarks = remarks.strip() if remarks and remarks.strip() != "" else None
    
    if manual_start_time is not None:
        journey.manual_start_time = datetime.fromisoformat(manual_start_time) if manual_start_time.strip() != "" else None
        
    if manual_arrived_time is not None:
        journey.manual_arrived_time = datetime.fromisoformat(manual_arrived_time) if manual_arrived_time.strip() != "" else None
        
    if manual_end_time is not None:
        journey.manual_end_time = datetime.fromisoformat(manual_end_time) if manual_end_time.strip() != "" else None

    if journey.is_off_duty == "N":
        db.query(OffDutyAssignment).filter(OffDutyAssignment.journey_id == journey.id).delete()

    journey.locked_by = None
    journey.lock_expires_at = None

    db.commit()
    return RedirectResponse(url=f"/journeys/ticket/{journey.id}", status_code=303)


@app.post("/journeys/ticket/release/{journey_id}")
def release_journey_ticket(journey_id: int, db: Session = Depends(get_db)):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if journey:
        journey.locked_by = None
        journey.lock_expires_at = None
        db.commit()
    return {"status": "success"}


@app.post("/journeys/ticket/{journey_id}/maintenance/add")
def add_journey_maintenance(
    journey_id: int, 
    maintenance_type: str = Form(...),  
    description: str = Form(...),
    location: str = Form(None),
    start_time: str = Form(...),
    end_time: str = Form(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    # الشرط الذكي: التحقق من عدم وجود توقف تشغيلي "مفتوح" في جدول JourneyStop
    open_stop = db.query(JourneyStop).filter(
        JourneyStop.journey_id == journey_id,
        JourneyStop.end_time.is_(None)
    ).first()
    
    if open_stop:
        raise HTTPException(
            status_code=400, 
            detail="لا يمكن تسجيل صيانة حالياً، يوجد توقف تشغيلي مفتوح لهذه الرحلة. يجب إنهاؤه أولاً."
        )

    try:
        start_dt = datetime.fromisoformat(start_time) if start_time else None
        end_dt = None
        duration_hours = 0.0
        
        if end_time and end_time.strip():
            end_dt = datetime.fromisoformat(end_time)
            if start_dt and end_dt <= start_dt:
                raise HTTPException(
                    status_code=400, detail="End time must be after start time."
                )
            if start_dt:
                duration_hours = round((end_dt - start_dt).total_seconds() / 3600.0, 2)
            
    except ValueError:
        raise HTTPException(
            status_code=400, detail="Invalid date and time format."
        )

    user_name = getattr(current_user, "username", None) or getattr(current_user, "email", None) or str(current_user)

    # حفظ البيانات في جدول JourneyMaintenance المنفصل
    new_maintenance = JourneyMaintenance(
        journey_id=journey_id,
        maintenance_type=maintenance_type,
        description=description,
        location=location,
        start_time=start_dt,
        end_time=end_dt,
        duration_hours=duration_hours,
        recorded_by=user_name
    )

    db.add(new_maintenance)
    db.commit()

    return RedirectResponse(
        url=f"/journeys/ticket/{journey_id}", status_code=status.HTTP_303_SEE_OTHER
    )


@app.get("/journeys/ticket/maintenance/delete/{maintenance_id}")
def delete_journey_maintenance(
    maintenance_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    # البحث المباشر في جدول JourneyMaintenance المنفصل
    maint = (
        db.query(JourneyMaintenance)
        .filter(JourneyMaintenance.id == maintenance_id)
        .first()
    )
    if not maint:
        raise HTTPException(
            status_code=404, detail="Maintenance record not found or already deleted."
        )

    journey_id = maint.journey_id
    db.delete(maint)
    db.commit()

    return RedirectResponse(
        url=f"/journeys/ticket/{journey_id}", status_code=status.HTTP_303_SEE_OTHER
    )


@app.post("/journeys/ticket/maintenance/update/{maintenance_id}")
def update_journey_maintenance(
    maintenance_id: int,
    maintenance_type: str = Form(...),
    description: str = Form(...),
    location: str = Form(None),
    start_time: str = Form(...),
    end_time: str = Form(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    # البحث المباشر في جدول JourneyMaintenance المنفصل
    maint = db.query(JourneyMaintenance).filter(
        JourneyMaintenance.id == maintenance_id
    ).first()
    
    if not maint:
        raise HTTPException(status_code=404, detail="Maintenance record not found.")

    try:
        start_dt = datetime.fromisoformat(start_time) if start_time else maint.start_time
        end_dt = None
        duration_hours = 0.0
        
        if end_time and end_time.strip():
            end_dt = datetime.fromisoformat(end_time)
            if end_dt <= start_dt:
                raise HTTPException(status_code=400, detail="End time must be after start time.")
            duration_hours = round((end_dt - start_dt).total_seconds() / 3600.0, 2)
            
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid date and time format.")

    user_name = getattr(current_user, "username", None) or getattr(current_user, "email", None) or str(current_user)

    maint.maintenance_type = maintenance_type
    maint.description = description
    maint.location = location
    maint.start_time = start_dt
    maint.end_time = end_dt
    maint.duration_hours = duration_hours
    maint.recorded_by = user_name

    db.commit()

    return RedirectResponse(
        url=f"/journeys/ticket/{maint.journey_id}", status_code=status.HTTP_303_SEE_OTHER
    )


# --- 1. صفحة عرض سجلات الصيانة والأعطال المركزية (مع فلترة صلاحيات الشركات) ---
def format_hours_to_hm(hours):
    if not hours or hours <= 0:
        return "00:00"
    total_minutes = int(round(hours * 60))
    h = total_minutes // 60
    m = total_minutes % 60
    return f"{h:02d}:{m:02d}"


@app.get("/maintenance-logs", response_class=HTMLResponse)
def maintenance_logs_page(
    request: Request,
    search_plate: Optional[str] = None,
    m_type: Optional[str] = None,
    description: Optional[str] = None,
    location: Optional[str] = None,
    status: Optional[str] = None,       # تم إضافة بارامتر الحالة
    driver_id: Optional[str] = None,    
    segment_id: Optional[str] = None,   
    company_id: Optional[str] = None,   
    client_id: Optional[str] = None,    
    start_from: Optional[str] = None,
    start_to: Optional[str] = None,
    end_from: Optional[str] = None,
    end_to: Optional[str] = None,
    sort_by: Optional[str] = "id",
    order: Optional[str] = "asc",
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user)
):
    error_message = None
    try:
        drv_id = int(driver_id) if driver_id and str(driver_id).isdigit() else None
        seg_id = int(segment_id) if segment_id and str(segment_id).isdigit() else None
        comp_id = int(company_id) if company_id and str(company_id).isdigit() else None
        cli_id = int(client_id) if client_id and str(client_id).isdigit() else None

        if not search_plate: search_plate = None
        if not m_type: m_type = None
        if not description: description = None
        if not location: location = None
        if not status: status = None
        if not start_from: start_from = None
        if not start_to: start_to = None
        if not end_from: end_from = None
        if not end_to: end_to = None

        allowed_company_ids = []
        allowed_client_ids = []
        if current_user:
            if hasattr(current_user, 'companies') and current_user.companies:
                for c in current_user.companies:
                    if c.id not in allowed_company_ids:
                        allowed_company_ids.append(c.id)
            if hasattr(current_user, 'clients') and current_user.clients:
                for cl in current_user.clients:
                    if cl.id not in allowed_client_ids:
                        allowed_client_ids.append(cl.id)
                    if hasattr(cl, 'company_id') and cl.company_id and cl.company_id not in allowed_company_ids:
                        allowed_company_ids.append(cl.company_id)

        base_query = db.query(JourneyMaintenance)\
            .join(Journey, JourneyMaintenance.journey_id == Journey.id)\
            .outerjoin(Vehicle, Journey.vehicle_id == Vehicle.id)\
            .outerjoin(Route, Journey.route_id == Route.id)\
            .outerjoin(Driver, Journey.driver_id == Driver.id)\
            .outerjoin(Segment, Journey.segment_id == Segment.id)
        
        filters = []
        if allowed_company_ids:
            filters.append(Vehicle.company_id.in_(allowed_company_ids))
            filters.append(Journey.company_id.in_(allowed_company_ids))
        if allowed_client_ids:
            filters.append(Route.client_company_id.in_(allowed_client_ids))
            if hasattr(Journey, 'client_id'):
                filters.append(Journey.client_id.in_(allowed_client_ids))
        
        if filters:
            base_query = base_query.filter(or_(*filters))
        else:
            base_query = base_query.filter(Journey.id == -1)

        companies_query = db.query(Company)
        if allowed_company_ids:
            companies_query = companies_query.filter(Company.id.in_(allowed_company_ids))
        else:
            companies_query = companies_query.filter(Company.id == -1)
        available_companies = companies_query.all()

        clients_query = db.query(Client)
        if allowed_client_ids:
            clients_query = clients_query.filter(Client.id.in_(allowed_client_ids))
        else:
            clients_query = clients_query.filter(Client.id == -1)
        available_clients = clients_query.all()

        available_descriptions = [row[0] for row in base_query.with_entities(JourneyMaintenance.description).distinct().all() if row[0]]
        available_locations = [row[0] for row in base_query.with_entities(JourneyMaintenance.location).distinct().all() if row[0]]
        
        drivers_query = db.query(Driver)\
            .join(Journey, Journey.driver_id == Driver.id)\
            .join(Vehicle, Journey.vehicle_id == Vehicle.id)\
            .outerjoin(Route, Journey.route_id == Route.id)
            
        segments_query = db.query(Segment)\
            .join(Journey, Journey.segment_id == Segment.id)\
            .join(Vehicle, Journey.vehicle_id == Vehicle.id)\
            .outerjoin(Route, Journey.route_id == Route.id)

        dropdown_filters = []
        if allowed_company_ids:
            dropdown_filters.append(Vehicle.company_id.in_(allowed_company_ids))
            dropdown_filters.append(Journey.company_id.in_(allowed_company_ids))
        if allowed_client_ids:
            dropdown_filters.append(Route.client_company_id.in_(allowed_client_ids))
            if hasattr(Journey, 'client_id'):
                dropdown_filters.append(Journey.client_id.in_(allowed_client_ids))

        if dropdown_filters:
            drivers_query = drivers_query.filter(or_(*dropdown_filters))
            segments_query = segments_query.filter(or_(*dropdown_filters))
        else:
            drivers_query = drivers_query.filter(Journey.id == -1)
            segments_query = segments_query.filter(Journey.id == -1)

        available_drivers = drivers_query.distinct().all()
        available_segments = segments_query.distinct().all()

        query = base_query

        if search_plate:
            query = query.filter(Vehicle.plate_number.ilike(f"%{search_plate}%"))
        if m_type:
            query = query.filter(JourneyMaintenance.maintenance_type == m_type)
        if description:
            query = query.filter(JourneyMaintenance.description == description)
        if location:
            query = query.filter(JourneyMaintenance.location == location)
        
        # ربط فلتر الحالة بقاعدة البيانات (Completed تعني وجود end_time، و Ongoing تعني عدم وجوده)
        if status == "Completed":
            query = query.filter(JourneyMaintenance.end_time.isnot(None))
        elif status == "Ongoing":
            query = query.filter(JourneyMaintenance.end_time.is_(None))
        
        if drv_id:
            query = query.filter(Journey.driver_id == drv_id)
        if seg_id:
            query = query.filter(Journey.segment_id == seg_id)
        
        if comp_id:
            query = query.filter(or_(Journey.company_id == comp_id, Vehicle.company_id == comp_id))

        if cli_id:
            client_filters = [Route.client_company_id == cli_id]
            if hasattr(Journey, 'client_id'):
                client_filters.append(Journey.client_id == cli_id)
            query = query.filter(or_(*client_filters))
        
        try:
            if start_from:
                dt_val = start_from.replace('T', ' ')[:19]
                query = query.filter(JourneyMaintenance.start_time >= datetime.strptime(dt_val, "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M"))
            if start_to:
                dt_val = start_to.replace('T', ' ')[:19]
                query = query.filter(JourneyMaintenance.start_time <= datetime.strptime(dt_val, "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M"))
            if end_from:
                dt_val = end_from.replace('T', ' ')[:19]
                query = query.filter(JourneyMaintenance.end_time >= datetime.strptime(dt_val, "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M"))
            if end_to:
                dt_val = end_to.replace('T', ' ')[:19]
                query = query.filter(JourneyMaintenance.end_time <= datetime.strptime(dt_val, "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M"))
        except Exception as dt_err:
            error_message = f"Invalid date format provided: {str(dt_err)}"

        sort_column_map = {
            "serial": Journey.serial,
            "plate": Vehicle.plate_number,
            "driver": Driver.name,
            "segment": Segment.name,
            "type": JourneyMaintenance.maintenance_type,
            "description": JourneyMaintenance.description,
            "location": JourneyMaintenance.location,
            "start_time": JourneyMaintenance.start_time,
            "end_time": JourneyMaintenance.end_time,
            "duration": JourneyMaintenance.duration_hours,
            "recorded_by": JourneyMaintenance.recorded_by,
        }
        
        sort_col = sort_column_map.get(sort_by, JourneyMaintenance.id)
        if order == "asc":
            query = query.order_by(sort_col.asc())
        else:
            query = query.order_by(sort_col.desc())

        logs = query.all()
        
        # تعيين الحالة باللغة الإنجليزية لكل سجل لعرضها في الـ Template
        for log in logs:
            log.status_text = "Completed" if log.end_time else "Ongoing"
            log.status_badge = "success" if log.end_time else "warning"

        total_hours = sum([log.duration_hours for log in logs if log.duration_hours])
        total_mins_sum = int(round(total_hours * 60))
        total_duration_formatted = f"{total_mins_sum // 60:02d}:{total_mins_sum % 60:02d}"

        return templates.TemplateResponse(
            request=request,
            name="maintenance_logs.html",
            context={
                "logs": logs,
                "total_duration": total_duration_formatted,
                "total_records": len(logs),
                "available_descriptions": available_descriptions,
                "available_locations": available_locations,
                "available_drivers": available_drivers,
                "available_segments": available_segments,
                "available_companies": available_companies,
                "available_clients": available_clients,
                "selected_company_id": comp_id,
                "selected_client_id": cli_id,
                "selected_driver_id": drv_id,
                "selected_segment_id": seg_id,
                "error_message": error_message,
                "request_params": request.query_params
            }
        )
    except Exception as e:
        print(f"Error in maintenance logs page: {str(e)}")
        return templates.TemplateResponse(
            request=request,
            name="maintenance_logs.html",
            context={
                "logs": [],
                "total_duration": "00:00",
                "total_records": 0,
                "available_descriptions": [],
                "available_locations": [],
                "available_drivers": [],
                "available_segments": [],
                "available_companies": [],
                "available_clients": [],
                "selected_company_id": None,
                "selected_client_id": None,
                "selected_driver_id": None,
                "selected_segment_id": None,
                "error_message": f"An error occurred: {str(e)}",
                "request_params": request.query_params
            }
        )


@app.get("/maintenance-logs/export")
def export_maintenance_logs(
    search_plate: Optional[str] = None,
    m_type: Optional[str] = None,
    description: Optional[str] = None,
    location: Optional[str] = None,
    status: Optional[str] = None,
    driver_id: Optional[str] = None,    # تم تعديلها إلى str لتجنب خطأ 422
    segment_id: Optional[str] = None,   # تم تعديلها إلى str لتجنب خطأ 422
    company_id: Optional[str] = None,   # تم تعديلها إلى str لتجنب خطأ 422
    client_id: Optional[str] = None,    # تم تعديلها إلى str لتجنب خطأ 422
    start_from: Optional[str] = None,
    start_to: Optional[str] = None,
    end_from: Optional[str] = None,
    end_to: Optional[str] = None,
    sort_by: Optional[str] = "id",
    order: Optional[str] = "asc",
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    # تحويل الـ IDs إلى أرقام صحيحة بأمان
    drv_id = int(driver_id) if driver_id and str(driver_id).isdigit() else None
    seg_id = int(segment_id) if segment_id and str(segment_id).isdigit() else None
    comp_id = int(company_id) if company_id and str(company_id).isdigit() else None
    cli_id = int(client_id) if client_id and str(client_id).isdigit() else None

    allowed_company_ids = []
    allowed_client_ids = []
    if current_user:
        if hasattr(current_user, "companies") and current_user.companies:
            for c in current_user.companies:
                if c.id not in allowed_company_ids:
                    allowed_company_ids.append(c.id)
        if hasattr(current_user, "clients") and current_user.clients:
            for cl in current_user.clients:
                if cl.id not in allowed_client_ids:
                    allowed_client_ids.append(cl.id)
                if hasattr(cl, "company_id") and cl.company_id and cl.company_id not in allowed_company_ids:
                    allowed_company_ids.append(cl.company_id)

    base_query = (
        db.query(JourneyMaintenance)
        .join(Journey, JourneyMaintenance.journey_id == Journey.id)
        .outerjoin(Vehicle, Journey.vehicle_id == Vehicle.id)
        .outerjoin(Route, Journey.route_id == Route.id)
        .outerjoin(Driver, Journey.driver_id == Driver.id)
        .outerjoin(Segment, Journey.segment_id == Segment.id)
    )

    filters = []
    if allowed_company_ids:
        filters.append(Vehicle.company_id.in_(allowed_company_ids))
        filters.append(Journey.company_id.in_(allowed_company_ids))
    if allowed_client_ids:
        filters.append(Route.client_company_id.in_(allowed_client_ids))
        if hasattr(Journey, 'client_id'):
            filters.append(Journey.client_id.in_(allowed_client_ids))
    
    if filters:
        base_query = base_query.filter(or_(*filters))
    else:
        base_query = base_query.filter(Journey.id == -1)

    query = base_query

    if search_plate:
        query = query.filter(Vehicle.plate_number.ilike(f"%{search_plate}%"))
    if m_type:
        query = query.filter(JourneyMaintenance.maintenance_type == m_type)
    if description:
        query = query.filter(JourneyMaintenance.description == description)
    if location:
        query = query.filter(JourneyMaintenance.location == location)
        
    if status == "Completed":
        query = query.filter(JourneyMaintenance.end_time.isnot(None))
    elif status == "Ongoing":
        query = query.filter(JourneyMaintenance.end_time.is_(None))

    if drv_id:
        query = query.filter(Journey.driver_id == drv_id)
    if seg_id:
        query = query.filter(Journey.segment_id == seg_id)
    
    if comp_id:
        query = query.filter(or_(Journey.company_id == comp_id, Vehicle.company_id == comp_id))
        
    if cli_id:
        client_filters = [Route.client_company_id == cli_id]
        if hasattr(Journey, 'client_id'):
            client_filters.append(Journey.client_id == cli_id)
        query = query.filter(or_(*client_filters))

    try:
        if start_from:
            dt_val = start_from.replace("T", " ")[:19]
            query = query.filter(
                JourneyMaintenance.start_time
                >= datetime.strptime(
                    dt_val,
                    "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M",
                )
            )
        if start_to:
            dt_val = start_to.replace("T", " ")[:19]
            query = query.filter(
                JourneyMaintenance.start_time
                <= datetime.strptime(
                    dt_val,
                    "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M",
                )
            )
        if end_from:
            dt_val = end_from.replace("T", " ")[:19]
            query = query.filter(
                JourneyMaintenance.end_time
                >= datetime.strptime(
                    dt_val,
                    "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M",
                )
            )
        if end_to:
            dt_val = end_to.replace("T", " ")[:19]
            query = query.filter(
                JourneyMaintenance.end_time
                <= datetime.strptime(
                    dt_val,
                    "%Y-%m-%d %H:%M:%S" if len(dt_val) > 16 else "%Y-%m-%d %H:%M",
                )
            )
    except Exception:
        pass

    sort_column_map = {
        "serial": Journey.serial,
        "plate": Vehicle.plate_number,
        "driver": Driver.name if hasattr(Driver, 'name') else Journey.driver_id,
        "segment": Segment.name if hasattr(Segment, 'name') else Journey.segment_id,
        "type": JourneyMaintenance.maintenance_type,
        "description": JourneyMaintenance.description,
        "location": JourneyMaintenance.location,
        "start_time": JourneyMaintenance.start_time,
        "end_time": JourneyMaintenance.end_time,
        "duration": JourneyMaintenance.duration_hours,
        "recorded_by": JourneyMaintenance.recorded_by,
    }

    sort_col = sort_column_map.get(sort_by, JourneyMaintenance.id)
    if order == "asc":
        query = query.order_by(sort_col.asc())
    else:
        query = query.order_by(sort_col.desc())

    logs = query.all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Maintenance Logs"

    ws.sheet_view.rightToLeft = True
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    total_font = Font(name="Segoe UI", size=11, bold=True)
    total_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    thick_top_border = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

    ws.merge_cells("A1:M1")
    title_cell = ws["A1"]
    title_cell.value = "Journey Maintenance & Breakdowns Report"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:M2")
    info_cell = ws["A2"]
    info_cell.value = f"Exported At: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    info_cell.font = Font(name="Segoe UI", size=9, italic=True, color="595959")
    info_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = [
        "#",
        "Journey Serial",
        "Vehicle Plate",
        "Driver",
        "Segment",
        "Type",
        "Description",
        "Location",
        "Start Time",
        "End Time",
        "Duration (HH:MM)",
        "Status",
        "Recorded By",
    ]

    ws.row_dimensions[4].height = 26
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    total_hours_sum = 0.0
    row_idx = 5
    for index, log in enumerate(logs, start=1):
        ws.row_dimensions[row_idx].height = 22

        duration_hrs = log.duration_hours or 0.0
        total_hours_sum += duration_hrs

        total_mins = int(round(duration_hrs * 60))
        dur_formatted = f"{total_mins // 60:02d}:{total_mins % 60:02d}"

        driver_name_val = "N/A"
        if log.journey and log.journey.driver:
            driver_name_val = getattr(log.journey.driver, 'name', str(log.journey.driver_id))

        segment_name_val = "N/A"
        if log.journey and log.journey.segment:
            segment_name_val = getattr(log.journey.segment, 'name', str(log.journey.segment_id))

        status_val = "Completed" if log.end_time else "Ongoing"

        row_data = [
            index,
            log.journey.serial if log.journey and log.journey.serial else f"#{log.journey_id}",
            log.journey.vehicle.plate_number if log.journey and log.journey.vehicle else "N/A",
            driver_name_val,
            segment_name_val,
            log.maintenance_type or "N/A",
            log.description or "N/A",
            log.location or "N/A",
            log.start_time.strftime("%Y-%m-%d %H:%M") if log.start_time else "N/A",
            log.end_time.strftime("%Y-%m-%d %H:%M") if log.end_time else "N/A",
            dur_formatted,
            status_val,
            log.recorded_by or "N/A",
        ]

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 9, 10, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(
                    horizontal="right" if col_idx > 1 else "center", vertical="center"
                )

        row_idx += 1

    ws.row_dimensions[row_idx].height = 24

    total_mins_sum = int(round(total_hours_sum * 60))
    total_duration_formatted = (
        f"{total_mins_sum // 60:02d}:{total_mins_sum % 60:02d}"
    )

    ws.merge_cells(
        start_row=row_idx, start_column=1, end_row=row_idx, end_column=10
    )
    label_cell = ws.cell(row=row_idx, column=1, value="Total Duration")
    label_cell.font = total_font
    label_cell.alignment = Alignment(horizontal="left", vertical="center")

    val_cell = ws.cell(row=row_idx, column=11, value=total_duration_formatted)
    val_cell.font = total_font
    val_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 14):
        c = ws.cell(row=row_idx, column=col_idx)
        c.fill = total_fill
        c.border = thick_top_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2]:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 6, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    headers_resp = {
        "Content-Disposition": 'attachment; filename="maintenance_logs.xlsx"'
    }
    return StreamingResponse(
        output,
        headers=headers_resp,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

@app.post("/journeys/ticket/{journey_id}/followup/add")
async def add_journey_followup(
    journey_id: int,
    request: Request,
    comment: str = Form(...),
    location: str = Form(...),
    db: Session = Depends(get_db)
):
    current_user = request.cookies.get("user_email") or "System"

    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey:
        raise HTTPException(status_code=404, detail="Journey not found")

    new_followup = JourneyFollowUp(
        journey_id=journey_id,
        employee_name=current_user,
        comment=comment,
        location=location
    )

    db.add(new_followup)
    db.commit()
    db.refresh(new_followup)

    return {"success": True, "message": "Follow-up added successfully"}

@app.get("/journeys/ticket/followup/delete/{followup_id}")
def delete_journey_followup(request: Request, followup_id: int, db: Session = Depends(get_db)):
    try:
        followup = db.query(JourneyFollowUp).filter(JourneyFollowUp.id == followup_id).first()
        if not followup:
            raise HTTPException(status_code=404, detail="Follow-up not found")

        logged_in_user = get_current_user(request, db)
        deleted_by_name = (
            logged_in_user.email
            if logged_in_user and hasattr(logged_in_user, "email")
            else "Admin"
        )

        journey_id = followup.journey_id
        followup.is_deleted = True 
        followup.deleted_by = deleted_by_name  
        db.commit()
        
        timestamp = int(time.time())
        return RedirectResponse(url=f"/journeys/ticket/{journey_id}?t={timestamp}", status_code=303)
        
    except Exception as e:
        print(f"ERROR IN DELETE: {e}")  
        raise e


@app.get("/journeys/ticket/followup/restore/{followup_id}")
def restore_journey_followup(followup_id: int, db: Session = Depends(get_db)):
    followup = db.query(JourneyFollowUp).filter(JourneyFollowUp.id == followup_id).first()
    if not followup:
        raise HTTPException(status_code=404, detail="Follow-up not found")

    journey_id = followup.journey_id
    followup.is_deleted = False 
    followup.deleted_by = None  
    db.commit()
    
    timestamp = int(time.time())
    return RedirectResponse(url=f"/journeys/ticket/{journey_id}?t={timestamp}", status_code=303)


@app.get("/journeys/ticket/{journey_id}/followup/deleted-list")
def get_deleted_followups(journey_id: int, db: Session = Depends(get_db)):
    try:
        deleted_followups = db.query(JourneyFollowUp).filter(
            JourneyFollowUp.journey_id == journey_id,
            JourneyFollowUp.is_deleted == True
        ).all()
        
        result = []
        for f in deleted_followups:
            result.append({
                "id": f.id,
                "employee_name": f.employee_name,
                "location": f.location or "",
                "comment": f.comment,
                "deleted_by": f.deleted_by or "Unknown",  
                "created_at": f.created_at.strftime('%Y-%m-%d %H:%M') if f.created_at else ''
            })
        return result
    except Exception as e:
        print(f"ERROR FETCHING DELETED FOLLOWUPS: {e}")
        raise HTTPException(status_code=500, detail="Internal server error")
    
@app.get("/journeys/copy/{journey_id}")
def copy_journey_ticket(journey_id: int, db: Session = Depends(get_db)):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey or journey.status not in ["Ended", "Cancelled"]:
        raise HTTPException(status_code=400, detail="لا يمكن نسخ رحلة غير منتهية أو ملغاة.")
    
    copy_url = (
        f"/journeys/new?company_id={journey.company_id}"
        f"&client_company_id={journey.client_company_id or ''}"
        f"&segment_id={journey.segment_id or ''}"
        f"&vehicle_id={journey.vehicle_id or ''}"
        f"&driver_id={journey.driver_id or ''}"
        f"&route_id={journey.route_id or ''}"
        f"&load_type_id={journey.load_type_id or ''}"
        f"&trailer_id={journey.trailer_id or ''}"
    )
    
    return RedirectResponse(url=copy_url, status_code=303)

@app.get("/api/vehicles/search")
def search_vehicles(q: str = "", db: Session = Depends(get_db)):
    vehicles = db.query(Vehicle).filter(Vehicle.plate_number.ilike(f"%{q}%")).all()
    return [{"id": v.id, "plate_number": v.plate_number, "license": v.license_expiry} for v in vehicles]

@app.get("/api/trailers/search")
def search_trailers(q: str, db: Session = Depends(get_db)):
    trailers = db.query(Trailer).filter(Trailer.trailer_number.ilike(f"%{q}%")).all()
    return [{"id": t.id, "trailer_number": t.trailer_number, "license": t.license_expiry} for t in trailers]

@app.get("/api/drivers/search")
def search_drivers(q: str = "", db: Session = Depends(get_db)):
    drivers = db.query(Driver).filter(Driver.name.ilike(f"%{q}%")).all()
    return [{"id": d.id, "name": d.name, "license": d.license_expiry} for d in drivers]

@app.get("/api/routes/search")
def search_routes(q: str = "", db: Session = Depends(get_db)):
    routes = db.query(Route).filter(
        or_(
            Route.origin.ilike(f"%{q}%"), 
            Route.destination.ilike(f"%{q}%")
        )
    ).all()
    return [{"id": r.id, "origin": r.origin, "destination": r.destination, "distance": r.distance_km} for r in routes]

def get_journey_context(journey_id: int, db: Session):
    journey = db.query(Journey).filter(Journey.id == journey_id).first()
    if not journey:
        return None
        
    route = db.query(Route).filter(Route.id == journey.route_id).first() if journey.route_id else journey.route
    
    distance_km = route.distance_km if (route and hasattr(route, 'distance_km') and route.distance_km) else 0.0
    
    setting = db.query(Setting).filter(
        Setting.company_id == None,
        Setting.client_company_id == None
    ).first()
    speed = float(setting.speed) if setting and hasattr(setting, 'speed') and setting.speed else 40.0
    rest_duration = float(setting.rest_time) if setting and hasattr(setting, 'rest_time') and setting.rest_time else 0.5
    sleep_hours_setting = float(setting.sleep_time) if setting and hasattr(setting, 'sleep_time') and setting.sleep_time else 8.0

    driving_hours_val = getattr(route, 'driving_hours', None) if route else None
    if not driving_hours_val and route and route.distance_km:
        driving_hours_val = route.distance_km / speed if speed > 0 else 0.0
    driving_hours_val = driving_hours_val or 0.0

    rest_time_val = getattr(route, 'rest_time', None) if route else None
    if rest_time_val is None and route and route.distance_km:
        rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
        rest_time_val = rest_stops * rest_duration
    rest_time_val = rest_time_val or 0.0

    total_time_val = getattr(route, 'total_time', None) if route else None
    if not total_time_val:
        total_time_val = driving_hours_val + rest_time_val

    driving_hours = f"{driving_hours_val:.1f} hrs"
    rest_time = f"{rest_time_val:.1f} hrs"
    total_time = f"{total_time_val:.1f} hrs"

    def format_dt(val, fmt='%Y-%m-%d %H:%M'):
        if not val:
            return 'N/A'
        if isinstance(val, str):
            return val[:16].replace('T', ' ')
        try:
            return val.strftime(fmt)
        except AttributeError:
            return str(val)

    def format_date_only(val):
        return format_dt(val, fmt='%Y-%m-%d')

    total_planned_driving = total_time_val
    if getattr(journey, 'include_sleep', 'N') == "Y":
        total_planned_driving += sleep_hours_setting

    if journey.manual_start_time and isinstance(journey.manual_start_time, datetime):
        extra_off_duty = 24.0 if getattr(journey, 'is_off_duty', 'N') == "Y" else 0.0
        journey.planned_end_time = journey.manual_start_time + timedelta(hours=total_planned_driving + extra_off_duty)
    else:
        journey.planned_end_time = None

    if journey.manual_start_time and journey.manual_arrived_time and isinstance(journey.manual_start_time, datetime) and isinstance(journey.manual_arrived_time, datetime):
        diff_actual = journey.manual_arrived_time - journey.manual_start_time
        actual_tot_hours = diff_actual.total_seconds() / 3600
        if getattr(journey, 'is_off_duty', 'N') != "Y" and getattr(journey, 'include_sleep', 'N') == "Y":
            actual_tot_hours -= sleep_hours_setting
            
        total_maintenance_hours = sum(m.duration_hours for m in journey.maintenances if not m.is_deleted and m.duration_hours)
        actual_tot_hours -= total_maintenance_hours

        if actual_tot_hours < 0:
            actual_tot_hours = 0.0
        h_act = int(actual_tot_hours)
        m_act = int((actual_tot_hours - h_act) * 60)
        journey.actual_driving_hours = f"{h_act:02d}:{m_act:02d}"
    else:
        journey.actual_driving_hours = "00:00"

    ref_end_time = journey.manual_end_time if getattr(journey, 'is_off_duty', 'N') == "Y" else journey.manual_arrived_time

    if ref_end_time and journey.planned_end_time and isinstance(ref_end_time, datetime) and isinstance(journey.planned_end_time, datetime):
        diff_seconds = (ref_end_time - journey.planned_end_time).total_seconds()
        
        if abs(diff_seconds) < 60:
            journey.delay_early = "00:00"
        else:
            prefix = "-" if diff_seconds < 0 else "+"
            total_mins = int(abs(diff_seconds) // 60)
            hours = total_mins // 60
            minutes = total_mins % 60
            journey.delay_early = f"{prefix}{hours:02d}:{minutes:02d}"
    else:
        journey.delay_early = "00:00"

    if journey.manual_start_time and journey.manual_end_time and isinstance(journey.manual_start_time, datetime) and isinstance(journey.manual_end_time, datetime):
        diff_total = journey.manual_end_time - journey.manual_start_time
        tot_seconds = diff_total.total_seconds()
        if tot_seconds < 0:
            tot_seconds = 0
        tot_h = int(tot_seconds // 3600)
        tot_m = int((tot_seconds % 3600) // 60)
        journey.total_journey_duration = f"{tot_h:02d}:{tot_m:02d}"
    else:
        journey.total_journey_duration = "00:00"

    contact_info = {
        "driver_phone": journey.driver.phone if (journey.driver and hasattr(journey.driver, 'phone') and journey.driver.phone) else 'N/A',
        "vehicle_phone": journey.vehicle.phone_number if (journey.vehicle and hasattr(journey.vehicle, 'phone_number') and journey.vehicle.phone_number) else (getattr(journey.vehicle, 'phone', 'N/A') if journey.vehicle else 'N/A')
    }

    licenses_info = {
        "driver_license": format_date_only(journey.driver.license_expiry) if journey.driver else 'N/A',
        "vehicle_license": format_date_only(journey.vehicle.license_expiry) if journey.vehicle else 'N/A',
        "trailer_license": format_date_only(journey.trailer.license_expiry) if journey.trailer else 'N/A'
    }

    followups = []
    if journey.follow_ups:
        for idx, fu in enumerate(journey.follow_ups, start=1):
            if not fu.is_deleted:
                fu.display_index = idx
                fu.formatted_created_at = format_dt(fu.created_at)
                followups.append(fu)

    # تجهيز التوقفات التشغيلية غير المحذوفة للطباعة
    suspended_stops = []
    if hasattr(journey, 'suspended_stops') and journey.suspended_stops:
        for idx, stop in enumerate(journey.suspended_stops, start=1):
            if not getattr(stop, 'is_deleted', False):
                stop.display_index = idx
                stop.formatted_start = format_dt(getattr(stop, 'start_time', None))
                stop.formatted_end = format_dt(getattr(stop, 'end_time', None))
                suspended_stops.append(stop)

    # تجهيز سجلات الصيانة غير المحذوفة للطباعة
    maintenances = []
    if hasattr(journey, 'maintenances') and journey.maintenances:
        for idx, maint in enumerate(journey.maintenances, start=1):
            if not getattr(maint, 'is_deleted', False):
                maint.display_index = idx
                maint.formatted_date = format_dt(getattr(maint, 'maintenance_date', None))
                maintenances.append(maint)

    journey.formatted_start_time = format_dt(journey.manual_start_time)
    journey.formatted_planned_end = format_dt(getattr(journey, 'planned_end_time', None))
    journey.formatted_arrived_time = format_dt(journey.manual_arrived_time)
    journey.formatted_end_time = format_dt(journey.manual_end_time)

    return {
        "journey": journey,
        "distance_km": distance_km,
        "driving_hours": driving_hours,
        "rest_time": rest_time,
        "total_time": total_time,
        "contact_info": contact_info,
        "licenses_info": licenses_info,
        "followups": followups,
        "suspended_stops": suspended_stops,
        "maintenances": maintenances
    }

@app.get("/journeys/print/{journey_id}", response_class=HTMLResponse)
def print_journey_ticket(journey_id: int, request: Request, db: Session = Depends(get_db)):
    context_data = get_journey_context(journey_id, db)
    if not context_data:
        raise HTTPException(status_code=404, detail="Journey not found")
    
    return templates.TemplateResponse(request, "print_ticket.html", context_data)


@app.get("/journeys/off-duty-calendar", response_class=HTMLResponse)
def off_duty_calendar(
    request: Request, 
    start_date: str = None, 
    view_mode: str = "week", 
    db: Session = Depends(get_db)
):
    try:
        # 1. التحقق من المستخدم الحالي وصلاحياته
        current_user = get_current_user(request, db)
        if not current_user:
            return RedirectResponse(url="/login", status_code=303)

        allowed_company_ids = []
        if hasattr(current_user, 'client_company_id') and current_user.client_company_id:
            allowed_company_ids.append(current_user.client_company_id)
            
        for table_name in ["user_companies", "user_company", "user_transport_companies"]:
            try:
                result = db.execute(
                    text(f"SELECT company_id FROM {table_name} WHERE user_id = :uid"),
                    {"uid": current_user.id}
                ).fetchall()
                for row in result:
                    if row[0] not in allowed_company_ids:
                        allowed_company_ids.append(row[0])
            except Exception:
                db.rollback()

        # جلب العملاء المرتبطين بالمستخدم الحالي (Client Filtering)
        allowed_client_ids = []
        if hasattr(current_user, 'client_id') and current_user.client_id:
            allowed_client_ids.append(current_user.client_id)
            
        for table_name in ["user_clients", "user_client"]:
            try:
                result = db.execute(
                    text(f"SELECT client_id FROM {table_name} WHERE user_id = :uid"),
                    {"uid": current_user.id}
                ).fetchall()
                for row in result:
                    if row[0] not in allowed_client_ids:
                        allowed_client_ids.append(row[0])
            except Exception:
                db.rollback()

        if start_date:
            current_start = date.fromisoformat(start_date)
        else:
            current_start = date.today()
            
        if view_mode == "day":
            date_range = [current_start]
            prev_date = (current_start - timedelta(days=1)).strftime('%Y-%m-%d')
            next_date = (current_start + timedelta(days=1)).strftime('%Y-%m-%d')
            first_day = current_start
            last_day = current_start
            
        elif view_mode == "month":
            year = current_start.year
            month = current_start.month
            last_day_of_month = calendar.monthrange(year, month)[1]
            
            first_day = date(year, month, 1)
            last_day = date(year, month, last_day_of_month)
            
            date_range = [first_day + timedelta(days=i) for i in range((last_day - first_day).days + 1)]
            
            if month == 1:
                prev_date = date(year - 1, 12, 1).strftime('%Y-%m-%d')
                next_date = date(year, 2, 1).strftime('%Y-%m-%d')
            elif month == 12:
                prev_date = date(year, 11, 1).strftime('%Y-%m-%d')
                next_date = date(year + 1, 1, 1).strftime('%Y-%m-%d')
            else:
                prev_date = date(year, month - 1, 1).strftime('%Y-%m-%d')
                next_date = date(year, month + 1, 1).strftime('%Y-%m-%d')
                
        else: 
            num_days = 7
            date_range = [current_start + timedelta(days=i) for i in range(num_days)]
            first_day = date_range[0]
            last_day = date_range[-1]
            prev_date = (current_start - timedelta(days=7)).strftime('%Y-%m-%d')
            next_date = (current_start + timedelta(days=7)).strftime('%Y-%m-%d')
            
        # 2. جلب الرحلات مع فلاتر الشركات والعملاء
        query = db.query(Journey)
        is_superuser = getattr(current_user, 'is_superuser', False)

        if not is_superuser:
            if allowed_company_ids:
                query = query.filter(Journey.company_id.in_(allowed_company_ids))
            if allowed_client_ids and hasattr(Journey, 'client_id'):
                query = query.filter(Journey.client_id.in_(allowed_client_ids))
                
            if not allowed_company_ids and not allowed_client_ids:
                query = query.filter(Journey.id == -1)
            
        raw_journeys = query.all()
        
        all_off_duty_pool = []
        for j in raw_journeys:
            val_od = str(getattr(j, 'is_off_duty', '')).strip().lower()
            is_od_active = val_od in ['y', 'yes', 'true', '1', 'نعم']
            status_val = str(getattr(j, 'status', '')).strip().lower()
            is_active = status_val == 'active'
            
            if is_od_active and is_active:
                all_off_duty_pool.append(j)
        
        calendar_dict = {d.strftime('%Y-%m-%d'): [] for d in date_range}
        
        setting = db.query(Setting).filter(
            Setting.company_id == None,
            Setting.client_company_id == None
        ).first()
        
        speed = float(setting.default_speed) if setting and setting.default_speed else 40.0
        rest_duration = float(setting.rest_time) if setting and setting.rest_time else 0.5
        sleep_hours_setting = float(setting.sleep_time) if setting and setting.sleep_time else 8.0

        added_per_day = {d_str: set() for d_str in calendar_dict}

        for journey in raw_journeys:
            if not journey.manual_start_time:
                continue
            
            val_od = str(getattr(journey, 'is_off_duty', '')).strip().lower()
            is_od_active = val_od in ['y', 'yes', 'true', '1', 'نعم']
            
            if not is_od_active:
                continue

            auto_date = (journey.manual_start_time + timedelta(days=1)).date()
            auto_date_str = auto_date.strftime('%Y-%m-%d')
            
            if auto_date_str in calendar_dict:
                if journey.id not in added_per_day[auto_date_str]:
                    calendar_dict[auto_date_str].append({
                        "assignment_id": None,
                        "id": journey.id,
                        "serial": journey.serial,
                        "plate_number": journey.vehicle.plate_number if journey.vehicle else 'N/A'
                    })
                    added_per_day[auto_date_str].add(journey.id)

        # جلب التعيينات اليدوية مع تطبيق نفس فلاتر الشركات والعملاء
        manual_assignments_query = db.query(OffDutyAssignment).join(Journey, OffDutyAssignment.journey_id == Journey.id)
        if not is_superuser:
            if allowed_company_ids:
                manual_assignments_query = manual_assignments_query.filter(Journey.company_id.in_(allowed_company_ids))
            if allowed_client_ids and hasattr(Journey, 'client_id'):
                manual_assignments_query = manual_assignments_query.filter(Journey.client_id.in_(allowed_client_ids))
            if not allowed_company_ids and not allowed_client_ids:
                manual_assignments_query = manual_assignments_query.filter(Journey.id == -1)
            
        manual_assignments = manual_assignments_query.all()
        
        for ma in manual_assignments:
            j = db.query(Journey).filter_by(id=ma.journey_id).first()
            ma_date_str = ma.duty_date.strftime('%Y-%m-%d') if isinstance(ma.duty_date, date) else str(ma.duty_date)
            if j and ma_date_str in calendar_dict:
                if j.id not in added_per_day[ma_date_str]:
                    calendar_dict[ma_date_str].append({
                        "assignment_id": ma.id,
                        "id": j.id,
                        "serial": j.serial,
                        "plate_number": j.vehicle.plate_number if j.vehicle else 'N/A'
                    })
                    added_per_day[ma_date_str].add(ma_date_str)

        journey_dates_map = {}
        for date_str, items in calendar_dict.items():
            duty_date = date.fromisoformat(date_str)
            if first_day <= duty_date <= last_day:
                for item in items:
                    jid = item['id']
                    if jid not in journey_dates_map:
                        journey_dates_map[jid] = []
                    journey_dates_map[jid].append(date_str)

        for jid in journey_dates_map:
            journey_dates_map[jid].sort()

        table_details = []
        total_delay_seconds = 0.0
        total_duration_seconds = 0.0

        for date_str, items in calendar_dict.items():
            duty_date = date.fromisoformat(date_str)
            if first_day <= duty_date <= last_day:
                for item in items:
                    journey = db.query(Journey).filter_by(id=item['id']).first()
                    if not journey:
                        continue

                    driver_name = journey.driver.name if journey.driver else 'N/A'
                    vehicle_plate = journey.vehicle.plate_number if journey.vehicle else 'N/A'
                    day_name = duty_date.strftime('%A')
                    
                    load_type = 'N/A'
                    if hasattr(journey, 'load_type') and journey.load_type:
                        load_type = getattr(journey.load_type, 'name', str(journey.load_type))
                    
                    company_name = 'N/A'
                    if hasattr(journey, 'company') and journey.company:
                        company_name = getattr(journey.company, 'name', 'N/A')
                    
                    segment_name = 'N/A'
                    if hasattr(journey, 'segment') and journey.segment:
                        segment_name = getattr(journey.segment, 'name', str(journey.segment))
                    elif journey.route and hasattr(journey.route, 'segment') and journey.route.segment:
                        seg_obj = journey.route.segment
                        segment_name = getattr(seg_obj, 'name', str(seg_obj))
                    
                    route_str = 'N/A'
                    if journey.route:
                        if hasattr(journey.route, 'full_name') and journey.route.full_name:
                            route_str = journey.route.full_name
                        else:
                            origin_name = journey.route.origin_location.name if hasattr(journey.route, 'origin_location') and journey.route.origin_location else getattr(journey.route, 'origin', '')
                            dest_name = journey.route.destination_location.name if hasattr(journey.route, 'destination_location') and journey.route.destination_location else getattr(journey.route, 'destination', '')
                            if origin_name or dest_name:
                                route_str = f"{origin_name} - {dest_name}"
                    
                    dep_time_str = journey.manual_start_time.strftime('%Y-%m-%d %H:%M') if journey.manual_start_time else 'N/A'
                    
                    route_obj = journey.route
                    driving_hours_val = getattr(route_obj, 'driving_hours', None) if route_obj else None
                    if not driving_hours_val and route_obj and route_obj.distance_km:
                        driving_hours_val = route_obj.distance_km / speed if speed > 0 else 0.0
                    driving_hours_val = driving_hours_val or 0.0

                    rest_time_val = getattr(route_obj, 'rest_time', None) if route_obj else None
                    if rest_time_val is None and route_obj and route_obj.distance_km:
                        rest_stops = int(driving_hours_val // 2) if driving_hours_val >= 2 else 0
                        rest_time_val = rest_stops * rest_duration
                    rest_time_val = rest_time_val or 0.0

                    total_time_val = getattr(route_obj, 'total_time', None) if route_obj else None
                    if not total_time_val:
                        total_time_val = driving_hours_val + rest_time_val

                    total_planned_driving = total_time_val
                    val_od = str(getattr(journey, 'is_off_duty', '')).strip().lower()
                    is_od_active = val_od in ['y', 'yes', 'true', '1', 'نعم']

                    if getattr(journey, 'include_sleep', 'N') == "Y":
                        total_planned_driving += sleep_hours_setting

                    planned_end_time = None
                    if journey.manual_start_time and isinstance(journey.manual_start_time, datetime):
                        extra_off_duty = 24.0 if is_od_active else 0.0
                        planned_end_time = journey.manual_start_time + timedelta(hours=total_planned_driving + extra_off_duty)
                    
                    planned_arr_str = planned_end_time.strftime('%Y-%m-%d %H:%M') if planned_end_time else 'N/A'

                    actual_arr_str = 'N/A'
                    if journey.manual_arrived_time and isinstance(journey.manual_arrived_time, datetime):
                        actual_arr_str = journey.manual_arrived_time.strftime('%Y-%m-%d %H:%M')

                    end_time_str = 'N/A'
                    if journey.manual_end_time and isinstance(journey.manual_end_time, datetime):
                        end_time_str = journey.manual_end_time.strftime('%Y-%m-%d %H:%M')

                    all_dates_for_journey = journey_dates_map.get(journey.id, [])
                    is_first_day_of_multiple = (len(all_dates_for_journey) > 1 and date_str == all_dates_for_journey[0])

                    if is_first_day_of_multiple:
                        delay_early_str = "لم يعد"
                        total_duration_str = "لم يعد"
                    else:
                        ref_end_time = journey.manual_end_time if is_od_active else journey.manual_arrived_time
                        delay_early_str = "00:00"
                        if ref_end_time and planned_end_time and isinstance(ref_end_time, datetime) and isinstance(planned_end_time, datetime):
                            diff_seconds = (ref_end_time - planned_end_time).total_seconds()
                            total_delay_seconds += diff_seconds
                            if abs(diff_seconds) >= 60:
                                prefix = " -" if diff_seconds < 0 else " +"
                                total_mins = int(abs(diff_seconds) // 60)
                                hours = total_mins // 60
                                minutes = total_mins % 60
                                delay_early_str = f"{prefix}{hours:02d}:{minutes:02d}"

                        total_duration_str = "00:00"
                        if journey.manual_start_time and journey.manual_end_time and isinstance(journey.manual_start_time, datetime) and isinstance(journey.manual_end_time, datetime):
                            diff_total = journey.manual_end_time - journey.manual_start_time
                            tot_seconds = diff_total.total_seconds()
                            if tot_seconds > 0:
                                total_duration_seconds += tot_seconds
                                tot_h = int(tot_seconds // 3600)
                                tot_m = int((tot_seconds % 3600) // 60)
                                total_duration_str = f"{tot_h:02d}:{tot_m:02d}"

                    table_details.append({
                        "assignment_id": item['assignment_id'],
                        "driver_name": driver_name,
                        "vehicle_plate": vehicle_plate,
                        "load_type": load_type,
                        "company": company_name,
                        "segment": segment_name,
                        "route_name": route_str,
                        "off_duty_day": day_name,
                        "departure_time": dep_time_str,
                        "actual_arrival_home": actual_arr_str,
                        "planned_arrival": planned_arr_str,
                        "end_time": end_time_str,
                        "delay_early": delay_early_str,
                        "total_duration": total_duration_str
                    })

        def format_seconds_to_hm(sec, is_signed=False):
            if sec == 0:
                return "00:00"
            prefix = ""
            if sec < 0:
                prefix = " -" if is_signed else "-"
            elif is_signed and sec > 0:
                prefix = " +"
            
            sec_abs = abs(sec)
            h = int(sec_abs // 3600)
            m = int((sec_abs % 3600) // 60)
            return f"{prefix}{h:02d}:{m:02d}"

        table_totals = {
            "delay_first": format_seconds_to_hm(total_delay_seconds, is_signed=True),
            "delay_early": format_seconds_to_hm(total_delay_seconds, is_signed=True),
            "total_duration": format_seconds_to_hm(total_duration_seconds, is_signed=False)
        }
        
        return templates.TemplateResponse(
            request=request,
            name="off_duty_calendar.html",
            context={
                "calendar_data": calendar_dict,
                "off_duty_journeys_table": table_details,
                "all_off_duty_pool": all_off_duty_pool,
                "table_totals": table_totals,
                "current_start": current_start.strftime('%Y-%m-%d'),
                "prev_date": prev_date,
                "next_date": next_date,
                "view_mode": view_mode
            }
        )
    except Exception as e:
        print("ERROR IN OFF DUTY CALENDAR:", str(e))
        raise e

@app.post("/journeys/off-duty-calendar/assign")
def assign_off_duty_journey(
    journey_id: int = Form(...),
    date_str: str = Form(...),
    db: Session = Depends(get_db)
):
    try:
        duty_date_obj = date.fromisoformat(date_str)
        existing = db.query(OffDutyAssignment).filter_by(journey_id=journey_id, duty_date=duty_date_obj).first()
        if not existing:
            new_assignment = OffDutyAssignment(
                journey_id=journey_id,
                duty_date=duty_date_obj
            )
            db.add(new_assignment)
            db.commit()
        return {"status": "success"}
    except Exception as e:
        db.rollback()
        print(f"ERROR IN ASSIGN OFF DUTY: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/journeys/off-duty-calendar/remove")
def remove_off_duty_assignment(
    assignment_id: int = Form(...),
    db: Session = Depends(get_db)
):
    try:
        assignment = db.query(OffDutyAssignment).filter(OffDutyAssignment.id == assignment_id).first()
        if not assignment:
            raise HTTPException(status_code=404, detail="Assignment not found")
        db.delete(assignment)
        db.commit()
        return {"status": "success"}
    except Exception as e:
        db.rollback()
        print(f"ERROR IN REMOVE OFF DUTY: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/app-management", response_class=HTMLResponse)
async def app_management_dashboard(request: Request, db: Session = Depends(get_db)):
    current_user = get_current_user(request, db)
    if not current_user:
        return RedirectResponse(url="/login", status_code=303)
        
    # قصر الصلاحية على المطور فقط وإزالة الأدمن
    is_developer = (
        getattr(current_user, "is_developer", False) or 
        str(getattr(current_user, "role", "")).lower() in ["developer", "dev"]
    )
    
    # إذا لم يكن مطوراً، يتم عرض صفحة الـ HTML المخصصة للخطأ 403
    if not is_developer:
        return templates.TemplateResponse(
            request=request, 
            name="403.html", 
            status_code=403
        )
        
    return templates.TemplateResponse(
        request=request, 
        name="app_management.html"
    )

@app.get("/about", response_class=HTMLResponse)
def about_page(request: Request):
    return templates.TemplateResponse(request, "about.html", {"request": request})

# --- مسارات التتبع الحي للأسطول (Fleet Live Tracking Endpoints) ---
from datetime import datetime

class LocationUpdateSchema(BaseModel):
    driver_id: int
    latitude: float
    longitude: float
    speed: float

@app.post("/api/driver/update-location/")
def update_driver_location(data: LocationUpdateSchema, db: Session = Depends(get_db)):
    """
    استقبال إحداثيات السائق وتحديثها أو حفظها في قاعدة البيانات
    """
    # التحقق هل السائق موجود أصلاً في جدول السائقين
    driver_exists = db.query(Driver).filter(Driver.id == data.driver_id).first()
    if not driver_exists:
        return {"status": "error", "message": "Driver not found in database"}

    log = db.query(DriverLocationLog).filter(DriverLocationLog.driver_id == data.driver_id).first()
    
    if log:
        log.latitude = data.latitude
        log.longitude = data.longitude
        log.speed = data.speed
        log.updated_at = datetime.utcnow()
    else:
        new_log = DriverLocationLog(
            driver_id=data.driver_id,
            latitude=data.latitude,
            longitude=data.longitude,
            speed=data.speed,
            updated_at=datetime.utcnow()
        )
        db.add(new_log)
    
    db.commit()
    return {"status": "success", "message": "Location updated successfully"}

@app.get("/fleet/live-map", response_class=HTMLResponse, tags=["Fleet Live Tracking"])
def render_fleet_live_map(request: Request):
    """
    عرض صفحة خريطة الأسطول العامة لتتبع جميع السائقين
    """
    template = templates.env.get_template("FleetLiveMap.html")
    return HTMLResponse(template.render({"request": request}))


@app.get("/api/drivers/live-locations", tags=["Fleet Live Tracking"])
def get_all_drivers_live_locations(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """
    جلب السائقين ومواقعهم اللحظية المقيدين حصراً بالعملاء المرتبطين بالمستخدم الحالي
    """
    is_admin_or_dev = (
        getattr(current_user, "is_admin", False) or 
        getattr(current_user, "is_developer", False) or 
        str(getattr(current_user, "role", "")).lower() in ["admin", "developer", "dev"]
    )

    if is_admin_or_dev:
        drivers = db.query(Driver).all()
    else:
        # جمع معرفات العملاء المربوطين بالمستخدم فقط
        allowed_client_ids = [c.id for c in getattr(current_user, "clients", [])]
        if hasattr(current_user, "client_company_id") and current_user.client_company_id:
            if current_user.client_company_id not in allowed_client_ids:
                allowed_client_ids.append(current_user.client_company_id)

        # فحص الجدول الوسيط للعملاء إن وجد
        for table_name in ["user_clients", "user_client"]:
            try:
                result = db.execute(
                    text(f"SELECT client_id FROM {table_name} WHERE user_id = :uid"), 
                    {"uid": current_user.id}
                ).fetchall()
                for row in result:
                    if row[0] not in allowed_client_ids:
                        allowed_client_ids.append(row[0])
            except Exception:
                db.rollback()

        # جلب السائقين التابعين لهؤلاء العملاء فقط
        if allowed_client_ids:
            drivers = db.query(Driver).filter(Driver.client_id.in_(allowed_client_ids)).all()
        else:
            drivers = []

    result = []
    
    for driver in drivers:
        # البحث عن أحدث موقع مسجل للسائق
        log = db.query(DriverLocationLog).filter(DriverLocationLog.driver_id == driver.id).first()
        
        if log:
            lat = log.latitude
            lng = log.longitude
            speed = log.speed or 0.0
            updated_at = log.updated_at.strftime("%Y-%m-%d %H:%M:%S") if log.updated_at else ""
        else:
            # إحداثيات افتراضية في القاهرة في حال لم يرسل المحاكي موقعاً بعد
            lat = 30.0444
            lng = 31.2357
            speed = 0.0
            updated_at = ""

        result.append({
            "driver_id": driver.id,
            "driver_name": getattr(driver, 'name', f"Driver #{driver.id}"),
            "latitude": lat,
            "longitude": lng,
            "speed": speed,
            "updated_at": updated_at
        })
        
    return result


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)