import datetime
from datetime import datetime, timezone
from database import Base
from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    Table,
    func,
)
from sqlalchemy.orm import relationship

# --- جداول الوسيطة لعلاقات Many-to-Many للمستخدمين ---
user_companies = Table(
    "user_companies",
    Base.metadata,
    Column("user_id", Integer, ForeignKey("users.id", ondelete="CASCADE"), primary_key=True),
    Column("company_id", Integer, ForeignKey("companies.id", ondelete="CASCADE"), primary_key=True)
)

user_clients = Table(
    "user_clients",
    Base.metadata,
    Column("user_id", Integer, ForeignKey("users.id", ondelete="CASCADE"), primary_key=True),
    Column("client_id", Integer, ForeignKey("clients.id", ondelete="CASCADE"), primary_key=True)
)


# --- 1. شركات النقل الرئيسية (Transport Companies) ---
class Company(Base):
    __tablename__ = "companies"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    
    clients = relationship(
        "Client", back_populates="company", cascade="all, delete-orphan"
    )
    vehicles = relationship("Vehicle", back_populates="company", cascade="all, delete-orphan")
    trailers = relationship("Trailer", back_populates="company", cascade="all, delete-orphan")
    drivers = relationship("Driver", back_populates="company", cascade="all, delete-orphan")
    routes = relationship("Route", back_populates="company", cascade="all, delete-orphan")
    locations = relationship("Location", back_populates="company", cascade="all, delete-orphan")
    load_types = relationship("LoadType", back_populates="company", cascade="all, delete-orphan")
    settings = relationship("Setting", back_populates="company", cascade="all, delete-orphan")
    journeys = relationship("Journey", foreign_keys="[Journey.company_id]", back_populates="company")
    fleet_assignments = relationship("FleetAssignment", back_populates="company", cascade="all, delete-orphan")
    driver_leaves = relationship("DriverLeave", back_populates="company", cascade="all, delete-orphan", overlaps="company_leaves")


# --- 2. الشركات العميلة (Clients) - جدول منفصل ---
class Client(Base):
    __tablename__ = "clients"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=False)

    company = relationship("Company", back_populates="clients")
    segments = relationship(
        "Segment", back_populates="client", cascade="all, delete-orphan"
    )
    vehicles = relationship("Vehicle", back_populates="client")
    trailers = relationship("Trailer", back_populates="client")
    drivers = relationship("Driver", back_populates="client")
    routes = relationship("Route", back_populates="client_company")
    locations = relationship("Location", back_populates="client", cascade="all, delete-orphan")
    load_types = relationship("LoadType", back_populates="client")
    settings = relationship("Setting", back_populates="client_company")
    journeys = relationship("Journey", foreign_keys="[Journey.client_company_id]", back_populates="client_company")
    fleet_assignments = relationship("FleetAssignment", back_populates="client", cascade="all, delete-orphan")


# --- 3. القطاعات / الأقسام (Segments) ---
class Segment(Base):
    __tablename__ = "segments"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    client_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=False)

    client = relationship("Client", back_populates="segments")
    vehicles = relationship("Vehicle", back_populates="segment")
    trailers = relationship("Trailer", back_populates="segment")
    drivers = relationship("Driver", back_populates="segment")
    journeys = relationship("Journey", back_populates="segment")
    fleet_assignments = relationship("FleetAssignment", back_populates="segment", cascade="all, delete-orphan")


# --- 4. السيارات (Vehicles) ---
class Vehicle(Base):
    __tablename__ = "vehicles"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    plate_number = Column(String, nullable=False)
    license_expiry = Column(String, nullable=True)
    phone_number = Column(String, nullable=True)
    status = Column(String, default="Active")
    
    vehicle_type = Column(String, nullable=True)
    
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    client_id = Column(Integer, ForeignKey("clients.id"), nullable=True)
    segment_id = Column(Integer, ForeignKey("segments.id"), nullable=True)

    company = relationship("Company", back_populates="vehicles")
    client = relationship("Client", back_populates="vehicles")
    segment = relationship("Segment", back_populates="vehicles")
    journeys = relationship("Journey", back_populates="vehicle")
    fleet_assignments = relationship("FleetAssignment", back_populates="vehicle", cascade="all, delete-orphan")


# --- 5. المقطورات (Trailers) ---
class Trailer(Base):
    __tablename__ = "trailers"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    trailer_number = Column(String, nullable=False)
    license_expiry = Column(String, nullable=True)
    status = Column(String, default="Active")

    company_id = Column(Integer, ForeignKey("companies.id"), nullable=True)
    client_id = Column(Integer, ForeignKey("clients.id"), nullable=True)
    segment_id = Column(Integer, ForeignKey("segments.id"), nullable=True)

    company = relationship("Company", back_populates="trailers")
    client = relationship("Client", back_populates="trailers")
    segment = relationship("Segment", back_populates="trailers")
    journeys = relationship("Journey", back_populates="trailer")
    fleet_assignments = relationship("FleetAssignment", back_populates="trailer", cascade="all, delete-orphan")


# --- 6. السائقين (Drivers) ---
class Driver(Base):
    __tablename__ = "drivers"
    
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    name = Column(String, nullable=False)
    username = Column(String, unique=True, index=True, nullable=True) # اسم المستخدم لتسجيل الدخول
    password = Column(String, nullable=True)                          # كلمة المرور لتسجيل الدخول
    phone = Column(String, nullable=True)
    license_expiry = Column(String, nullable=True)
    status = Column(String, default="Active")

    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=True)
    client_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)
    segment_id = Column(Integer, ForeignKey("segments.id", ondelete="CASCADE"), nullable=True)

    company = relationship("Company", back_populates="drivers")
    client = relationship("Client", back_populates="drivers")
    segment = relationship("Segment", back_populates="drivers")
    journeys = relationship("Journey", back_populates="driver")
    fleet_assignments = relationship("FleetAssignment", back_populates="driver", cascade="all, delete-orphan")
    leaves = relationship("DriverLeave", back_populates="driver", cascade="all, delete-orphan", overlaps="driver_leaves")
    location_log = relationship("DriverLocationLog", back_populates="driver", uselist=False, cascade="all, delete-orphan")


# --- 6 مكرر. جدول التتبع الحي للسائقين (Driver Location Logs) ---
class DriverLocationLog(Base):
    __tablename__ = "driver_location_logs"

    id = Column(Integer, primary_key=True, index=True)
    driver_id = Column(Integer, ForeignKey("drivers.id", ondelete="CASCADE"), nullable=False, unique=True)
    latitude = Column(Float, nullable=False)
    longitude = Column(Float, nullable=False)
    speed = Column(Float, default=0.0, nullable=True)
    updated_at = Column(DateTime, default=lambda: datetime.now(), onupdate=lambda: datetime.now())

    driver = relationship("Driver", back_populates="location_log")


# --- 7. جدول إجازات السائقين (Driver Leaves) ---
class DriverLeave(Base):
    __tablename__ = "driver_leaves"

    id = Column(Integer, primary_key=True, index=True)
    driver_id = Column(Integer, ForeignKey("drivers.id"), nullable=False)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    start_date = Column(String, nullable=False)
    end_date = Column(String, nullable=True)     
    days_count = Column(Integer, nullable=True)   
    notes = Column(String, nullable=True)
    last_client = Column(String, nullable=True)
    status = Column(String, default="Active")     

    driver = relationship("Driver", backref="driver_leaves", overlaps="leaves")  
    company = relationship("Company", backref="company_leaves", overlaps="driver_leaves") 


# --- 8. جدول الأماكن المحددة من الخريطة (Locations) ---
class Location(Base):
    __tablename__ = "locations"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    location_type = Column(String, default="main", nullable=True)  
    segment = Column(String, nullable=True)
    latitude = Column(Float, nullable=False)
    longitude = Column(Float, nullable=False)
    
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=False)
    client_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)

    company = relationship("Company", back_populates="locations")
    client = relationship("Client", back_populates="locations")
    
    routes_as_origin = relationship("Route", foreign_keys="[Route.origin_id]", back_populates="origin_location", cascade="all, delete-orphan")
    routes_as_destination = relationship("Route", foreign_keys="[Route.destination_id]", back_populates="destination_location", cascade="all, delete-orphan")


# --- 9. خطوط السير (Routes) ---
class Route(Base):
    __tablename__ = "routes"
    
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    
    route_type_name = Column(String, nullable=True)
    
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=True)
    client_company_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)

    origin_id = Column(Integer, ForeignKey("locations.id", ondelete="CASCADE"), nullable=False)
    destination_id = Column(Integer, ForeignKey("locations.id", ondelete="CASCADE"), nullable=False)
    
    distance_km = Column(Float, default=0.0)
    estimated_hours = Column(Float, default=0.0)

    company = relationship("Company", back_populates="routes")
    client_company = relationship("Client", back_populates="routes")
    origin_location = relationship("Location", foreign_keys=[origin_id], back_populates="routes_as_origin")
    destination_location = relationship("Location", foreign_keys=[destination_id], back_populates="routes_as_destination")
    
    via_points = relationship("RouteViaPoint", back_populates="route", cascade="all, delete-orphan", order_by="RouteViaPoint.sequence_order")
    
    journeys = relationship("Journey", back_populates="route", cascade="all, delete-orphan")

    @property
    def full_name(self):
        origin_name = self.origin_location.name if self.origin_location else ""
        dest_name = self.destination_location.name if self.destination_location else ""
        
        if self.route_type_name:
            return f"{origin_name} - {dest_name} ({self.route_type_name})"
        return f"{origin_name} - {dest_name}"


# --- 9 مكرر. جدول النقاط الوسيطة لخطوط السير (Route Via Points) ---
class RouteViaPoint(Base):
    __tablename__ = "route_via_points"
    
    id = Column(Integer, primary_key=True, index=True)
    route_id = Column(Integer, ForeignKey("routes.id", ondelete="CASCADE"), nullable=False)
    location_id = Column(Integer, ForeignKey("locations.id", ondelete="CASCADE"), nullable=False)
    sequence_order = Column(Integer, default=0, nullable=False)

    route = relationship("Route", back_populates="via_points")
    location = relationship("Location")


# --- 10. أنواع الأحمال (Load Types) ---
class LoadType(Base):
    __tablename__ = "load_types"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=True)
    client_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)

    company = relationship("Company", back_populates="load_types")
    client = relationship("Client", back_populates="load_types")
    journeys = relationship("Journey", back_populates="load_type")


# --- 11. المستخدمين والصلاحيات (Users) ---
class User(Base):
    __tablename__ = "users"
    
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True, nullable=False) 
    email = Column(String, unique=True, index=True, nullable=False) 
    password = Column(String, nullable=False)
    role = Column(String, default="User")
    permissions = Column(String, nullable=True)
    
    companies = relationship("Company", secondary=user_companies, backref="users")
    clients = relationship("Client", secondary=user_clients, backref="users")
    
    journeys_created = relationship("Journey", foreign_keys="[Journey.created_by_user_id]", back_populates="creator_user")
    follow_ups = relationship("JourneyFollowUp", foreign_keys="[JourneyFollowUp.user_id]", back_populates="user")


# --- 12. الإعدادات المخصصة (Settings) ---
class Setting(Base):
    __tablename__ = "settings"
    
    id = Column(Integer, primary_key=True, index=True)
    
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=True)
    client_company_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)
    
    default_speed = Column(Float, default=40.0)
    rest_time = Column(Float, default=0.5)
    sleep_time = Column(Float, default=8.0)
    max_driving_hours = Column(Float, default=10.0)
    follow_up_interval = Column(Float, default=2.0)

    company = relationship("Company", back_populates="settings")
    client_company = relationship("Client", back_populates="settings")


# --- 13. جدول التسكين اليدوي والثابت للأسطول (Fleet Assignments) ---
class FleetAssignment(Base):
    __tablename__ = "fleet_assignments"
    
    id = Column(Integer, primary_key=True, index=True)
    company_id = Column(Integer, ForeignKey("companies.id", ondelete="CASCADE"), nullable=True)
    client_id = Column(Integer, ForeignKey("clients.id", ondelete="CASCADE"), nullable=True)
    segment_id = Column(Integer, ForeignKey("segments.id", ondelete="CASCADE"), nullable=True)
    
    vehicle_id = Column(Integer, ForeignKey("vehicles.id", ondelete="CASCADE"), nullable=False)
    driver_id = Column(Integer, ForeignKey("drivers.id", ondelete="CASCADE"), nullable=False)
    trailer_id = Column(Integer, ForeignKey("trailers.id", ondelete="CASCADE"), nullable=True)
    
    updated_at = Column(DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    company = relationship("Company", back_populates="fleet_assignments")
    client = relationship("Client", back_populates="fleet_assignments")
    segment = relationship("Segment", back_populates="fleet_assignments")
    vehicle = relationship("Vehicle", back_populates="fleet_assignments")
    driver = relationship("Driver", back_populates="fleet_assignments")
    trailer = relationship("Trailer", back_populates="fleet_assignments")


# --- 14. الرحلات (Journeys) ---
class Journey(Base):
    __tablename__ = "journeys"
    id = Column(Integer, primary_key=True, index=True)
    serial = Column(String, unique=True, index=True, nullable=True)
    
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    client_company_id = Column(Integer, ForeignKey("clients.id"), nullable=True)
    
    segment_id = Column(Integer, ForeignKey("segments.id"), nullable=True)
    driver_id = Column(Integer, ForeignKey("drivers.id"), nullable=False)
    vehicle_id = Column(Integer, ForeignKey("vehicles.id"), nullable=False)
    trailer_id = Column(Integer, ForeignKey("trailers.id"), nullable=True)
    route_id = Column(Integer, ForeignKey("routes.id"), nullable=False)
    load_type_id = Column(Integer, ForeignKey("load_types.id"), nullable=False)

    created_by = Column(String, nullable=True)  
    created_by_user_id = Column(Integer, ForeignKey("users.id"), nullable=True)  
    
    status = Column(String, default="Planned")
    remarks = Column(String, nullable=True)

    include_sleep = Column(String, default="N")
    is_off_duty = Column(String, default="N")

    locked_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    lock_expires_at = Column(DateTime, nullable=True)

    departure_time = Column(DateTime, nullable=True)
    arrival_time = Column(DateTime, nullable=True)
    manual_start_time = Column(DateTime, nullable=True)
    planned_end_time = Column(DateTime, nullable=True)
    manual_arrived_time = Column(DateTime, nullable=True)
    actual_hours = Column(String, nullable=True)
    delayed_duration = Column(String, nullable=True)
    manual_end_time = Column(DateTime, nullable=True)
    total_journey_hours = Column(String, nullable=True)
    
    created_at = Column(DateTime, default=lambda: datetime.now())  

    company = relationship("Company", foreign_keys=[company_id], back_populates="journeys")
    client_company = relationship("Client", foreign_keys=[client_company_id], back_populates="journeys")
    segment = relationship("Segment", back_populates="journeys")
    driver = relationship("Driver", back_populates="journeys")
    vehicle = relationship("Vehicle", back_populates="journeys")
    trailer = relationship("Trailer", back_populates="journeys")
    route = relationship("Route", back_populates="journeys")
    load_type = relationship("LoadType", back_populates="journeys")
    
    locker = relationship("User", foreign_keys=[locked_by])
    creator_user = relationship("User", foreign_keys=[created_by_user_id], back_populates="journeys_created")

    follow_ups = relationship(
        "JourneyFollowUp", back_populates="journey", cascade="all, delete-orphan"
    )
    maintenances = relationship(
        "JourneyMaintenance", back_populates="journey", cascade="all, delete-orphan"
    )
    off_duty_assignments = relationship(
        "OffDutyAssignment", back_populates="journey", cascade="all, delete-orphan"
    )
    stops = relationship(
        "JourneyStop", back_populates="journey", cascade="all, delete-orphan"
    )


# --- 15. جدول متابعة الرحلات (Journey Follow-ups) ---
class JourneyFollowUp(Base):
    __tablename__ = "journey_follow_ups"
    id = Column(Integer, primary_key=True, index=True)
    journey_id = Column(Integer, ForeignKey("journeys.id"), nullable=False)
    
    employee_name = Column(String, nullable=False)  
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)  

    location = Column(String, nullable=True)
    comment = Column(String, nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now())  

    is_deleted = Column(Boolean, default=False)
    deleted_by = Column(String, nullable=True)

    journey = relationship("Journey", back_populates="follow_ups")
    user = relationship("User", foreign_keys=[user_id], back_populates="follow_ups")


# --- 16. جدول تعيينات الإجازات (Off Duty Assignments) ---
class OffDutyAssignment(Base):
    __tablename__ = "off_duty_assignments"
    id = Column(Integer, primary_key=True, index=True)
    journey_id = Column(Integer, ForeignKey("journeys.id"), nullable=False)
    duty_date = Column(String, nullable=False)

    journey = relationship("Journey", back_populates="off_duty_assignments")


class JourneyMaintenance(Base):
    __tablename__ = "journey_maintenances"
    
    id = Column(Integer, primary_key=True, index=True)
    journey_id = Column(
        Integer,
        ForeignKey("journeys.id", ondelete="CASCADE"),
        nullable=False,
    )

    driver_name = Column(String(255), nullable=True)
    segment = Column(String(255), nullable=True)

    maintenance_type = Column(String(50), nullable=False)
    description = Column(Text, nullable=False)
    location = Column(String(255), nullable=True)
    start_time = Column(DateTime, nullable=False)
    end_time = Column(DateTime, nullable=True)
    duration_hours = Column(Float, nullable=True, default=0.0) 
    recorded_by = Column(String(255), nullable=True)
    is_deleted = Column(Boolean, default=False)

    journey = relationship("Journey", back_populates="maintenances")


# --- 18. جدول توقفات الرحلة (Journey Stops) ---
class JourneyStop(Base):
    __tablename__ = "journey_stops"
    
    id = Column(Integer, primary_key=True, index=True)
    journey_id = Column(Integer, ForeignKey("journeys.id", ondelete="CASCADE"), nullable=False)
    location = Column(String, nullable=True)              
    reason = Column(String, nullable=True)               
    start_time = Column(DateTime, nullable=False)      
    end_time = Column(DateTime, nullable=True)         
    hours_count = Column(Float, nullable=True)         
    recorded_by = Column(String, nullable=False)       
    
    journey = relationship("Journey", back_populates="stops")