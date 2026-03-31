"""
hotel_config.py  —  RevPar MD Platform (AWS Version)
======================================================
Central registry for all hotels.

AWS SETUP:
  - S3_BUCKET: the name of your S3 bucket (revpar-md-data)
  - AWS_REGION: the region your bucket lives in (us-east-1)
  - Hotel data files live at: s3://revpar-md-data/<folder_name>/
  - Hotel photos live at:     s3://revpar-md-data/photos/<filename>.jpg
  - Auth files live at:       s3://revpar-md-data/auth/

  To add a hotel: add an entry to HOTELS and upload its data folder to S3.
  Set active=False for hotels not yet configured (shows as "Coming Soon").
"""

import os
import io
import boto3
from botocore.exceptions import ClientError

# ── AWS CONFIG ────────────────────────────────────────────────────────────────
S3_BUCKET  = os.environ.get("S3_BUCKET", "revpar-md-data")
AWS_REGION = os.environ.get("AWS_REGION", "us-east-1")

def _s3():
    return boto3.client("s3", region_name=AWS_REGION)

def _photo_url(filename: str) -> str:
    """Return the public-readable S3 URL for a hotel photo."""
    return f"https://{S3_BUCKET}.s3.{AWS_REGION}.amazonaws.com/photos/{filename}"

# ── HOTEL REGISTRY ────────────────────────────────────────────────────────────
HOTELS = [
    {
        "id":           "hampton_lino_lakes",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Lino Lakes, MN",
        "folder_name":  "Hampton_Lino_Lakes",
        "brand":        "Hilton",
        "total_rooms":  112,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       True,
        "photo_path":   _photo_url("hampton_lino_lakes.jpg"),
    },
    {
        "id":           "hampton_superior",
        "display_name": "Hampton Inn Superior",
        "subtitle":     "Superior, WI",
        "folder_name":  "Hampton_Superior",
        "brand":        "Hilton",
        "total_rooms":  83,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       True,
        "photo_path":   _photo_url("Hampton_Superior.jpg"),
    },
    {
        "id":           "hilton_checkers_la",
        "display_name": "Hilton Checkers LA",
        "subtitle":     "Los Angeles, CA",
        "folder_name":  "Hilton_Checkers_LA",
        "brand":        "Hilton",
        "total_rooms":  193,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       True,
        "photo_path":   _photo_url("Hilton_Checkers_LA.jpg"),
    },
    {
        "id":           "hampton_cherry_creek",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Denver Cherry Creek, CO",
        "folder_name":  "Hampton_Cherry_Creek",
        "brand":        "Hilton",
        "total_rooms":  132,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       True,
        "photo_path":   _photo_url("hampton_cherry_creek.jpg"),
    },
    {
        "id":           "holiday_inn_express_superior",
        "display_name": "Holiday Inn Express & Suites",
        "subtitle":     "Superior, WI",
        "folder_name":  "Holiday_Inn_Express_Superior",
        "brand":        "IHG",
        "total_rooms":  84,
        "brand_color":  "#003580",
        "accent_color": "#0055A5",
        "active":       True,
        "photo_path":   _photo_url("holiday_inn_express_superior.jpg"),
    },
    {
        "id":           "hotel_indigo_rochester",
        "display_name": "Hotel Indigo",
        "subtitle":     "Rochester, MN",
        "folder_name":  "Hotel_Indigo_Rochester",
        "brand":        "IHG",
        "total_rooms":  178,
        "brand_color":  "#003580",
        "accent_color": "#0055A5",
        "active":       True,
        "photo_path":   _photo_url("hotel_indigo_rochester.jpg"),
    },

    # ── COMING SOON — S3 folders not yet created ───────────────────────────────
    {
        "id":           "hampton_chandler",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Chandler, AZ",
        "folder_name":  "Hampton_Chandler",
        "brand":        "Hilton",
        "total_rooms":  153,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Hampton_Chandler.jpg"),
    },
    {
        "id":           "homewood_chandler",
        "display_name": "Homewood Suites",
        "subtitle":     "Chandler, AZ",
        "folder_name":  "Homewood_Chandler",
        "brand":        "Hilton",
        "total_rooms":  133,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Homewood_Chandler.jpg"),
    },
    {
        "id":           "homewood_sarasota",
        "display_name": "Homewood Suites",
        "subtitle":     "Sarasota, FL",
        "folder_name":  "Homewood_Sarasota",
        "brand":        "Hilton",
        "total_rooms":  100,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Homewood_Sarasota.jpg"),
    },
    {
        "id":           "hyatt_place_fort_myers",
        "display_name": "Hyatt Place",
        "subtitle":     "Fort Myers, FL",
        "folder_name":  "Hyatt_Place_Fort_Myers",
        "brand":        "Hyatt",
        "total_rooms":  148,
        "brand_color":  "#96231b",
        "accent_color": "#c0392b",
        "active":       False,
        "photo_path":   _photo_url("Hyatt_Place_Fort_Myers.jpg"),
    },
    {
        "id":           "hyatt_place_cape_canaveral",
        "display_name": "Hyatt Place",
        "subtitle":     "Cape Canaveral, FL",
        "folder_name":  "Hyatt_Place_Cape_Canaveral",
        "brand":        "Hyatt",
        "total_rooms":  150,
        "brand_color":  "#96231b",
        "accent_color": "#c0392b",
        "active":       False,
        "photo_path":   _photo_url("Hyatt_Place_Cape_Canaveral.jpg"),
    },
    {
        "id":           "hampton_onalaska",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Onalaska, WI",
        "folder_name":  "Hampton_Onalaska",
        "brand":        "Hilton",
        "total_rooms":  107,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Hampton_Onalaska.jpg"),
    },
    {
        "id":           "doubletree_tampa",
        "display_name": "DoubleTree by Hilton",
        "subtitle":     "Tampa, FL",
        "folder_name":  "DoubleTree_Tampa",
        "brand":        "Hilton",
        "total_rooms":  291,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Tampa_Doubletree.jpg"),
    },
    {
        "id":           "hampton_lakewood",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Lakewood, CO",
        "folder_name":  "Hampton_Lakewood",
        "brand":        "Hilton",
        "total_rooms":  179,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Hampton_Lakewood.jpg"),
    },
    {
        "id":           "tapestry_virginia_crossings",
        "display_name": "Tapestry Collection by Hilton",
        "subtitle":     "Virginia Crossings, VA",
        "folder_name":  "Tapestry_Virginia_Crossings",
        "brand":        "Hilton",
        "total_rooms":  183,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Tapestry_Virginia_Crossings.jpg"),
    },
    {
        "id":           "hampton_holly_springs",
        "display_name": "Hampton Inn & Suites",
        "subtitle":     "Holly Springs, NC",
        "folder_name":  "Hampton_Holly_Springs",
        "brand":        "Hilton",
        "total_rooms":  124,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Hampton_Holly_Springs.jpg"),
    },
    {
        "id":           "hgi_duncanville",
        "display_name": "Hilton Garden Inn",
        "subtitle":     "Duncanville, TX",
        "folder_name":  "HGI_Duncanville",
        "brand":        "Hilton",
        "total_rooms":  142,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("HGI_Duncanville.jpg"),
    },
    {
        "id":           "embassy_brooklyn_center",
        "display_name": "Embassy Suites",
        "subtitle":     "Brooklyn Center, MN",
        "folder_name":  "Embassy_Brooklyn_Center",
        "brand":        "Hilton",
        "total_rooms":  175,
        "brand_color":  "#003087",
        "accent_color": "#0066CC",
        "active":       False,
        "photo_path":   _photo_url("Embassy_Brooklyn_Center.jpg"),
    },
]

# ── FILE ROLE DETECTION ───────────────────────────────────────────────────────
FILE_ROLES = {
    "master":   [".xlsb", "revpar"],
    "str":      ["str.xlsx"],
    "budget":   ["budget.xlsx"],
    "year":     ["year.xlsx", "data glance", "data_glance"],
    "groups":   ["group_wash", "group wash", "groupwash", "group summary", "groups.xlsx"],
    "rates":    ["rates.xlsx"],
    "booking":  ["booking_reports", "srp_activity", "srp activity", "booking reports"],
    "srp_pace": ["srp pace", "srp_pace", "property_segment_data", "property segment data",
                 "dss property", "dss_property", "property detail", "property_detail"],
    "pickup1":  ["1.xlsx"],
    "pickup7":  ["7.xlsx"],
    "strategic":["strategicanalysis", "strategic_analysis"],
    "corp_segments": ["corp_segments", "corp_segment"],
    "lighthouse_events": ["lighthouse events", "lighthouse_events"],
}

FILE_LABELS = {
    "master":   "RevPar MD Master",
    "str":      "STR Report",
    "budget":   "Budget",
    "year":     "Year Data",
    "groups":   "Group Wash Report",
    "rates":    "Rates",
    "booking":  "SRP Activity",
    "srp_pace": "SRP Pace",
    "pickup1":  "Overnight Pickup",
    "pickup7":  "7-Day Pickup",
    "strategic":"Strategic Analysis",
    "corp_segments":"Corp Segments",
    "lighthouse_events":"Lighthouse Events",
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_hotel_folder(hotel: dict) -> str:
    """Return the S3 prefix (folder path) for a hotel's data files."""
    return hotel["folder_name"] + "/"


def list_s3_files(hotel: dict) -> list[str]:
    """Return a list of filenames (not full keys) inside the hotel's S3 folder."""
    prefix = get_hotel_folder(hotel)
    try:
        s3 = _s3()
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=prefix)
        objects = resp.get("Contents", [])
        # Strip the prefix to get just the filename
        return [
            obj["Key"][len(prefix):]
            for obj in objects
            if obj["Key"] != prefix and "/" not in obj["Key"][len(prefix):]
        ]
    except ClientError:
        return []


def detect_files(hotel: dict) -> dict:
    """Return {role: s3_key or None} for all file roles."""
    result   = {role: None for role in FILE_ROLES}
    prefix   = get_hotel_folder(hotel)
    filenames = list_s3_files(hotel)

    for role, patterns in FILE_ROLES.items():
        for fname in filenames:
            name_lower = fname.lower()
            for pat in patterns:
                if pat.startswith("."):
                    if name_lower.endswith(pat):
                        result[role] = prefix + fname
                        break
                else:
                    if pat in name_lower:
                        result[role] = prefix + fname
                        break
            if result[role]:
                break
    return result


def file_status(hotel: dict) -> dict:
    files     = detect_files(hotel)
    n_present = sum(1 for v in files.values() if v is not None)
    n_total   = len(FILE_ROLES)
    return {
        "files":   files,
        "present": n_present,
        "total":   n_total,
        "ready":   n_present == n_total,
        "pct":     int(100 * n_present / n_total),
    }


def download_file_bytes(s3_key: str) -> io.BytesIO | None:
    """
    Download a file from S3 and return it as a BytesIO object.
    Used by data_loader.py instead of open(path, 'rb').
    Returns None if the file does not exist or cannot be read.
    """
    try:
        s3  = _s3()
        obj = s3.get_object(Bucket=S3_BUCKET, Key=s3_key)
        return io.BytesIO(obj["Body"].read())
    except ClientError:
        return None


def read_json_from_s3(s3_key: str) -> dict | list | None:
    """Read and parse a JSON file from S3. Returns None on failure."""
    import json
    buf = download_file_bytes(s3_key)
    if buf is None:
        return None
    try:
        return json.loads(buf.read().decode("utf-8"))
    except Exception:
        return None


def write_json_to_s3(s3_key: str, data: dict | list) -> bool:
    """Serialize data to JSON and write it to S3. Returns True on success."""
    import json
    try:
        s3      = _s3()
        payload = json.dumps(data, indent=2).encode("utf-8")
        s3.put_object(Bucket=S3_BUCKET, Key=s3_key, Body=payload,
                      ContentType="application/json")
        return True
    except ClientError:
        return False


def read_yaml_from_s3(s3_key: str) -> dict | None:
    """Read and parse a YAML file from S3. Returns None on failure."""
    import yaml
    buf = download_file_bytes(s3_key)
    if buf is None:
        return None
    try:
        return yaml.safe_load(buf.read().decode("utf-8"))
    except Exception:
        return None


def write_yaml_to_s3(s3_key: str, data: dict) -> bool:
    """Serialize data to YAML and write it to S3. Returns True on success."""
    import yaml
    try:
        s3      = _s3()
        payload = yaml.dump(data, default_flow_style=False,
                            allow_unicode=True).encode("utf-8")
        s3.put_object(Bucket=S3_BUCKET, Key=s3_key, Body=payload,
                      ContentType="application/x-yaml")
        return True
    except ClientError:
        return False
