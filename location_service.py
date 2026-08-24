"""Small, cached OpenStreetMap lookup for valuation-report nearby facilities."""

import json
import math
import os
import urllib.parse
import urllib.request


OVERPASS_URL = os.getenv("OVERPASS_API_URL", "https://overpass-api.de/api/interpreter")
USER_AGENT = os.getenv(
    "MAP_LOOKUP_USER_AGENT",
    "SVAI-Saksham-Valuation/1.0 (contact: sakshamvaluer@yahoo.com)",
)


def _number(value):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _distance_km(lat1, lon1, lat2, lon2):
    radius = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def _element_position(element):
    center = element.get("center") or element
    return _number(center.get("lat")), _number(center.get("lon"))


def _name(tags, fallback):
    return (
        tags.get("name:en") or tags.get("name") or tags.get("official_name")
        or tags.get("ref") or fallback
    )


def _nearest(elements, latitude, longitude, predicate, fallback):
    choices = []
    for element in elements:
        tags = element.get("tags") or {}
        if not predicate(tags):
            continue
        lat, lon = _element_position(element)
        if lat is None or lon is None:
            continue
        distance = _distance_km(latitude, longitude, lat, lon)
        choices.append((distance, _name(tags, fallback), tags))
    return min(choices, default=None, key=lambda item: item[0])


def _display(item):
    if not item:
        return "Not available in map data", ""
    distance, name, _ = item
    return str(name), f"{distance:.1f} Km"


def nearby_facilities(latitude, longitude, timeout=25):
    """Return report-ready nearest-place fields using one end-user-triggered query."""
    latitude, longitude = _number(latitude), _number(longitude)
    if latitude is None or longitude is None:
        return {}
    radius = max(1000, int(os.getenv("NEARBY_LOOKUP_RADIUS_METERS", "20000")))
    station_radius = max(radius, int(os.getenv("NEARBY_STATION_RADIUS_METERS", "100000")))
    facility_radius = max(radius, int(os.getenv("NEARBY_FACILITY_RADIUS_METERS", "30000")))
    query = f"""
[out:json][timeout:20];
(
  nwr(around:{station_radius},{latitude},{longitude})[railway=station];
  nwr(around:{facility_radius},{latitude},{longitude})[amenity=hospital];
  nwr(around:{facility_radius},{latitude},{longitude})[amenity~"^(school|college)$"];
  way(around:{radius},{latitude},{longitude})[highway~"^(motorway|trunk|primary|secondary|tertiary)$"];
  nwr(around:{radius},{latitude},{longitude})[amenity~"^(police|fire_station|bank|marketplace|fuel|bus_station)$"];
);
out center tags;
""".strip()
    request = urllib.request.Request(
        OVERPASS_URL,
        data=urllib.parse.urlencode({"data": query}).encode("utf-8"),
        headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            elements = json.load(response).get("elements", [])
    except Exception:
        return {}

    station = _nearest(elements, latitude, longitude, lambda t: t.get("railway") == "station", "Railway Station")
    hospital = _nearest(elements, latitude, longitude, lambda t: t.get("amenity") == "hospital", "Hospital")
    road = _nearest(
        elements, latitude, longitude,
        lambda t: t.get("highway") in {"motorway", "trunk", "primary", "secondary", "tertiary"},
        "Major Road",
    )
    education = _nearest(
        elements, latitude, longitude,
        lambda t: t.get("amenity") in {"school", "college"},
        "School / College",
    )
    other = _nearest(
        elements, latitude, longitude,
        lambda t: t.get("amenity") in {"police", "fire_station", "bank", "marketplace", "fuel", "bus_station"},
        "Nearby Facility",
    )
    result = {"nearby_lookup_source": "OpenStreetMap contributors"}
    for prefix, item in (
        ("nearest_railway_station", station),
        ("nearest_hospital", hospital),
        ("nearest_major_road", road),
        ("nearest_school_college", education),
        ("other_nearby_facility", other),
    ):
        result[prefix], result[f"{prefix}_distance"] = _display(item)
    # Community dominance is a sensitive demographic judgement and cannot be
    # established reliably from coordinates or public map POIs.
    result["community_dominated_area"] = "Not assessed from map data"
    return result
