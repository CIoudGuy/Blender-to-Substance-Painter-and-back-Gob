bl_info = {
    "name": "GoB SP Bridge",
    "author": "Cloud Guy | cloud_was_taken on Discord",
    "version": (0, 2, 0),
    "blender": (4, 5, 0),
    "location": "View3D > Sidebar > GoB SP",
    "description": "Send FBX to Substance 3D Painter and import meshes/textures back",
    "category": "Import-Export",
}

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import urllib.error
import urllib.request
from pathlib import Path

import bpy
from bpy.props import BoolProperty, FloatProperty, StringProperty, PointerProperty
from bpy.types import AddonPreferences, Operator, Panel


BRIDGE_ENV_VAR = "GOB_SP_BRIDGE_DIR"
BRIDGE_ROOT_HINT_FILENAME = "bridge_root.json"
BRIDGE_SHARED_HINT_DIRNAME = ".gob_sp_bridge"
MANIFEST_FILENAME = "bridge.json"
BLENDER_EXPORT_FILENAME = "b2sp.fbx"
BLENDER_HIGH_FILENAME = "b2sp_hi.fbx"
SP_EXPORT_FILENAME = "sp2b.fbx"
ACTIVE_SP_INFO_FILENAME = "active_sp.json"
ACTIVE_BLENDER_INFO_FILENAME = "active_blender.json"
LINKS_FILENAME = "project_links.json"
PROJECT_META_DIRNAME = ".gob_meta"
TEMP_DIRNAME = ".gob_temp"
TEMP_SP_PREFIX = "gob_unsaved_sp_"
TEMP_BLENDER_PREFIX = "gob_unsaved_bl_"
TEMP_SP_SUFFIX = ".spp"
TEMP_BLENDER_SUFFIX = ".blend"
ACTIVE_SP_INFO_MAX_AGE = 120.0
ACTIVE_BLENDER_INFO_MAX_AGE = 120.0
UPDATE_URL = (
    "https://raw.githubusercontent.com/CIoudGuy/Blender-to-Substance-Painter-and-back-Gob/"
    "refs/heads/main/version.json"
)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".tga", ".exr"}
CACHE_WARN_BYTES = 35 * 1024 ** 3
DEFAULT_CACHE_LIMIT_GB = 35.0
UI_LINK_CACHE_TTL = 0.75
_temp_session_id = None
_temp_blender_file = None
_last_blender_file = None
_project_dir_cache = {}
_ui_link_cache = {
    "timestamp": 0.0,
    "blender_file": "",
    "project_dir": "",
    "active_info": None,
    "auto_sp_project": "",
    "linked_sp_project": "",
    "auto_is_temp": False,
    "auto_exists": False,
}

MAP_KEYWORDS = [
    ("basecolor", "base_color"),
    ("base_color", "base_color"),
    ("basec", "base_color"),
    ("basecol", "base_color"),
    ("basecolour", "base_color"),
    ("base_map", "base_color"),
    ("basemap", "base_color"),
    ("albedo", "base_color"),
    ("diffuse", "base_color"),
    ("materialparams", "orm"),
    ("materialparam", "orm"),
    ("metallic", "metallic"),
    ("metalness", "metallic"),
    ("roughness", "roughness"),
    ("glossiness", "glossiness"),
    ("smoothness", "glossiness"),
    ("specular", "specular"),
    ("reflection", "specular"),
    ("normal", "normal"),
    ("ambientocclusion", "ao"),
    ("occlusion", "ao"),
    ("opacity", "opacity"),
    ("alpha", "opacity"),
    ("transparent", "opacity"),
    ("transparency", "opacity"),
    ("cutout", "opacity"),
    ("emissive", "emission"),
    ("emission", "emission"),
    ("height", "height"),
    ("displacement", "height"),
    ("color", "base_color"),
    ("metal", "metallic"),
    ("rough", "roughness"),
    ("gloss", "glossiness"),
    ("ao", "ao"),
    ("disp", "height"),
    ("nrm", "normal"),
]

DISCORD_INVITE_URL = "https://discord.gg/BE7k9Xxm5z"
BUG_REPORT_URL = (
    "https://github.com/CIoudGuy/Blender-to-Substance-Painter-and-back-Gob/issues"
)


def windows_documents_dir():
    if os.name != "nt":
        return None
    try:
        import ctypes
        CSIDL_PERSONAL = 5
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(260)
        result = ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf
        )
        if result == 0 and buf.value:
            return buf.value
    except Exception:
        return None
    return None


def default_bridge_dir():
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        return env_path
    docs = windows_documents_dir()
    if docs:
        return str(Path(docs) / "GoB_SP_Bridge")
    return str(Path.home() / "Documents" / "GoB_SP_Bridge")


def documents_bridge_root():
    docs = windows_documents_dir()
    if docs:
        return Path(docs) / "GoB_SP_Bridge"
    if sys.platform == "darwin":
        icloud_docs = (
            Path.home()
            / "Library"
            / "Mobile Documents"
            / "com~apple~CloudDocs"
            / "Documents"
        )
        if icloud_docs.is_dir():
            return icloud_docs / "GoB_SP_Bridge"
    return Path.home() / "Documents" / "GoB_SP_Bridge"


def sanitize_name(name):
    if not name:
        return "untitled"
    safe = []
    for ch in name:
        if ord(ch) < 128 and (ch.isalnum() or ch in "-_"):
            safe.append(ch)
        else:
            safe.append("_")
    result = "".join(safe).strip("_")
    return result or "untitled"


def normalize_path(path):
    if not path:
        return ""
    try:
        return os.path.abspath(os.path.expanduser(str(path)))
    except Exception:
        return str(path)


def normalize_path_key(path):
    return normalize_path(path).lower()


def temp_session_id():
    global _temp_session_id
    if _temp_session_id is None:
        _temp_session_id = f"{os.getpid()}_{int(time.time())}"
    return _temp_session_id


def bridge_temp_dir(prefs=None):
    return get_bridge_root(prefs) / TEMP_DIRNAME


def ensure_placeholder_file(path):
    if not path:
        return
    try:
        path = Path(path)
    except Exception:
        return
    try:
        ensure_dir(path.parent)
        path.touch(exist_ok=True)
    except OSError:
        return


def temp_blender_file_path(prefs=None):
    global _temp_blender_file
    if _temp_blender_file:
        return _temp_blender_file
    temp_dir = bridge_temp_dir(prefs)
    filename = f"{TEMP_BLENDER_PREFIX}{temp_session_id()}{TEMP_BLENDER_SUFFIX}"
    temp_path = temp_dir / filename
    ensure_placeholder_file(temp_path)
    _temp_blender_file = str(temp_path)
    return _temp_blender_file


def get_blender_file_path_or_temp(prefs=None):
    if bpy.data.filepath:
        return bpy.data.filepath
    return temp_blender_file_path(prefs)


def is_temp_file(path, prefix, suffix, prefs=None):
    if not path:
        return False
    try:
        path_obj = Path(path)
    except Exception:
        return False
    name = path_obj.name.lower()
    if not (name.startswith(prefix) and name.endswith(suffix)):
        return False
    try:
        return normalize_path(path_obj.parent).lower() == normalize_path(bridge_temp_dir(prefs)).lower()
    except Exception:
        return False


def is_temp_blender_file(path, prefs=None):
    return is_temp_file(path, TEMP_BLENDER_PREFIX, TEMP_BLENDER_SUFFIX, prefs)


def is_temp_sp_project_file(path, prefs=None):
    return is_temp_file(path, TEMP_SP_PREFIX, TEMP_SP_SUFFIX, prefs)


def project_meta_dir(project_dir):
    return Path(project_dir) / PROJECT_META_DIRNAME


def project_manifest_path(project_dir):
    if not project_dir:
        return None
    return project_meta_dir(project_dir) / MANIFEST_FILENAME


def legacy_project_manifest_path(project_dir):
    if not project_dir:
        return None
    return Path(project_dir) / MANIFEST_FILENAME


def find_project_manifest_path(project_dir):
    if not project_dir:
        return None
    new_path = project_manifest_path(project_dir)
    if new_path and new_path.exists():
        return new_path
    legacy_path = legacy_project_manifest_path(project_dir)
    if legacy_path and legacy_path.exists():
        return legacy_path
    return new_path


def project_dir_from_manifest_path(manifest_path):
    if not manifest_path:
        return None
    path = Path(manifest_path)
    if path.parent.name == PROJECT_META_DIRNAME:
        return path.parent.parent
    return path.parent


def project_dir_cache_key(blender_file):
    if not blender_file:
        return ""
    return normalize_path_key(blender_file)


def cached_project_dir(blender_file):
    key = project_dir_cache_key(blender_file)
    if not key:
        return None
    cached = _project_dir_cache.get(key)
    return Path(cached) if cached else None


def set_cached_project_dir(blender_file, project_dir):
    key = project_dir_cache_key(blender_file)
    if not key or not project_dir:
        return
    _project_dir_cache[key] = str(project_dir)


def manifest_matches_blender_file(manifest, blender_file):
    if not manifest or not blender_file:
        return False
    manifest_bl = get_manifest_blender_file(manifest)
    return bool(manifest_bl and paths_match(manifest_bl, blender_file))


def resolve_project_dir_for_blender(context, prefs, blender_file):
    cached = cached_project_dir(blender_file)
    if cached:
        return cached
    if blender_file:
        linked_dir = project_dir_from_linked_sp(blender_file, prefs)
        if linked_dir:
            set_cached_project_dir(blender_file, linked_dir)
            return linked_dir
    base_dir = get_bridge_root(prefs) / get_project_name(context)
    if blender_file and base_dir.exists():
        manifest_path = find_project_manifest_path(base_dir)
        manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
        if manifest_matches_blender_file(manifest, blender_file):
            set_cached_project_dir(blender_file, base_dir)
            return base_dir
    if blender_file:
        manifest_path = find_manifest_for_blender_file(
            get_candidate_bridge_roots(prefs),
            blender_file,
        )
        if manifest_path:
            project_dir = project_dir_from_manifest_path(manifest_path)
            set_cached_project_dir(blender_file, project_dir)
            return project_dir
        project_dir = unique_project_dir(base_dir, blender_file, prefs)
        set_cached_project_dir(blender_file, project_dir)
        return project_dir
    return base_dir


def unique_project_dir(base_dir, blender_file, prefs):
    if not base_dir.exists():
        return base_dir
    if blender_file:
        manifest_path = find_project_manifest_path(base_dir)
        manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
        if manifest_matches_blender_file(manifest, blender_file):
            return base_dir
    root = base_dir.parent
    base_name = base_dir.name
    index = 1
    while True:
        candidate = root / f"{base_name}{index}"
        if candidate.exists():
            if blender_file:
                manifest_path = find_project_manifest_path(candidate)
                manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
                if manifest_matches_blender_file(manifest, blender_file):
                    return candidate
            index += 1
            continue
        return candidate


def project_dir_for_send(context, prefs, blender_file):
    if blender_file:
        cached = cached_project_dir(blender_file)
        if cached:
            return cached
        linked_dir = project_dir_from_linked_sp(blender_file, prefs)
        if linked_dir:
            set_cached_project_dir(blender_file, linked_dir)
            return linked_dir
        manifest_path = find_manifest_for_blender_file(
            get_candidate_bridge_roots(prefs),
            blender_file,
        )
        if manifest_path:
            project_dir = project_dir_from_manifest_path(manifest_path)
            set_cached_project_dir(blender_file, project_dir)
            return project_dir
    base_dir = get_bridge_root(prefs) / get_project_name(context)
    project_dir = unique_project_dir(base_dir, blender_file, prefs)
    if blender_file:
        set_cached_project_dir(blender_file, project_dir)
    return project_dir


def link_registry_paths(prefs=None):
    roots = []
    docs_root = documents_bridge_root()
    if docs_root:
        roots.append(Path(docs_root))
    for root in get_candidate_bridge_roots(prefs):
        if not root:
            continue
        try:
            root_path = Path(root)
        except TypeError:
            continue
        if root_path.exists():
            roots.append(root_path)
    unique = []
    seen = set()
    for root in roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(root)
    return [root / LINKS_FILENAME for root in unique]


def load_link_registry(prefs=None):
    for path in link_registry_paths(prefs):
        if not path or not path.exists():
            continue
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except (OSError, json.JSONDecodeError):
            continue
        if isinstance(data, dict):
            return data
    return {}


def save_link_registry(data, prefs=None):
    paths = link_registry_paths(prefs)
    if not paths:
        return
    primary = paths[0]
    ensure_dir(primary.parent)
    try:
        with open(primary, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, ensure_ascii=True)
    except OSError:
        return
    for path in paths[1:]:
        if not path.exists():
            continue
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(data, handle, indent=2, ensure_ascii=True)
        except OSError:
            continue


def update_link_registry(sp_project_file=None, blender_file=None, prefs=None):
    if not sp_project_file or not blender_file:
        return
    data = load_link_registry(prefs)
    sp_key = normalize_path_key(sp_project_file)
    bl_key = normalize_path_key(blender_file)
    sp_map = data.get("sp_to_blender")
    if not isinstance(sp_map, dict):
        sp_map = {}
    bl_map = data.get("blender_to_sp")
    if not isinstance(bl_map, dict):
        bl_map = {}
    sp_map[sp_key] = str(blender_file)
    bl_map[bl_key] = str(sp_project_file)
    data["sp_to_blender"] = sp_map
    data["blender_to_sp"] = bl_map
    save_link_registry(data, prefs)


def paths_match(left, right):
    if not left or not right:
        return False
    return normalize_path(left).lower() == normalize_path(right).lower()


def parse_suffixes(text):
    if not text:
        return []
    parts = [part.strip().lower() for part in text.split(",")]
    return [part for part in parts if part]


def is_name_with_suffix(name, suffixes):
    lname = name.lower()
    for suffix in suffixes:
        if lname.endswith(suffix):
            return True
    return False


def collection_in_scene(scene, collection):
    if not scene or not collection:
        return False
    root = getattr(scene, "collection", None)
    if not root:
        return False
    if collection == root:
        return True
    if hasattr(root, "children_recursive"):
        return collection in root.children_recursive
    return collection in root.children


def _scene_collection_poll(scene, collection):
    return collection_in_scene(scene, collection)


def _find_layer_collections(layer_collection, target_collection, results):
    if not layer_collection or not target_collection:
        return
    if layer_collection.collection == target_collection:
        results.append(layer_collection)
    for child in layer_collection.children:
        _find_layer_collections(child, target_collection, results)


def collect_collection_meshes(collection, selected_only=False, selected_names=None):
    if not collection:
        return []
    try:
        objects = collection.all_objects
    except AttributeError:
        objects = collection.objects
    results = []
    seen = set()
    for obj in objects:
        if obj.type != "MESH":
            continue
        if selected_only and selected_names and obj.name not in selected_names:
            continue
        if obj.name in seen:
            continue
        seen.add(obj.name)
        results.append(obj)
    return results


def collect_low_poly_objects(context, prefs):
    scene = context.scene
    selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
    selected_names = None
    if selected_only:
        selected_names = {
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        }
    low_collection = getattr(scene, "gob_sp_low_poly_collection", None)
    if not collection_in_scene(scene, low_collection):
        low_collection = None
    if low_collection:
        collection_meshes = collect_collection_meshes(
            low_collection,
            selected_only=selected_only,
            selected_names=selected_names,
        )
        if collection_meshes:
            return collection_meshes
    suffixes = parse_suffixes(getattr(prefs, "low_poly_suffixes", ""))
    search_pool = context.selected_objects if selected_only else context.scene.objects
    if suffixes:
        candidates = [
            obj for obj in search_pool
            if obj.type == "MESH" and is_name_with_suffix(obj.name, suffixes)
        ]
        if candidates:
            return candidates
        if selected_only:
            return [obj for obj in search_pool if obj.type == "MESH"]
    return [obj for obj in context.selected_objects if obj.type == "MESH"]


def get_prefs(context):
    addon = context.preferences.addons.get(__name__)
    return addon.preferences if addon else None


def get_bridge_root(prefs):
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        return Path(env_path).expanduser()
    path = prefs.bridge_dir if prefs and prefs.bridge_dir else default_bridge_dir()
    return Path(path).expanduser()


def get_project_name(context):
    if bpy.data.filepath:
        return sanitize_name(Path(bpy.data.filepath).stem)
    if context.active_object:
        return sanitize_name(context.active_object.name)
    return "untitled"


def get_project_dir(context, prefs):
    blender_file = get_blender_file_path_or_temp(prefs)
    return resolve_project_dir_for_blender(context, prefs, blender_file)


def get_project_dir_fast(context, prefs):
    blender_file = get_blender_file_path_or_temp(prefs)
    cached = cached_project_dir(blender_file)
    if cached:
        return cached
    base_dir = get_bridge_root(prefs) / get_project_name(context)
    if blender_file and base_dir.exists():
        manifest_path = find_project_manifest_path(base_dir)
        manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
        if manifest_matches_blender_file(manifest, blender_file):
            set_cached_project_dir(blender_file, base_dir)
            return base_dir
    return base_dir


def bridge_root_hint_path():
    return Path(default_bridge_dir()) / BRIDGE_ROOT_HINT_FILENAME


def shared_bridge_root_hint_path():
    return Path.home() / BRIDGE_SHARED_HINT_DIRNAME / BRIDGE_ROOT_HINT_FILENAME


def read_bridge_root_hint(path):
    if not path:
        return None
    path = Path(path)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    root = data.get("bridge_root")
    if not root:
        return None
    return Path(root).expanduser()


def write_bridge_root_hint(root_path):
    if not root_path:
        return
    hint_paths = [bridge_root_hint_path(), shared_bridge_root_hint_path()]
    payload = {"bridge_root": str(Path(root_path).expanduser())}
    for hint_path in hint_paths:
        try:
            ensure_dir(hint_path.parent)
            with open(hint_path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, indent=2, ensure_ascii=True)
        except OSError:
            continue




def ensure_dir(path):
    Path(path).mkdir(parents=True, exist_ok=True)


def write_manifest(path, data):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2)


def read_manifest(path):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, json.JSONDecodeError):
        return None


def read_active_sp_info(path):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    if not data.get("project_open"):
        return None
    project_dir = data.get("project_dir")
    if not project_dir:
        return None
    timestamp = data.get("timestamp")
    try:
        timestamp = float(timestamp)
    except (TypeError, ValueError):
        timestamp = 0.0
    if not timestamp:
        try:
            timestamp = Path(path).stat().st_mtime
        except OSError:
            timestamp = 0.0
    return {
        "project_dir": Path(project_dir),
        "project_name": data.get("project_name"),
        "timestamp": timestamp,
        "sp_project_file": data.get("sp_project_file"),
        "blender_file": data.get("blender_file"),
    }


def find_active_sp_project_info(prefs, max_age=ACTIVE_SP_INFO_MAX_AGE):
    now = time.time()
    best = None
    best_time = 0.0
    for root in get_candidate_bridge_roots(prefs):
        candidate = Path(root) / ACTIVE_SP_INFO_FILENAME
        if not candidate.exists():
            continue
        info = read_active_sp_info(candidate)
        if not info:
            continue
        ts = info.get("timestamp", 0.0) or 0.0
        if max_age and ts and now - ts > max_age:
            continue
        if ts > best_time:
            best_time = ts
            best = info
    return best


def active_blender_info_paths(prefs=None, project_dir=None):
    roots = []
    docs_root = documents_bridge_root()
    if docs_root:
        roots.append(Path(docs_root))
    for root in get_candidate_bridge_roots(prefs):
        if not root:
            continue
        try:
            root_path = Path(root)
        except TypeError:
            continue
        roots.append(root_path)
    if project_dir:
        roots.append(project_meta_dir(project_dir))
    unique = []
    seen = set()
    for root in roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(root / ACTIVE_BLENDER_INFO_FILENAME)
    return unique


def write_active_blender_info(context=None, prefs=None):
    if context is None:
        context = bpy.context
    if context is None:
        return
    prefs = prefs or get_prefs(context)
    project_dir = get_project_dir(context, prefs)
    info = {
        "timestamp": time.time(),
        "project_open": True,
        "project_name": get_project_name(context),
        "project_dir": str(project_dir),
    }
    blender_file = get_blender_file_path_or_temp(prefs)
    if blender_file:
        info["blender_file"] = blender_file
    for path in active_blender_info_paths(prefs, project_dir):
        try:
            ensure_dir(path.parent)
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(info, handle, indent=2, ensure_ascii=True)
        except OSError:
            continue


def update_manifest_blender_file(old_blender_file, new_blender_file, prefs=None):
    if not old_blender_file or not new_blender_file:
        return
    manifest_path = find_manifest_for_blender_file(
        get_candidate_bridge_roots(prefs),
        old_blender_file,
    )
    if not manifest_path:
        return
    manifest = read_manifest(manifest_path)
    if not isinstance(manifest, dict):
        return
    manifest["blender_file"] = str(new_blender_file)
    project_dir = project_dir_from_manifest_path(manifest_path)
    target_path = project_manifest_path(project_dir)
    if target_path:
        ensure_dir(target_path.parent)
        write_manifest(target_path, manifest)


def sync_saved_blender_file(context=None, prefs=None):
    global _last_blender_file
    if context is None:
        context = bpy.context
    if context is None:
        _last_blender_file = None
        return
    prefs = prefs or get_prefs(context)
    current_real = bpy.data.filepath
    current = current_real or temp_blender_file_path(prefs)
    if _last_blender_file is None:
        _last_blender_file = current
        return
    if current_real and not paths_match(current_real, _last_blender_file):
        active_info = resolve_active_sp_project_info(context, prefs)
        sp_project_file = ""
        if active_info:
            sp_project_file = str(active_info.get("sp_project_file") or "")
        if not sp_project_file:
            sp_project_file = get_linked_sp_project_path(
                get_project_dir(context, prefs),
                active_info=active_info,
                blender_file=_last_blender_file,
                prefs=prefs,
            )
        if sp_project_file:
            update_link_registry(
                sp_project_file=sp_project_file,
                blender_file=current_real,
                prefs=prefs,
            )
        update_manifest_blender_file(_last_blender_file, current_real, prefs=prefs)
        project_dir = cached_project_dir(_last_blender_file)
        if not project_dir:
            project_dir = resolve_project_dir_for_blender(context, prefs, _last_blender_file)
        if project_dir:
            set_cached_project_dir(current_real, project_dir)
        _last_blender_file = current_real
        return
    _last_blender_file = current


def _update_active_blender_info(_context=None):
    try:
        context = bpy.context
        prefs = get_prefs(context) if context else None
        sync_saved_blender_file(context, prefs)
        write_active_blender_info(context, prefs)
    except Exception:
        pass
    return None


def _active_blender_heartbeat():
    _update_active_blender_info()
    return 30.0


def get_candidate_bridge_roots(prefs):
    roots = []
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        roots.append(Path(env_path))
    if prefs and prefs.bridge_dir:
        roots.append(Path(prefs.bridge_dir))
    hint = read_bridge_root_hint(bridge_root_hint_path())
    if hint:
        roots.append(hint)
    shared_hint = read_bridge_root_hint(shared_bridge_root_hint_path())
    if shared_hint:
        roots.append(shared_hint)
    docs = windows_documents_dir()
    if docs:
        roots.append(Path(docs) / "GoB_SP_Bridge")
    roots.append(Path.home() / "Documents" / "GoB_SP_Bridge")
    if sys.platform == "darwin":
        icloud_docs = (
            Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs" / "Documents"
        )
        roots.append(icloud_docs / "GoB_SP_Bridge")
    for var in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial"):
        env = os.environ.get(var)
        if env:
            roots.append(Path(env) / "Documents" / "GoB_SP_Bridge")
    unique = []
    seen = set()
    for root in roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(root)
    return unique


def find_latest_manifest(bridge_roots, source=None):
    best_path = None
    best_time = -1.0
    for root in bridge_roots:
        if not root or not root.exists():
            continue
        for candidate in root.rglob(MANIFEST_FILENAME):
            try:
                mtime = candidate.stat().st_mtime
            except OSError:
                continue
            if source:
                manifest = read_manifest(candidate)
                if not manifest or manifest.get("source") != source:
                    continue
            if mtime > best_time:
                best_time = mtime
                best_path = candidate
    return best_path


def find_manifest_for_blender_file(bridge_roots, blender_file, source=None):
    if not blender_file:
        return None
    best_path = None
    best_time = -1.0
    for root in bridge_roots:
        if not root or not root.exists():
            continue
        for candidate in root.rglob(MANIFEST_FILENAME):
            try:
                mtime = candidate.stat().st_mtime
            except OSError:
                continue
            manifest = read_manifest(candidate)
            if not manifest:
                continue
            if source and manifest.get("source") != source:
                continue
            manifest_blender = manifest.get("blender_file")
            if not manifest_blender or not paths_match(manifest_blender, blender_file):
                continue
            if mtime > best_time:
                best_time = mtime
                best_path = candidate
    return best_path


def find_manifest_for_sp_project_file(bridge_roots, sp_project_file, source=None):
    if not sp_project_file:
        return None
    best_path = None
    best_time = -1.0
    for root in bridge_roots:
        if not root or not root.exists():
            continue
        for candidate in root.rglob(MANIFEST_FILENAME):
            try:
                mtime = candidate.stat().st_mtime
            except OSError:
                continue
            manifest = read_manifest(candidate)
            if not manifest:
                continue
            if source and manifest.get("source") != source:
                continue
            manifest_sp = manifest.get("sp_project_file") or manifest.get("sp_project_path")
            if not manifest_sp or not paths_match(manifest_sp, sp_project_file):
                continue
            if mtime > best_time:
                best_time = mtime
                best_path = candidate
    return best_path


def project_dir_from_linked_sp(blender_file, prefs):
    if not blender_file:
        return None
    registry = load_link_registry(prefs)
    sp_project_file = registry.get("blender_to_sp", {}).get(
        normalize_path_key(blender_file)
    )
    if not sp_project_file:
        return None
    manifest_path = find_manifest_for_sp_project_file(
        get_candidate_bridge_roots(prefs),
        sp_project_file,
    )
    if not manifest_path:
        return None
    return project_dir_from_manifest_path(manifest_path)


def build_mesh_signature(low_objects, high_objects=None):
    low_names = sorted({obj.name for obj in (low_objects or []) if obj})
    high_names = sorted({obj.name for obj in (high_objects or []) if obj})
    return {"low": low_names, "high": high_names}


def normalize_mesh_signature(value):
    if not value:
        return {"low": [], "high": []}
    if isinstance(value, dict):
        low = sorted(str(v) for v in (value.get("low") or []) if v)
        high = sorted(str(v) for v in (value.get("high") or []) if v)
        return {"low": low, "high": high}
    if isinstance(value, (list, tuple, set)):
        low = sorted(str(v) for v in value if v)
        return {"low": low, "high": []}
    return {"low": [], "high": []}


def mesh_signature_matches(manifest, signature):
    if not signature or not isinstance(manifest, dict):
        return False
    manifest_sig = normalize_mesh_signature(manifest.get("mesh_signature"))
    if not manifest_sig["low"] and not manifest_sig["high"]:
        return False
    return manifest_sig == normalize_mesh_signature(signature)


def find_latest_saved_sp_project_for_blender(bridge_roots, blender_file):
    if not blender_file:
        return ""
    best_file = ""
    best_time = -1.0
    for root in bridge_roots:
        if not root or not root.exists():
            continue
        for candidate in root.rglob(MANIFEST_FILENAME):
            try:
                mtime = candidate.stat().st_mtime
            except OSError:
                continue
            manifest = read_manifest(candidate)
            if not manifest:
                continue
            manifest_blender = get_manifest_blender_file(manifest)
            if not manifest_blender or not paths_match(manifest_blender, blender_file):
                continue
            sp_project_file = get_manifest_sp_project_file(manifest)
            if not sp_project_file or is_temp_sp_project_file(sp_project_file):
                continue
            try:
                if not Path(sp_project_file).is_file():
                    continue
            except OSError:
                continue
            if mtime > best_time:
                best_time = mtime
                best_file = sp_project_file
    return best_file


def find_manifest_for_mesh_signature(bridge_roots, blender_file, signature, source="blender"):
    if not blender_file or not signature:
        return None
    best_path = None
    best_time = -1.0
    for root in bridge_roots:
        if not root or not root.exists():
            continue
        for candidate in root.rglob(MANIFEST_FILENAME):
            try:
                mtime = candidate.stat().st_mtime
            except OSError:
                continue
            manifest = read_manifest(candidate)
            if not manifest:
                continue
            if source and manifest.get("source") != source:
                continue
            manifest_blender = get_manifest_blender_file(manifest)
            if not manifest_blender or not paths_match(manifest_blender, blender_file):
                continue
            if not mesh_signature_matches(manifest, signature):
                continue
            if mtime > best_time:
                best_time = mtime
                best_path = candidate
    return best_path


def project_dir_signature_matches(project_dir, signature):
    if not project_dir or not signature:
        return True
    manifest_path = find_project_manifest_path(project_dir)
    if not manifest_path or not manifest_path.exists():
        return True
    manifest = read_manifest(manifest_path)
    if not isinstance(manifest, dict):
        return True
    if "mesh_signature" not in manifest:
        return True
    return mesh_signature_matches(manifest, signature)


def resolve_sp_project_candidate(sp_project_file, blender_file, prefs=None):
    if not sp_project_file:
        return ""
    if is_temp_sp_project_file(sp_project_file, prefs):
        fallback = find_latest_saved_sp_project_for_blender(
            get_candidate_bridge_roots(prefs),
            blender_file,
        )
        return fallback or sp_project_file
    try:
        if Path(sp_project_file).is_file():
            return sp_project_file
    except OSError:
        return ""
    fallback = find_latest_saved_sp_project_for_blender(
        get_candidate_bridge_roots(prefs),
        blender_file,
    )
    return fallback or ""


def get_manifest_sp_project_file(manifest):
    if not isinstance(manifest, dict):
        return ""
    value = manifest.get("sp_project_file") or manifest.get("sp_project_path")
    return str(value) if value else ""


def get_manifest_link_sp_project_file(manifest):
    if not isinstance(manifest, dict):
        return ""
    value = manifest.get("link_sp_project_file") or manifest.get("linked_sp_project_file")
    return str(value) if value else ""


def get_manifest_blender_file(manifest):
    if not isinstance(manifest, dict):
        return ""
    value = manifest.get("blender_file")
    return str(value) if value else ""


def resolve_active_sp_project_info(context, prefs):
    project_dir = get_project_dir(context, prefs)
    if project_dir:
        info = read_active_sp_info(project_meta_dir(project_dir) / ACTIVE_SP_INFO_FILENAME)
        if info:
            return info
    active_info = find_active_sp_project_info(prefs)
    if not active_info:
        return None
    blender_file = get_blender_file_path_or_temp(prefs)
    blender_file_is_temp = is_temp_blender_file(blender_file, prefs)
    sp_project_file = str(active_info.get("sp_project_file") or "")
    linked_sp_project = ""
    if blender_file:
        linked_sp_project = get_linked_sp_project_path(
            project_dir,
            active_info=None,
            blender_file=blender_file,
            prefs=prefs,
        )
    if linked_sp_project and sp_project_file:
        if not paths_match(sp_project_file, linked_sp_project):
            return None
    if blender_file:
        if paths_match(active_info.get("blender_file"), blender_file):
            return active_info
        if sp_project_file:
            registry = load_link_registry(prefs)
            linked_blender = registry.get("sp_to_blender", {}).get(
                normalize_path_key(sp_project_file)
            )
            if linked_blender and paths_match(linked_blender, blender_file):
                return active_info
            manifest_path = find_manifest_for_sp_project_file(
                get_candidate_bridge_roots(prefs),
                sp_project_file,
            )
            if manifest_path:
                manifest = read_manifest(manifest_path)
                manifest_blender = get_manifest_blender_file(manifest)
                if manifest_blender and paths_match(manifest_blender, blender_file):
                    return active_info
        if project_dir and sp_project_file:
            manifest = read_manifest(find_project_manifest_path(project_dir))
            if manifest and paths_match(get_manifest_sp_project_file(manifest), sp_project_file):
                return active_info
        if not blender_file_is_temp:
            return None
    if project_dir and active_info.get("project_dir") == project_dir:
        return active_info
    if not blender_file_is_temp:
        current_name = get_project_name(context)
        if (active_info.get("project_name") and current_name and
                active_info["project_name"].lower() == current_name.lower()):
            return active_info
    if project_dir and sp_project_file:
        manifest = read_manifest(find_project_manifest_path(project_dir))
        if manifest and paths_match(get_manifest_sp_project_file(manifest), sp_project_file):
            return active_info
    return None


def get_linked_sp_project_path(
    project_dir,
    active_info=None,
    blender_file=None,
    prefs=None,
):
    if active_info:
        sp_project_file = str(active_info.get("sp_project_file") or "")
        if sp_project_file:
            return sp_project_file
    if blender_file:
        registry = load_link_registry(prefs)
        sp_project_file = registry.get("blender_to_sp", {}).get(
            normalize_path_key(blender_file)
        )
        if sp_project_file:
            return str(sp_project_file)
    sp_project_file = ""
    if blender_file:
        manifest_path = find_manifest_for_blender_file(
            get_candidate_bridge_roots(prefs),
            blender_file,
        )
        if manifest_path:
            manifest = read_manifest(manifest_path)
            sp_project_file = get_manifest_sp_project_file(manifest)
            if sp_project_file:
                return str(sp_project_file)
    if project_dir:
        manifest = read_manifest(find_project_manifest_path(project_dir))
        sp_project_file = get_manifest_sp_project_file(manifest)
    return str(sp_project_file) if sp_project_file else ""


def resolve_linked_sp_project_file(
    project_dir,
    active_info=None,
    blender_file=None,
    prefs=None,
):
    sp_project_file = get_linked_sp_project_path(
        project_dir,
        active_info=active_info,
        blender_file=blender_file,
        prefs=prefs,
    )
    return resolve_sp_project_candidate(sp_project_file, blender_file, prefs)


def get_linked_sp_project_path_fast(
    project_dir,
    active_info=None,
    blender_file=None,
    prefs=None,
):
    if active_info:
        sp_project_file = str(active_info.get("sp_project_file") or "")
        if sp_project_file:
            return sp_project_file
    if blender_file:
        registry = load_link_registry(prefs)
        sp_project_file = registry.get("blender_to_sp", {}).get(
            normalize_path_key(blender_file)
        )
        if sp_project_file:
            return str(sp_project_file)
    if project_dir:
        manifest = read_manifest(find_project_manifest_path(project_dir))
        sp_project_file = get_manifest_sp_project_file(manifest)
        if sp_project_file:
            return str(sp_project_file)
    return ""


def resolve_linked_sp_project_file_fast(
    project_dir,
    active_info=None,
    blender_file=None,
    prefs=None,
):
    sp_project_file = get_linked_sp_project_path_fast(
        project_dir,
        active_info=active_info,
        blender_file=blender_file,
        prefs=prefs,
    )
    if not sp_project_file:
        return ""
    if is_temp_sp_project_file(sp_project_file, prefs):
        return sp_project_file
    try:
        if Path(sp_project_file).is_file():
            return sp_project_file
    except OSError:
        return ""
    return ""


def folder_size_bytes(path):
    if not path or not path.exists():
        return 0
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for filename in filenames:
            try:
                total += (Path(dirpath) / filename).stat().st_size
            except OSError:
                continue
    return total


def bridge_cache_size_bytes(prefs):
    return folder_size_bytes(get_bridge_root(prefs))


def project_cache_size_bytes(context, prefs):
    return folder_size_bytes(get_project_dir(context, prefs))


def clear_cache_dir(path):
    if not path.exists():
        return "empty"
    try:
        shutil.rmtree(path)
    except OSError:
        return "error"
    ensure_dir(path)
    return "cleared"


def clear_cache_dir_except(root, keep_paths=None):
    if not root.exists():
        return "empty"
    keep = set()
    if keep_paths:
        for path in keep_paths:
            if not path:
                continue
            try:
                path_obj = Path(path).resolve()
            except OSError:
                continue
            try:
                if root.resolve() not in path_obj.parents and path_obj != root.resolve():
                    continue
            except OSError:
                continue
            keep.add(str(path_obj).lower())
    try:
        for child in root.iterdir():
            try:
                child_key = str(child.resolve()).lower()
            except OSError:
                child_key = str(child).lower()
            if child.is_file() and child.name in {BRIDGE_ROOT_HINT_FILENAME, ACTIVE_SP_INFO_FILENAME}:
                continue
            if child_key in keep:
                continue
            try:
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink()
            except OSError:
                return "error"
    except OSError:
        return "error"
    ensure_dir(root)
    return "cleared"


def cache_limit_bytes(prefs):
    if not prefs:
        return 0
    try:
        limit_gb = float(getattr(prefs, "cache_limit_gb", 0.0))
    except (TypeError, ValueError):
        return 0
    if limit_gb <= 0:
        return 0
    return limit_gb * 1024 ** 3


def format_bytes(value):
    size = float(value or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024.0 or unit == "TB":
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


def local_version_string():
    version = bl_info.get("version")
    if isinstance(version, (tuple, list)):
        return ".".join(str(part) for part in version)
    return str(version or "0.0.0")


def parse_version(value):
    parts = re.findall(r"\d+", str(value))
    return tuple(int(part) for part in parts) if parts else (0,)


def is_version_newer(remote, local):
    return parse_version(remote) > parse_version(local)


def check_for_updates():
    try:
        with urllib.request.urlopen(UPDATE_URL, timeout=4) as response:
            data = json.load(response)
    except (OSError, json.JSONDecodeError, urllib.error.URLError) as exc:
        return {"status": "error", "error": str(exc)}
    if not isinstance(data, dict):
        return {"status": "error", "error": "Invalid update data"}
    blender_info = data.get("blender") or {}
    if not isinstance(blender_info, dict):
        return {"status": "error", "error": "Missing Blender update data"}
    remote_version = str(blender_info.get("version") or "").strip()
    if not remote_version:
        return {"status": "error", "error": "Missing remote version"}
    local_version = local_version_string()
    if not is_version_newer(remote_version, local_version):
        return {
            "status": "none",
            "local_version": local_version,
            "remote_version": remote_version,
        }
    return {
        "status": "update",
        "info": {
            "version": remote_version,
            "download_url": blender_info.get("download_url"),
            "notes": data.get("notes"),
            "local_version": local_version,
        },
    }


def detect_map_type(stem_lower):
    for keyword in ("opacity", "alpha", "transparency", "transparent", "cutout"):
        if keyword in stem_lower:
            return "opacity", keyword
    match = re.search(r"(?:^|[._\\-])base(?:$|[._\\-])", stem_lower)
    if match:
        return "base_color", match.group(0)
    if "materialparams" in stem_lower or "materialparam" in stem_lower:
        return "orm", "materialparams"
    if "maskmap" in stem_lower:
        return "orm", "maskmap"
    if "occlusionroughnessmetallic" in stem_lower:
        return "orm", "occlusionroughnessmetallic"
    if "occlusionroughnessmetal" in stem_lower:
        return "orm", "occlusionroughnessmetal"
    match = re.search(r"occlusion[._\\-]?roughness[._\\-]?metallic", stem_lower)
    if match:
        return "orm", match.group(0)
    match = re.search(r"occlusion[._\\-]?roughness[._\\-]?metal", stem_lower)
    if match:
        return "orm", match.group(0)
    match = re.search(r"(?:^|[._\\-])arm(?:$|[._\\-])", stem_lower)
    if match:
        return "orm", match.group(0)
    match = re.search(r"(?:^|[._\\-])orm(?:$|[._\\-])", stem_lower)
    if match:
        return "orm", match.group(0)
    match = re.search(r"metallic[._\\-]?roughness", stem_lower)
    if match:
        return "metallic_roughness", match.group(0)
    match = re.search(r"roughness[._\\-]?metallic", stem_lower)
    if match:
        return "metallic_roughness", match.group(0)
    match = re.search(r"metallic[._\\-]?smoothness", stem_lower)
    if match:
        return "metallic_smoothness", match.group(0)
    match = re.search(r"specular[._\\-]?smoothness", stem_lower)
    if match:
        return "specular_smoothness", match.group(0)
    match = re.search(r"specular[._\\-]?gloss", stem_lower)
    if match:
        return "specular_smoothness", match.group(0)
    if re.search(r"specgloss", stem_lower):
        return "specular_smoothness", "specgloss"
    match = re.search(r"mask[._\\-]?map", stem_lower)
    if match:
        return "mask", match.group(0)
    for keyword, map_type in MAP_KEYWORDS:
        if keyword in stem_lower:
            return map_type, keyword
    if "rgb" in stem_lower:
        return "base_color", "rgb"
    return None, None


def should_invert_normal_y(path, manifest=None):
    if manifest:
        fmt = manifest.get("normal_map_format") or manifest.get("normal_format")
        if fmt:
            fmt_lower = str(fmt).lower()
            if "directx" in fmt_lower or "d3d" in fmt_lower or fmt_lower == "dx":
                return True
            if "opengl" in fmt_lower or "ogl" in fmt_lower or fmt_lower == "gl":
                return False
        if "normal_map_y_invert" in manifest:
            return bool(manifest.get("normal_map_y_invert"))
    name = Path(path).stem.lower()
    if "directx" in name or "d3d" in name or "_dx" in name or name.endswith("dx"):
        return True
    if "opengl" in name or "ogl" in name or "_gl" in name or name.endswith("gl"):
        return False
    return False



def guess_texture_set_name(stem, keyword, fallback=None):
    if not keyword:
        return stem
    lower = stem.lower()
    idx = lower.find(keyword)
    if idx <= 0:
        return fallback or stem
    base = stem[:idx].rstrip(" _-.")
    return base or fallback or stem


def normalize_texset_name(name):
    if not name:
        return name
    match = re.match(r"(?i)^(b2sp|sp2b)[._-]+(.+)$", name)
    if match:
        return match.group(2)
    return name


def normalize_match_name(name):
    if not name:
        return ""
    name = normalize_texset_name(str(name))
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def map_keyword_in_name(name):
    if not name:
        return False
    map_type, _ = detect_map_type(name.lower())
    return map_type is not None


def guess_texset_from_path(path_obj):
    parts = [part for part in path_obj.parts[:-1] if part]
    for part in reversed(parts):
        lower = part.lower()
        if lower in {"textures", "texture", "maps", "map", "export", "exports", "output"}:
            continue
        if map_keyword_in_name(lower):
            continue
        return part
    return None


def gather_texture_paths(manifest):
    paths = []
    if not manifest:
        return paths
    textures_dir = manifest.get("textures_dir")
    base_dir = Path(textures_dir).expanduser() if textures_dir else None
    if isinstance(manifest.get("textures"), list):
        for raw in manifest["textures"]:
            if not raw:
                continue
            path = Path(raw).expanduser()
            if not path.is_absolute() and base_dir:
                path = base_dir / path
            paths.append(str(path))
    if base_dir:
        for ext in IMAGE_EXTS:
            for path in base_dir.rglob(f"*{ext}"):
                if path.is_file():
                    paths.append(str(path))
    seen = set()
    unique = []
    for path in paths:
        if path in seen:
            continue
        seen.add(path)
        unique.append(path)
    return unique


def group_textures(texture_paths):
    grouped = {}
    for path in texture_paths:
        path_obj = Path(path)
        stem = path_obj.stem
        stem_lower = stem.lower()
        map_type, keyword = detect_map_type(stem_lower)
        if not map_type:
            path_key = re.sub(r"[\\\\/]+", "_", str(path_obj).lower())
            map_type, keyword = detect_map_type(path_key)
        if not map_type:
            continue
        if map_type == "opacity":
            for base_hint in ("basecolor", "base_color", "albedo", "diffuse", "color"):
                if base_hint in stem_lower:
                    map_type = "base_color"
                    keyword = base_hint
                    break
        fallback = None
        lower_parts = [part.lower() for part in path_obj.parts]
        if "textures" in lower_parts:
            idx = len(lower_parts) - 1 - lower_parts[::-1].index("textures")
            if idx + 1 < len(path_obj.parts):
                fallback = path_obj.parts[idx + 1]
        texset = guess_texture_set_name(stem, keyword, fallback=fallback)
        if texset == stem and map_keyword_in_name(stem_lower):
            guessed = fallback or guess_texset_from_path(path_obj)
            if guessed:
                texset = guessed
        texset = normalize_texset_name(texset)
        if texset:
            texset = texset.strip()
        grouped.setdefault(texset, {})[map_type] = path
    return grouped


def load_image(path):
    try:
        image = bpy.data.images.load(path, check_existing=True)
    except RuntimeError:
        return None
    try:
        image.reload()
    except RuntimeError:
        pass
    return image


def build_material(mat, maps, normal_y_invert=False, manifest=None):
    mat["gob_bridge_material"] = True
    mat.use_nodes = True
    if hasattr(mat, "blend_method"):
        mat.blend_method = "OPAQUE"
    if hasattr(mat, "shadow_method"):
        try:
            mat.shadow_method = "OPAQUE"
        except Exception:
            pass
    if hasattr(mat, "alpha_threshold"):
        try:
            mat.alpha_threshold = 0.5
        except Exception:
            pass
    if hasattr(mat, "use_backface_culling"):
        mat.use_backface_culling = False
    if hasattr(mat, "show_transparent_back"):
        mat.show_transparent_back = True
    nodes = mat.node_tree.nodes
    links = mat.node_tree.links
    nodes.clear()

    output = nodes.new("ShaderNodeOutputMaterial")
    output.location = (500, 0)
    principled = nodes.new("ShaderNodeBsdfPrincipled")
    principled.location = (200, 0)
    links.new(principled.outputs["BSDF"], output.inputs["Surface"])

    base_node = None
    ao_node = None
    height_node = None
    gloss_node = None
    normal_node = None
    emission_node = None
    opacity_node = None
    metallic_node = None
    roughness_node = None
    orm_node = None
    mask_node = None
    metallic_roughness_node = None
    metallic_smoothness_node = None
    specular_node = None
    specular_smoothness_node = None

    y = 300
    step = -220
    for map_type in ("base_color", "orm", "metallic_roughness", "metallic_smoothness",
                     "mask", "ao", "metallic", "roughness", "glossiness",
                     "specular_smoothness", "specular", "normal", "height",
                     "opacity", "emission"):
        if map_type not in maps:
            continue
        tex = nodes.new("ShaderNodeTexImage")
        tex.location = (-400, y)
        y += step
        image = load_image(maps[map_type])
        if not image:
            continue
        tex.image = image
        if map_type in {"normal", "roughness", "metallic", "ao", "height",
                        "opacity", "glossiness", "orm", "metallic_roughness",
                        "metallic_smoothness", "mask", "specular_smoothness",
                        "specular"}:
            try:
                image.colorspace_settings.name = "Non-Color"
            except TypeError:
                pass
        if map_type == "base_color":
            base_node = tex
        elif map_type == "orm":
            orm_node = tex
        elif map_type == "metallic_roughness":
            metallic_roughness_node = tex
        elif map_type == "metallic_smoothness":
            metallic_smoothness_node = tex
        elif map_type == "mask":
            mask_node = tex
        elif map_type == "ao":
            ao_node = tex
        elif map_type == "metallic":
            metallic_node = tex
        elif map_type == "roughness":
            roughness_node = tex
        elif map_type == "glossiness":
            gloss_node = tex
        elif map_type == "specular_smoothness":
            specular_smoothness_node = tex
        elif map_type == "specular":
            specular_node = tex
        elif map_type == "normal":
            normal_node = tex
        elif map_type == "height":
            height_node = tex
        elif map_type == "opacity":
            opacity_node = tex
        elif map_type == "emission":
            emission_node = tex

    ao_output = ao_node.outputs["Color"] if ao_node else None
    metallic_output = metallic_node.outputs["Color"] if metallic_node else None
    roughness_output = roughness_node.outputs["Color"] if roughness_node else None
    specular_output = specular_node.outputs["Color"] if specular_node else None
    if orm_node:
        separate = nodes.new("ShaderNodeSeparateRGB")
        separate.location = (-220, -300)
        links.new(orm_node.outputs["Color"], separate.inputs["Image"])
        if ao_output is None:
            ao_output = separate.outputs["R"]
        if roughness_output is None:
            roughness_output = separate.outputs["G"]
        if metallic_output is None:
            metallic_output = separate.outputs["B"]
    if metallic_roughness_node:
        separate = nodes.new("ShaderNodeSeparateRGB")
        separate.location = (-220, -120)
        links.new(metallic_roughness_node.outputs["Color"], separate.inputs["Image"])
        if roughness_output is None:
            roughness_output = separate.outputs["G"]
        if metallic_output is None:
            metallic_output = separate.outputs["B"]
    if metallic_smoothness_node:
        separate = nodes.new("ShaderNodeSeparateRGB")
        separate.location = (-220, -180)
        links.new(metallic_smoothness_node.outputs["Color"], separate.inputs["Image"])
        if metallic_output is None:
            metallic_output = separate.outputs["R"]
        if roughness_output is None:
            invert = nodes.new("ShaderNodeInvert")
            invert.inputs["Fac"].default_value = 1.0
            invert.location = (-120, -200)
            links.new(metallic_smoothness_node.outputs["Alpha"], invert.inputs["Color"])
            roughness_output = invert.outputs["Color"]
    if mask_node:
        separate = nodes.new("ShaderNodeSeparateRGB")
        separate.location = (-220, -240)
        links.new(mask_node.outputs["Color"], separate.inputs["Image"])
        if metallic_output is None:
            metallic_output = separate.outputs["R"]
        if ao_output is None:
            ao_output = separate.outputs["G"]
        if roughness_output is None:
            invert = nodes.new("ShaderNodeInvert")
            invert.inputs["Fac"].default_value = 1.0
            invert.location = (-120, -260)
            links.new(mask_node.outputs["Alpha"], invert.inputs["Color"])
            roughness_output = invert.outputs["Color"]
    if specular_smoothness_node:
        if specular_output is None:
            specular_output = specular_smoothness_node.outputs["Color"]
        if roughness_output is None:
            invert = nodes.new("ShaderNodeInvert")
            invert.inputs["Fac"].default_value = 1.0
            invert.location = (-120, -320)
            links.new(specular_smoothness_node.outputs["Alpha"], invert.inputs["Color"])
            roughness_output = invert.outputs["Color"]

    allow_basecolor_alpha = bool(manifest and manifest.get("basecolor_has_opacity"))
    opacity_output = None
    if opacity_node:
        opacity_output = opacity_node.outputs["Color"]
    elif base_node and allow_basecolor_alpha:
        image = getattr(base_node, "image", None)
        try:
            if image and getattr(image, "channels", 0) >= 4:
                opacity_output = base_node.outputs["Alpha"]
        except Exception:
            pass

    if base_node and ao_output:
        mix = nodes.new("ShaderNodeMixRGB")
        mix.blend_type = "MULTIPLY"
        mix.inputs["Fac"].default_value = 1.0
        mix.location = (-50, 200)
        links.new(base_node.outputs["Color"], mix.inputs["Color1"])
        links.new(ao_output, mix.inputs["Color2"])
        links.new(mix.outputs["Color"], principled.inputs["Base Color"])
    elif base_node:
        links.new(base_node.outputs["Color"], principled.inputs["Base Color"])

    if metallic_output:
        links.new(metallic_output, principled.inputs["Metallic"])

    if roughness_output:
        links.new(roughness_output, principled.inputs["Roughness"])
    elif gloss_node:
        invert = nodes.new("ShaderNodeInvert")
        invert.location = (-100, -260)
        links.new(gloss_node.outputs["Color"], invert.inputs["Color"])
        links.new(invert.outputs["Color"], principled.inputs["Roughness"])

    if specular_output:
        specular_input = None
        for socket in principled.inputs:
            if socket.name in {"Specular", "Specular IOR Level"}:
                specular_input = socket
                break
        if specular_input:
            if getattr(specular_output, "type", "") != "VALUE":
                rgb_to_bw = nodes.new("ShaderNodeRGBToBW")
                rgb_to_bw.location = (-60, -340)
                links.new(specular_output, rgb_to_bw.inputs["Color"])
                links.new(rgb_to_bw.outputs["Val"], specular_input)
            else:
                links.new(specular_output, specular_input)

    if normal_node:
        normal_map = nodes.new("ShaderNodeNormalMap")
        normal_map.location = (-50, -520)
        if normal_y_invert:
            separate = nodes.new("ShaderNodeSeparateRGB")
            separate.location = (-250, -520)
            invert = nodes.new("ShaderNodeInvert")
            invert.location = (-200, -640)
            combine = nodes.new("ShaderNodeCombineRGB")
            combine.location = (-100, -520)
            links.new(normal_node.outputs["Color"], separate.inputs["Image"])
            links.new(separate.outputs["R"], combine.inputs["R"])
            links.new(separate.outputs["G"], invert.inputs["Color"])
            links.new(invert.outputs["Color"], combine.inputs["G"])
            links.new(separate.outputs["B"], combine.inputs["B"])
            links.new(combine.outputs["Image"], normal_map.inputs["Color"])
        else:
            links.new(normal_node.outputs["Color"], normal_map.inputs["Color"])
        links.new(normal_map.outputs["Normal"], principled.inputs["Normal"])

    if height_node:
        disp = nodes.new("ShaderNodeDisplacement")
        disp.inputs["Scale"].default_value = 0.1
        disp.location = (200, -520)
        links.new(height_node.outputs["Color"], disp.inputs["Height"])
        links.new(disp.outputs["Displacement"], output.inputs["Displacement"])

    if emission_node:
        emission_input = None
        for socket in principled.inputs:
            if socket.name in {"Emission", "Emission Color"}:
                emission_input = socket
                break
        if emission_input:
            links.new(emission_node.outputs["Color"], emission_input)
        emission_strength = None
        for socket in principled.inputs:
            if socket.name == "Emission Strength":
                emission_strength = socket
                break
        if emission_strength:
            emission_strength.default_value = 1.0

    if opacity_output:
        links.new(opacity_output, principled.inputs["Alpha"])
        mat.blend_method = "CLIP"
        if hasattr(mat, "alpha_threshold"):
            mat.alpha_threshold = 0.5
        if hasattr(mat, "show_transparent_back"):
            mat.show_transparent_back = False
        if hasattr(mat, "shadow_method"):
            try:
                mat.shadow_method = "HASHED"
            except Exception:
                pass

    return mat


def get_or_build_material(name, maps, normal_y_invert=False, manifest=None):
    mat = bpy.data.materials.get(name)
    if not mat:
        mat = bpy.data.materials.new(name=name)
    return build_material(mat, maps, normal_y_invert=normal_y_invert, manifest=manifest)


def assign_material_to_object(obj, material, texset_name, all_groups):
    if obj.type != "MESH":
        return
    target_slot = None
    texset_key = normalize_match_name(texset_name)
    if obj.material_slots:
        for idx, slot in enumerate(obj.material_slots):
            if slot.material and normalize_match_name(slot.material.name) == texset_key:
                target_slot = idx
                break
    if target_slot is None:
        if len(all_groups) == 1 and obj.material_slots:
            target_slot = 0
    if target_slot is None:
        if not obj.material_slots:
            obj.data.materials.append(material)
        else:
            obj.data.materials[0] = material
        return
    obj.data.materials[target_slot] = material


def find_signature_targets(context, manifest):
    if not manifest:
        return []
    signature = normalize_mesh_signature(manifest.get("mesh_signature"))
    low_names = {name for name in (signature.get("low") or []) if name}
    if not low_names:
        return []
    return [
        obj for obj in context.scene.objects
        if obj.type == "MESH" and obj.name in low_names
    ]


def find_texture_targets(context, grouped):
    if not grouped:
        return []
    keys = {normalize_match_name(key) for key in grouped if key}
    matches = []
    for obj in context.scene.objects:
        if obj.type != "MESH":
            continue
        matched = False
        for slot in obj.material_slots:
            if slot.material and normalize_match_name(slot.material.name) in keys:
                matched = True
                break
        if not matched:
            lname = normalize_match_name(obj.name)
            if any(key and key in lname for key in keys):
                matched = True
        if matched:
            matches.append(obj)
    if matches:
        return matches
    if len(keys) == 1:
        if context.active_object and context.active_object.type == "MESH":
            return [context.active_object]
        return [obj for obj in context.scene.objects if obj.type == "MESH"]
    return []


def apply_textures_to_objects(objects, grouped, manifest=None, strict=False):
    if not grouped:
        return
    materials = {}
    material_entries = []
    for texset, maps in grouped.items():
        mat_name = texset
        normal_path = maps.get("normal")
        normal_y_invert = bool(normal_path and should_invert_normal_y(normal_path, manifest=manifest))
        mat = get_or_build_material(
            mat_name,
            maps,
            normal_y_invert=normal_y_invert,
            manifest=manifest,
        )
        key = normalize_match_name(texset)
        if key:
            materials.setdefault(key, mat)
        material_entries.append((key, mat, texset))

    groups = list(material_entries)
    mesh_targets = [obj for obj in objects if obj.type == "MESH"]
    single_target = len(mesh_targets) == 1
    for obj in mesh_targets:
        assigned = False
        for idx, slot in enumerate(obj.material_slots):
            if not slot.material:
                continue
            key = normalize_match_name(slot.material.name)
            if key and key in materials:
                obj.material_slots[idx].material = materials[key]
                assigned = True
        if assigned:
            continue
        obj_key = normalize_match_name(obj.name)
        for key, mat, texset in groups:
            if key and obj_key and key in obj_key:
                assign_material_to_object(obj, mat, texset, materials)
                assigned = True
                break
        if not assigned and not strict:
            if single_target and obj.material_slots and groups:
                for idx, entry in enumerate(groups):
                    _, mat, _ = entry
                    if idx < len(obj.material_slots):
                        obj.material_slots[idx].material = mat
                    else:
                        obj.data.materials.append(mat)
                assigned = True
            elif groups:
                assign_material_to_object(obj, groups[0][1], groups[0][2], materials)


def build_fbx_export_kwargs(prefs):
    if not prefs:
        return {}
    props = bpy.ops.export_scene.fbx.get_rna_type().properties
    kwargs = {}

    def set_if(prop_name, value):
        if prop_name in props:
            kwargs[prop_name] = value

    set_if("global_scale", max(0.0001, float(prefs.fbx_export_scale)))
    set_if("apply_unit_scale", bool(prefs.fbx_apply_unit_scale))
    set_if("apply_scale_options", "FBX_SCALE_UNITS")
    if not prefs.fbx_export_custom_normals:
        set_if("use_custom_normals", False)
    return kwargs


def remove_uv_layers(mesh):
    try:
        layers = list(mesh.uv_layers)
    except AttributeError:
        return
    for layer in layers:
        try:
            mesh.uv_layers.remove(layer)
        except RuntimeError:
            continue


def object_is_valid(obj):
    try:
        name = obj.name
    except ReferenceError:
        return False
    return name in bpy.data.objects


def unique_object_name(base):
    name = base
    idx = 1
    while name in bpy.data.objects:
        name = f"{base}_{idx}"
        idx += 1
    return name


def export_fbx_objects(filepath, objects, prefs=None, strip_uvs=False):
    export_objs = [obj for obj in objects if object_is_valid(obj) and obj.type == "MESH"]
    if not export_objs:
        return False
    temp_objects = []
    renamed_objects = []
    if strip_uvs:
        for obj in export_objs:
            orig_name = obj.name
            temp_name = unique_object_name(f"{orig_name}__gob_src")
            try:
                obj.name = temp_name
                renamed_objects.append((obj, orig_name))
            except RuntimeError:
                renamed_objects.append((obj, orig_name))
            dup = obj.copy()
            dup.data = obj.data.copy()
            dup.name = orig_name
            bpy.context.scene.collection.objects.link(dup)
            temp_objects.append(dup)
        export_objs = temp_objects
    if strip_uvs:
        for obj in export_objs:
            remove_uv_layers(obj.data)
    obj_states = []
    layer_states = []
    collection_states = []
    seen_objs = set()
    seen_layers = set()
    seen_collections = set()
    view_layer = bpy.context.view_layer
    for obj in export_objs:
        try:
            obj_key = obj.as_pointer()
        except Exception:
            obj_key = id(obj)
        if obj_key not in seen_objs:
            seen_objs.add(obj_key)
            obj_states.append((
                obj,
                getattr(obj, "hide_viewport", False),
                getattr(obj, "hide_render", False),
                getattr(obj, "hide_select", False),
            ))
            try:
                obj.hide_set(False)
            except Exception:
                pass
            try:
                obj.hide_viewport = False
            except Exception:
                pass
            try:
                obj.hide_render = False
            except Exception:
                pass
            try:
                obj.hide_select = False
            except Exception:
                pass
        for collection in obj.users_collection:
            try:
                col_key = collection.as_pointer()
            except Exception:
                col_key = id(collection)
            if col_key not in seen_collections:
                seen_collections.add(col_key)
                collection_states.append((
                    collection,
                    getattr(collection, "hide_viewport", False),
                    getattr(collection, "hide_render", False),
                    getattr(collection, "hide_select", False),
                ))
                try:
                    collection.hide_viewport = False
                except Exception:
                    pass
                try:
                    collection.hide_render = False
                except Exception:
                    pass
                try:
                    collection.hide_select = False
                except Exception:
                    pass
            if view_layer and view_layer.layer_collection:
                matches = []
                _find_layer_collections(view_layer.layer_collection, collection, matches)
                for layer in matches:
                    try:
                        layer_key = layer.as_pointer()
                    except Exception:
                        layer_key = id(layer)
                    if layer_key in seen_layers:
                        continue
                    seen_layers.add(layer_key)
                    layer_states.append((
                        layer,
                        getattr(layer, "exclude", False),
                        getattr(layer, "hide_viewport", False),
                    ))
                    try:
                        layer.exclude = False
                    except Exception:
                        pass
                    try:
                        layer.hide_viewport = False
                    except Exception:
                        pass
    prev_selected = [obj for obj in bpy.context.selected_objects if object_is_valid(obj)]
    prev_active = bpy.context.view_layer.objects.active
    for obj in prev_selected:
        try:
            obj.select_set(False)
        except ReferenceError:
            continue
    for obj in export_objs:
        try:
            obj.select_set(True)
        except ReferenceError:
            continue
    if export_objs:
        bpy.context.view_layer.objects.active = export_objs[0]
    export_kwargs = build_fbx_export_kwargs(prefs)
    try:
        bpy.ops.export_scene.fbx(
            filepath=str(filepath),
            use_selection=True,
            use_mesh_modifiers=True,
            mesh_smooth_type="FACE",
            add_leaf_bones=False,
            bake_space_transform=False,
            **export_kwargs,
        )
    finally:
        for layer, was_excluded, was_hidden in layer_states:
            try:
                layer.exclude = was_excluded
            except Exception:
                pass
            try:
                layer.hide_viewport = was_hidden
            except Exception:
                pass
        for collection, was_hidden, was_render, was_select in collection_states:
            try:
                collection.hide_viewport = was_hidden
            except Exception:
                pass
            try:
                collection.hide_render = was_render
            except Exception:
                pass
            try:
                collection.hide_select = was_select
            except Exception:
                pass
        for obj, was_hidden, was_render, was_select in obj_states:
            if not object_is_valid(obj):
                continue
            try:
                obj.hide_viewport = was_hidden
            except Exception:
                pass
            try:
                obj.hide_render = was_render
            except Exception:
                pass
            try:
                obj.hide_select = was_select
            except Exception:
                pass
        for obj in export_objs:
            try:
                obj.select_set(False)
            except ReferenceError:
                continue
        if temp_objects:
            for obj in temp_objects:
                mesh_data = obj.data
                try:
                    bpy.data.objects.remove(obj, do_unlink=True)
                except RuntimeError:
                    pass
                try:
                    if mesh_data:
                        bpy.data.meshes.remove(mesh_data, do_unlink=True)
                except RuntimeError:
                    pass
        for obj, orig_name in renamed_objects:
            if object_is_valid(obj):
                try:
                    obj.name = orig_name
                except RuntimeError:
                    pass
        for obj in prev_selected:
            try:
                obj.select_set(True)
            except ReferenceError:
                continue
        bpy.context.view_layer.objects.active = prev_active
    return True


def export_selected_fbx(filepath, prefs=None, strip_uvs=False):
    return export_fbx_objects(
        filepath,
        bpy.context.selected_objects,
        prefs=prefs,
        strip_uvs=strip_uvs,
    )


def object_has_uvs(obj):
    if obj.type != "MESH":
        return False
    return bool(obj.data.uv_layers)


def mesh_triangle_count(obj):
    if obj.type != "MESH":
        return 0
    mesh = obj.data
    try:
        mesh.calc_loop_triangles()
        return len(mesh.loop_triangles)
    except Exception:
        try:
            return len(mesh.polygons)
        except Exception:
            return 0


def split_meshes_by_triangles(objects):
    mesh_items = [(obj, mesh_triangle_count(obj)) for obj in objects if obj.type == "MESH"]
    if not mesh_items:
        return [], []
    if len(mesh_items) == 1:
        return [mesh_items[0][0]], []
    mesh_items.sort(key=lambda item: item[1])
    if mesh_items[0][1] == mesh_items[-1][1]:
        return [obj for obj, _ in mesh_items], []
    best_index = 0
    best_gap = -1
    for idx in range(len(mesh_items) - 1):
        gap = mesh_items[idx + 1][1] - mesh_items[idx][1]
        if gap > best_gap:
            best_gap = gap
            best_index = idx
    low = [obj for obj, _ in mesh_items[:best_index + 1]]
    high = [obj for obj, _ in mesh_items[best_index + 1:]]
    return low, high


def collect_high_poly_objects(context, prefs, low_objects):
    candidates = collect_high_poly_candidates(context, prefs)
    if not low_objects:
        return candidates
    low_set = {obj.name for obj in low_objects}
    return [obj for obj in candidates if obj.name not in low_set]


def collect_high_poly_candidates(context, prefs):
    scene = context.scene
    objects = []
    selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
    selected_names = None
    if selected_only:
        selected_names = {
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        }
    high_collection = getattr(scene, "gob_sp_high_poly_collection", None)
    if not collection_in_scene(scene, high_collection):
        high_collection = None
    if high_collection:
        objects = collect_collection_meshes(
            high_collection,
            selected_only=selected_only,
            selected_names=selected_names,
        )
        if objects:
            return objects
    suffixes = parse_suffixes(getattr(prefs, "high_poly_suffixes", ""))
    if suffixes:
        for obj in context.scene.objects:
            if obj.type != "MESH":
                continue
            if is_name_with_suffix(obj.name, suffixes):
                objects.append(obj)
    for obj in context.scene.objects:
        if obj.type == "MESH" and obj.get("gob_high_poly"):
            objects.append(obj)
    if selected_only and selected_names:
        objects = [obj for obj in objects if obj.name in selected_names]
    unique = []
    seen = set()
    for obj in objects:
        if obj.name in seen:
            continue
        seen.add(obj.name)
        unique.append(obj)
    return unique


def import_fbx(filepath):
    before = {obj.name for obj in bpy.data.objects}
    bpy.ops.import_scene.fbx(filepath=str(filepath))
    return [obj for obj in bpy.data.objects if obj.name not in before]


def find_sp_exe(_prefs):
    for env_var in ("SUBSTANCE_PAINTER_EXE", "ADOBE_SUBSTANCE_PAINTER_EXE"):
        env_path = os.environ.get(env_var)
        if env_path:
            env_candidate = Path(env_path).expanduser()
            if env_candidate.is_file():
                return str(env_candidate)
            if sys.platform == "darwin" and env_candidate.suffix.lower() == ".app" and env_candidate.is_dir():
                return str(env_candidate)

    if os.name == "nt":
        program_files = os.environ.get("ProgramFiles", r"C:\\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)")
        candidates = [
            Path(program_files) / "Adobe" / "Adobe Substance 3D Painter" / "Adobe Substance 3D Painter.exe",
            Path(program_files) / "Adobe" / "Adobe Substance 3D Painter" / "Substance 3D Painter.exe",
            Path(program_files) / "Adobe" / "Substance 3D Painter" / "Substance 3D Painter.exe",
            Path(program_files) / "Adobe" / "Substance 3D Painter 11.1.1" / "Substance 3D Painter.exe",
            Path(program_files) / "Allegorithmic" / "Substance Painter" / "Substance Painter.exe",
        ]
        if program_files_x86:
            candidates.extend([
                Path(program_files_x86) / "Adobe" / "Adobe Substance 3D Painter" / "Adobe Substance 3D Painter.exe",
                Path(program_files_x86) / "Adobe" / "Substance 3D Painter" / "Substance 3D Painter.exe",
            ])
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)

        adobe_bases = []
        for base in (program_files, program_files_x86):
            if base:
                adobe_bases.append(Path(base) / "Adobe")
        for base in adobe_bases:
            if not base.exists():
                continue
            for exe in base.rglob("*Painter*.exe"):
                name = exe.name.lower()
                if "painter" in name and "substance" in name:
                    return str(exe)
    elif sys.platform == "darwin":
        app_candidates = [
            Path("/Applications/Adobe Substance 3D Painter.app"),
            Path("/Applications/Substance 3D Painter.app"),
            Path("/Applications/Allegorithmic/Substance Painter.app"),
            Path.home() / "Applications" / "Adobe Substance 3D Painter.app",
            Path.home() / "Applications" / "Substance 3D Painter.app",
            Path.home() / "Applications" / "Substance Painter.app",
        ]
        for candidate in app_candidates:
            if candidate.is_dir():
                return str(candidate)
        for root in (Path("/Applications"), Path.home() / "Applications"):
            if not root.exists():
                continue
            for app in root.glob("*.app"):
                name = app.name.lower()
                if "painter" in name and "substance" in name:
                    return str(app)
    return None


def open_sp_project_file(project_file, sp_exe=None):
    if not project_file:
        return False
    try:
        if sys.platform == "darwin":
            if sp_exe and sp_exe.lower().endswith(".app"):
                subprocess.Popen(["open", "-a", sp_exe, project_file])
            else:
                subprocess.Popen(["open", project_file])
            return True
        if os.name == "nt":
            if sp_exe and Path(sp_exe).is_file():
                subprocess.Popen([sp_exe, project_file])
            else:
                os.startfile(project_file)
            return True
        if sp_exe and Path(sp_exe).is_file():
            subprocess.Popen([sp_exe, project_file])
        else:
            subprocess.Popen(["xdg-open", project_file])
        return True
    except OSError:
        return False


def macos_app_executable(app_path):
    if not app_path:
        return None
    path = Path(app_path)
    if path.suffix.lower() != ".app":
        return None
    macos_dir = path / "Contents" / "MacOS"
    if not macos_dir.is_dir():
        return None
    preferred = macos_dir / path.stem
    if preferred.is_file():
        return preferred
    for candidate in macos_dir.iterdir():
        if candidate.is_file():
            return candidate
    return None


def launch_sp_instance(sp_exe=None, new_instance=False, force_token=None):
    if not sp_exe:
        return False
    try:
        env = None
        token = str(force_token or "").strip()
        if token:
            env = os.environ.copy()
            env["GOB_SP_FORCE_NEW_TOKEN"] = token
        if sys.platform == "darwin":
            if sp_exe.lower().endswith(".app"):
                if token:
                    exec_path = macos_app_executable(sp_exe)
                    if exec_path:
                        subprocess.Popen([str(exec_path)], env=env)
                        return True
                cmd = ["open"]
                if new_instance:
                    cmd.append("-n")
                cmd.extend(["-a", sp_exe])
                subprocess.Popen(cmd, env=env)
                return True
        if os.name == "nt" and new_instance:
            cmd = ["cmd", "/c", "start", "", sp_exe]
            subprocess.Popen(cmd, env=env)
            return True
        cmd = [sp_exe]
        subprocess.Popen(cmd, env=env)
        return True
    except OSError:
        return False


def open_path_in_file_manager(path):
    if not path:
        return False
    try:
        bpy.ops.wm.path_open(filepath=str(path))
        return True
    except Exception:
        pass
    try:
        if os.name == "nt" and hasattr(os, "startfile"):
            os.startfile(str(path))
            return True
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
            return True
        subprocess.Popen(["xdg-open", str(path)])
        return True
    except Exception:
        return False


def is_sp_running():
    try:
        output = subprocess.check_output(
            ["tasklist", "/FO", "CSV"],
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except OSError:
        return False
    return ("Adobe Substance 3D Painter.exe" in output or
            "Substance 3D Painter.exe" in output)


_update_check_in_progress = False
_update_check_result = None
_update_check_show_no_update = False
_update_check_show_popup = False
_last_update_info = None
_update_status_kind = "idle"
_update_status_text = "Update: not checked yet"
_update_status_time = 0.0
_last_export_warning = ""
_pending_export_popup = False
_cache_size_check_time = 0.0
_cache_size_global = 0
_cache_size_local = 0
_cache_size_project_root = None


def _set_update_status(kind, text, info=None):
    global _update_status_kind
    global _update_status_text
    global _update_status_time
    global _last_update_info
    _update_status_kind = kind
    _update_status_text = text
    _update_status_time = time.time()
    if info:
        _last_update_info = info
    elif kind != "update":
        _last_update_info = None


def _update_worker():
    global _update_check_result
    _update_check_result = check_for_updates()


def _show_update_popup(info):
    if not info:
        return
    global _last_update_info
    _last_update_info = info
    wm = bpy.context.window_manager
    if not wm:
        return

    def draw(self, _context):
        layout = self.layout
        layout.label(
            text=f"Update available: {info['version']} (current {info['local_version']})"
        )
        notes = info.get("notes")
        if notes:
            for line in str(notes).splitlines():
                if line.strip():
                    layout.label(text=line.strip())
        if info.get("download_url"):
            layout.operator(GOB_OT_OpenUpdateURL.bl_idname, text="Open Download Page")

    wm.popup_menu(draw, title="GoB SP Bridge Update", icon="INFO")


def _show_simple_popup(title, message, icon="INFO"):
    wm = bpy.context.window_manager
    if not wm:
        return

    def draw(self, _context):
        layout = self.layout
        for line in str(message).splitlines():
            if line.strip():
                layout.label(text=line.strip())

    wm.popup_menu(draw, title=title, icon=icon)


def _set_export_warning(message):
    global _last_export_warning
    _last_export_warning = message or ""


def _queue_export_warning_popup(message):
    global _pending_export_popup
    if _pending_export_popup:
        return
    _pending_export_popup = True

    def _show_popup():
        global _pending_export_popup
        _pending_export_popup = False
        if message:
            _show_simple_popup("GoB SP Bridge", message, icon="ERROR")
        return None

    bpy.app.timers.register(_show_popup, first_interval=0.01)


def _enforce_selected_suffix_policy(context, prefs, operator=None):
    if not (prefs and prefs.export_selected_only and prefs.export_low_poly and prefs.export_high_poly):
        return True
    scene = context.scene
    if getattr(scene, "gob_sp_low_poly_collection", None) or getattr(scene, "gob_sp_high_poly_collection", None):
        _set_export_warning("")
        return True
    selected_meshes = [obj for obj in context.selected_objects if obj.type == "MESH"]
    if not selected_meshes:
        return True
    if prefs.experimental_auto_split_selected:
        low_objects, high_objects = split_meshes_by_triangles(selected_meshes)
        if low_objects and high_objects:
            _set_export_warning("")
            return True
        prefs.export_high_poly = False
        message = (
            "Experimental auto-split needs both low and high meshes in the selection "
            "(different triangle counts). High poly export was turned off."
        )
        _set_export_warning(message)
        if operator:
            operator.report({"WARNING"}, message)
        _queue_export_warning_popup(message)
        return False
    low_suffixes = parse_suffixes(prefs.low_poly_suffixes) or ["_low"]
    high_suffixes = parse_suffixes(prefs.high_poly_suffixes) or ["_high"]
    has_low = False
    has_high = False
    has_unknown = False
    for obj in selected_meshes:
        is_low = is_name_with_suffix(obj.name, low_suffixes)
        is_high = is_name_with_suffix(obj.name, high_suffixes)
        if is_low:
            has_low = True
        if is_high:
            has_high = True
        if not (is_low or is_high):
            has_unknown = True
    if has_low and has_high and not has_unknown:
        _set_export_warning("")
        return True
    prefs.export_high_poly = False
    message = (
        "Export Selected Only requires selected meshes named with low/high suffixes "
        "(for example: _low and _high). High poly export was turned off."
    )
    _set_export_warning(message)
    if operator:
        operator.report({"WARNING"}, message)
    _queue_export_warning_popup(message)
    return False


def _on_export_selected_only_update(self, context):
    if not self.export_selected_only:
        if self.experimental_auto_split_selected:
            self.experimental_auto_split_selected = False
        _set_export_warning("")
        return
    _enforce_selected_suffix_policy(context, self)


def _on_export_low_poly_update(self, context):
    if not self.export_low_poly:
        if self.experimental_auto_split_selected:
            self.experimental_auto_split_selected = False
        _set_export_warning("")
        return
    if self.experimental_auto_split_selected:
        _enforce_selected_suffix_policy(context, self)


def _on_export_high_poly_update(self, context):
    if not self.export_high_poly:
        if self.experimental_auto_split_selected:
            self.experimental_auto_split_selected = False
        _set_export_warning("")
        return
    _enforce_selected_suffix_policy(context, self)


def _on_experimental_auto_split_update(self, context):
    if not self.experimental_auto_split_selected:
        if self.export_low_poly:
            self.export_low_poly = False
        if self.export_high_poly:
            self.export_high_poly = False
        if self.export_selected_only:
            self.export_selected_only = False
        _set_export_warning("")
        return
    if not self.export_selected_only:
        self.export_selected_only = True
    if not self.export_low_poly:
        self.export_low_poly = True
    if not self.export_high_poly:
        self.export_high_poly = True
    _enforce_selected_suffix_policy(context, self)


def _on_auto_clear_cache_update(self, _context):
    if not self.auto_clear_cache:
        return
    limit = getattr(self, "cache_limit_gb", DEFAULT_CACHE_LIMIT_GB)
    message = (
        "Warning: auto-clear removes cached projects (keeps the current project) "
        f"when total cache exceeds {limit:.1f} GB."
    )
    _show_simple_popup("GoB SP Bridge", message)


def _update_poll():
    global _update_check_in_progress
    global _update_check_result
    global _update_check_show_no_update
    global _update_check_show_popup
    if _update_check_result is None:
        return 0.5
    result = _update_check_result
    _update_check_result = None
    _update_check_in_progress = False
    if result.get("status") == "update":
        info = result.get("info")
        _set_update_status("update", f"Update available: {info['version']}", info=info)
        if _update_check_show_popup:
            _show_update_popup(info)
    elif _update_check_show_no_update:
        if result.get("status") == "none":
            local = result.get("local_version") or local_version_string()
            _set_update_status("up_to_date", f"Up to date ({local})")
            _show_simple_popup("GoB SP Bridge", f"You're up to date ({local}).")
        else:
            error = result.get("error") or "Update check failed."
            _set_update_status("error", f"Update check failed: {error}")
            _show_simple_popup("GoB SP Bridge", error, icon="ERROR")
    elif result.get("status") == "none":
        local = result.get("local_version") or local_version_string()
        _set_update_status("up_to_date", f"Up to date ({local})")
    else:
        error = result.get("error") or "Update check failed."
        _set_update_status("error", f"Update check failed: {error}")
    _update_check_show_no_update = False
    _update_check_show_popup = False
    return None


def start_update_check(show_no_update=False, show_popup=True):
    global _update_check_in_progress
    global _update_check_show_no_update
    global _update_check_show_popup
    if _update_check_in_progress:
        return
    _update_check_in_progress = True
    _update_check_show_no_update = show_no_update
    _update_check_show_popup = show_popup
    _set_update_status("checking", "Update: checking...")
    thread = threading.Thread(target=_update_worker, daemon=True)
    thread.start()
    bpy.app.timers.register(_update_poll, first_interval=0.5)


def get_cached_cache_sizes(context, prefs, max_age=5.0):
    global _cache_size_check_time
    global _cache_size_global
    global _cache_size_local
    global _cache_size_project_root
    now = time.time()
    project_dir = str(get_project_dir(context, prefs))
    if (
        _cache_size_project_root != project_dir
        or now - _cache_size_check_time > max_age
    ):
        _cache_size_project_root = project_dir
        _cache_size_global = bridge_cache_size_bytes(prefs)
        _cache_size_local = project_cache_size_bytes(context, prefs)
        if prefs and getattr(prefs, "auto_clear_cache", False):
            limit_bytes = cache_limit_bytes(prefs)
            if limit_bytes and _cache_size_global > limit_bytes:
                keep_paths = [get_project_dir(context, prefs)]
                result = clear_cache_dir_except(get_bridge_root(prefs), keep_paths=keep_paths)
                if result == "cleared":
                    _cache_size_global = bridge_cache_size_bytes(prefs)
                    _cache_size_local = project_cache_size_bytes(context, prefs)
        _cache_size_check_time = now
    return _cache_size_global, _cache_size_local


def refresh_cache_sizes(context, prefs):
    global _cache_size_check_time
    global _cache_size_global
    global _cache_size_local
    global _cache_size_project_root
    project_dir = str(get_project_dir(context, prefs))
    _cache_size_project_root = project_dir
    _cache_size_global = bridge_cache_size_bytes(prefs)
    _cache_size_local = project_cache_size_bytes(context, prefs)
    _cache_size_check_time = time.time()


def _init_scene_ui_prefs(_context=None):
    prefs = get_prefs(bpy.context) if bpy.context else None
    default_show = prefs.ui_show_export_settings if prefs else True
    for scene in bpy.data.scenes:
        if not getattr(scene, "gob_sp_ui_export_settings_initialized", False):
            scene.gob_sp_ui_show_export_settings = default_show
            scene.gob_sp_ui_export_settings_initialized = True
    _update_active_blender_info()
    return None


class GOBSPPreferences(AddonPreferences):
    bl_idname = __name__

    bridge_dir: StringProperty(
        name="Bridge Folder",
        subtype="DIR_PATH",
        default=default_bridge_dir(),
    )
    auto_launch_sp: BoolProperty(
        name="Auto-launch Substance Painter",
        default=True,
    )
    open_linked_sp_project: BoolProperty(
        name="Open Linked SP Project",
        description="Open the linked .spp file when sending to Substance Painter",
        default=False,
    )
    force_new_sp_project_on_send: BoolProperty(
        name="Open New Painter Instance",
        description=(
            "Launch a new Substance Painter instance and create a new project"
            " instead of reusing the current one"
        ),
        default=False,
    )
    export_high_poly: BoolProperty(
        name="Export High Poly",
        default=True,
        update=_on_export_high_poly_update,
    )
    export_low_poly: BoolProperty(
        name="Export Low Poly",
        default=True,
        update=_on_export_low_poly_update,
    )
    export_selected_only: BoolProperty(
        name="Only Selected Meshes",
        description="Limit low/high exports to the current selection",
        default=False,
        update=_on_export_selected_only_update,
    )
    experimental_auto_split_selected: BoolProperty(
        name="Auto-split by Triangle Count",
        description="Split selected meshes into low/high using triangle counts",
        default=False,
        update=_on_experimental_auto_split_update,
    )
    low_poly_suffixes: StringProperty(
        name="Low Poly Suffixes",
        description="Comma-separated suffixes for low poly objects (must be at end)",
        default="_low",
    )
    high_poly_suffixes: StringProperty(
        name="High Poly Suffixes",
        description="Comma-separated suffixes for high poly objects (must be at end)",
        default="_high",
    )
    fbx_export_scale: FloatProperty(
        name="FBX Export Scale",
        default=1.0,
        min=0.001,
        max=1000.0,
    )
    fbx_apply_unit_scale: BoolProperty(
        name="Apply Unit Scale",
        default=True,
    )
    fbx_export_custom_normals: BoolProperty(
        name="Export Custom Normals",
        default=True,
    )
    ui_show_export_settings: BoolProperty(
        name="Show Export Settings",
        default=True,
    )
    ui_show_project_link: BoolProperty(
        name="Show Project Link",
        default=True,
    )
    ui_show_fbx_settings: BoolProperty(
        name="Show FBX Export Settings",
        default=False,
    )
    ui_show_cache: BoolProperty(
        name="Show Cache",
        default=False,
    )
    auto_clear_cache: BoolProperty(
        name="Auto-clear Global Cache",
        description="Remove cached projects (keeps the current one) when over the limit",
        default=False,
        update=_on_auto_clear_cache_update,
    )
    cache_limit_gb: FloatProperty(
        name="Global Cache Limit (GB)",
        default=DEFAULT_CACHE_LIMIT_GB,
        min=1.0,
        max=2048.0,
    )

    def draw(self, _context):
        layout = self.layout
        layout.label(text="Bridge")
        layout.prop(self, "bridge_dir")
        layout.prop(self, "auto_launch_sp")
        layout.separator()
        layout.label(text="Use the GoB SP panel for export options")


class GOB_OT_SendToSP(Operator):
    bl_idname = "gob_sp.send_to_substance_painter"
    bl_label = "Send to Substance Painter"

    def execute(self, context):
        prefs = get_prefs(context)
        write_active_blender_info(context, prefs)
        blender_file = get_blender_file_path_or_temp(prefs)
        force_new_project = bool(prefs and prefs.force_new_sp_project_on_send)
        _enforce_selected_suffix_policy(context, prefs, operator=self)
        if prefs and prefs.export_selected_only and prefs.experimental_auto_split_selected:
            selected_meshes = [obj for obj in context.selected_objects if obj.type == "MESH"]
            low_objects, high_candidates = split_meshes_by_triangles(selected_meshes)
        else:
            low_objects = collect_low_poly_objects(context, prefs)
            high_candidates = []
            if prefs and prefs.export_high_poly:
                high_candidates = collect_high_poly_candidates(context, prefs)
        if prefs and prefs.export_high_poly and high_candidates and low_objects:
            high_names = {obj.name for obj in high_candidates}
            low_objects = [obj for obj in low_objects if obj.name not in high_names]
        if not low_objects and (not prefs or prefs.export_low_poly):
            self.report({"ERROR"}, "Select or name at least one low poly mesh")
            return {"CANCELLED"}

        if low_objects and any(not object_has_uvs(obj) for obj in low_objects):
            self.report({"ERROR"}, "Missing UVs: unwrap in Blender before export")
            return {"CANCELLED"}

        high_signature_objects = []
        if prefs and prefs.export_high_poly:
            high_signature_objects = high_candidates
        mesh_signature = build_mesh_signature(low_objects, high_signature_objects)
        signature_manifest = None
        signature_project_dir = None
        signature_sp_project = ""
        if blender_file and mesh_signature and not force_new_project:
            signature_manifest_path = find_manifest_for_mesh_signature(
                get_candidate_bridge_roots(prefs),
                blender_file,
                mesh_signature,
                source="blender",
            )
            if signature_manifest_path:
                signature_project_dir = project_dir_from_manifest_path(signature_manifest_path)
                signature_manifest = read_manifest(signature_manifest_path)
                signature_sp_project = get_manifest_sp_project_file(signature_manifest)

        active_info = resolve_active_sp_project_info(context, prefs)
        if force_new_project:
            active_info = None
        if signature_project_dir:
            project_dir = signature_project_dir
            if active_info and active_info.get("project_dir") != project_dir:
                active_info = None
        else:
            if active_info and not project_dir_signature_matches(active_info.get("project_dir"), mesh_signature):
                active_info = None
            if force_new_project:
                base_dir = get_bridge_root(prefs) / get_project_name(context)
                project_dir = unique_project_dir(base_dir, None, prefs)
            else:
                project_dir = (
                    active_info["project_dir"]
                    if active_info
                    else project_dir_for_send(context, prefs, blender_file)
                )
        write_bridge_root_hint(project_dir.parent)
        ensure_dir(project_dir)

        export_path = project_dir / BLENDER_EXPORT_FILENAME
        old_manifest = read_manifest(find_project_manifest_path(project_dir))
        old_mesh = old_manifest.get("mesh_fbx") if old_manifest else None
        linked_sp_project_hint = signature_sp_project or get_linked_sp_project_path(
            project_dir,
            active_info=active_info,
            blender_file=blender_file,
            prefs=prefs,
        )
        if signature_sp_project:
            linked_sp_project = resolve_sp_project_candidate(
                signature_sp_project,
                blender_file,
                prefs=prefs,
            )
        else:
            linked_sp_project = resolve_linked_sp_project_file(
                project_dir,
                active_info=active_info,
                blender_file=blender_file,
                prefs=prefs,
            )
        if force_new_project:
            linked_sp_project_hint = ""
            linked_sp_project = ""
        sp_project_file = ""
        if active_info:
            sp_project_file = str(active_info.get("sp_project_file") or "")
        if not sp_project_file:
            sp_project_file = signature_sp_project or linked_sp_project
        if sp_project_file and blender_file and not force_new_project:
            update_link_registry(sp_project_file=sp_project_file, blender_file=blender_file, prefs=prefs)

        if not prefs or prefs.export_low_poly:
            if not low_objects:
                self.report({"ERROR"}, "Low poly export enabled but no meshes found")
                return {"CANCELLED"}
            strip_uvs = False
            export_fbx_objects(
                export_path,
                low_objects,
                prefs=prefs,
                strip_uvs=strip_uvs,
            )
        elif not old_mesh:
            self.report({"ERROR"}, "Low poly export disabled and no previous low mesh found")
            return {"CANCELLED"}

        high_export_path = None
        if prefs and prefs.export_high_poly:
            high_objects = high_candidates
            if high_objects:
                high_export_path = project_dir / BLENDER_HIGH_FILENAME
                exported = export_fbx_objects(high_export_path, high_objects, prefs=prefs)
                if not exported or not high_export_path.exists():
                    self.report({"WARNING"}, "High poly export failed or produced no FBX")
                    high_export_path = None

        force_new_token = ""
        if force_new_project:
            force_new_token = uuid.uuid4().hex

        manifest_path = project_manifest_path(project_dir)
        if manifest_path:
            ensure_dir(manifest_path.parent)
        sp_running = is_sp_running()
        manifest = {
            "version": 1,
            "source": "blender",
            "project": get_project_name(context),
            "mesh_fbx": str(export_path) if (not prefs or prefs.export_low_poly) else old_mesh,
            "timestamp": time.time(),
        }
        if mesh_signature:
            manifest["mesh_signature"] = mesh_signature
        if blender_file:
            manifest["blender_file"] = blender_file
        else:
            previous_blender_file = get_manifest_blender_file(old_manifest)
            if previous_blender_file:
                manifest["blender_file"] = previous_blender_file
        if linked_sp_project_hint:
            manifest["sp_project_file"] = linked_sp_project_hint
        if force_new_project:
            manifest["force_new_project"] = True
            if force_new_token:
                manifest["force_new_token"] = force_new_token
        manifest["auto_import"] = True
        manifest["auto_import_at"] = time.time()
        if high_export_path:
            manifest["high_mesh_fbx"] = str(high_export_path)
        if prefs and prefs.export_high_poly:
            manifest["high_mesh_exported"] = bool(high_export_path)
        write_manifest(manifest_path, manifest)

        sp_exe = find_sp_exe(prefs) if prefs else None
        opened_project = False
        active_any = find_active_sp_project_info(prefs)
        active_sp_file = ""
        if active_any:
            active_sp_file = str(active_any.get("sp_project_file") or "")
        already_open = bool(
            sp_running
            and linked_sp_project
            and active_sp_file
            and paths_match(active_sp_file, linked_sp_project)
        )
        should_force_open = bool(
            sp_running
            and linked_sp_project
            and active_sp_file
            and not paths_match(active_sp_file, linked_sp_project)
        )
        if force_new_project:
            if sp_exe:
                opened_project = launch_sp_instance(
                    sp_exe,
                    new_instance=True,
                    force_token=force_new_token,
                )
                if not opened_project:
                    self.report({"WARNING"}, "Failed to launch Substance Painter")
            else:
                self.report({"WARNING"}, "Substance Painter executable not found")
        else:
            if (linked_sp_project and not already_open and
                    (should_force_open or (prefs and prefs.open_linked_sp_project))):
                if is_temp_sp_project_file(linked_sp_project, prefs):
                    if not sp_running and sp_exe:
                        try:
                            if sys.platform == "darwin" and sp_exe.lower().endswith(".app"):
                                subprocess.Popen(["open", "-a", sp_exe])
                            else:
                                subprocess.Popen([sp_exe])
                            opened_project = True
                        except OSError:
                            self.report({"WARNING"}, "Failed to launch Substance Painter")
                else:
                    opened_project = open_sp_project_file(linked_sp_project, sp_exe=sp_exe)
                    if not opened_project:
                        self.report({"WARNING"}, "Failed to open linked Substance Painter project")

            if prefs and prefs.auto_launch_sp and not sp_running and not opened_project:
                if sp_exe:
                    try:
                        if sys.platform == "darwin" and sp_exe.lower().endswith(".app"):
                            subprocess.Popen(["open", "-a", sp_exe])
                        else:
                            subprocess.Popen([sp_exe])
                    except OSError:
                        self.report({"WARNING"}, "Failed to launch Substance Painter")
                else:
                    self.report({"WARNING"}, "Substance Painter executable not found")

        self.report({"INFO"}, "Exported FBX for Substance Painter")
        return {"FINISHED"}


class GOB_OT_ImportFromSP(Operator):
    bl_idname = "gob_sp.import_from_substance_painter"
    bl_label = "Import from Substance Painter"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        prefs = get_prefs(context)
        roots = get_candidate_bridge_roots(prefs)
        project_dir = get_project_dir(context, prefs)
        manifest_path = None
        manifest = None
        current_blender_file = get_blender_file_path_or_temp(prefs)
        current_is_temp = is_temp_blender_file(current_blender_file, prefs)
        active_info = resolve_active_sp_project_info(context, prefs)
        if not active_info and (not current_blender_file or current_is_temp):
            active_info = find_active_sp_project_info(prefs)
        sp_project_file = active_info.get("sp_project_file") if active_info else ""
        if sp_project_file:
            candidate = find_manifest_for_sp_project_file(
                roots,
                sp_project_file,
                source="substance_painter",
            )
            if candidate:
                manifest_path = candidate
                manifest = read_manifest(manifest_path)
                if manifest and current_blender_file:
                    manifest_blender = get_manifest_blender_file(manifest)
                    if manifest_blender and not paths_match(manifest_blender, current_blender_file):
                        manifest = None
                        manifest_path = None
        if not manifest or manifest.get("source") != "substance_painter":
            blender_file = current_blender_file
            if blender_file:
                candidate = find_manifest_for_blender_file(
                    roots,
                    blender_file,
                    source="substance_painter",
                )
                if candidate:
                    manifest_path = candidate
                    manifest = read_manifest(manifest_path)
        if not manifest or manifest.get("source") != "substance_painter":
            if not bpy.data.filepath:
                candidate = find_project_manifest_path(project_dir)
                if candidate and candidate.exists():
                    manifest_path = candidate
                    manifest = read_manifest(manifest_path)
        if not manifest or manifest.get("source") != "substance_painter":
            self.report({"ERROR"}, "No Substance Painter bridge manifest found for this project")
            return {"CANCELLED"}
        if not manifest:
            self.report({"ERROR"}, "Failed to read bridge manifest")
            return {"CANCELLED"}
        project_dir = project_dir_from_manifest_path(manifest_path)
        sp_project_file = get_manifest_sp_project_file(manifest)
        link_sp_project_file = get_manifest_link_sp_project_file(manifest)
        blender_file = get_manifest_blender_file(manifest) or current_blender_file
        if link_sp_project_file and blender_file:
            update_link_registry(
                sp_project_file=link_sp_project_file,
                blender_file=blender_file,
                prefs=prefs,
            )
        elif sp_project_file and blender_file:
            update_link_registry(
                sp_project_file=sp_project_file,
                blender_file=blender_file,
                prefs=prefs,
            )

        mesh_path = manifest.get("mesh_fbx")
        mesh_exported = bool(mesh_path) or bool(manifest.get("mesh_exported"))
        if mesh_path:
            mesh_path = Path(mesh_path)
            if not mesh_path.is_absolute():
                mesh_path = project_dir / mesh_path
            mesh_path = str(mesh_path)
        if not mesh_path and mesh_exported:
            fallback = project_dir / SP_EXPORT_FILENAME
            if fallback.exists():
                mesh_path = str(fallback)
        new_objects = []
        if mesh_path and Path(mesh_path).is_file():
            new_objects = import_fbx(mesh_path)

        texture_paths = gather_texture_paths(manifest)
        targets = list(new_objects)
        signature_targets = find_signature_targets(context, manifest)
        if signature_targets:
            existing = {obj.name for obj in targets}
            for obj in signature_targets:
                if obj.name not in existing:
                    targets.append(obj)
                    existing.add(obj.name)
        grouped = group_textures(texture_paths) if texture_paths else {}
        strict = False
        if grouped:
            matched_targets = find_texture_targets(context, grouped)
            if matched_targets:
                if targets:
                    existing = {obj.name for obj in targets}
                    for obj in matched_targets:
                        if obj.name not in existing:
                            targets.append(obj)
                            existing.add(obj.name)
                else:
                    targets = matched_targets
        if not targets and grouped:
            targets = find_texture_targets(context, grouped)
            if not targets:
                targets = [obj for obj in context.scene.objects if obj.type == "MESH"]
                strict = True
        if not targets and grouped:
            self.report(
                {"WARNING"},
                "No mesh targets found; match material or object names to texture sets",
            )
        if texture_paths and targets:
            apply_textures_to_objects(targets, grouped, manifest=manifest, strict=strict)

        self.report({"INFO"}, "Imported assets from Substance Painter")
        return {"FINISHED"}


class GOB_OT_OpenExportFolder(Operator):
    bl_idname = "gob_sp.open_export_folder"
    bl_label = "Open Export Folder"

    def execute(self, context):
        prefs = get_prefs(context)
        active_info = resolve_active_sp_project_info(context, prefs)
        target_dir = active_info["project_dir"] if active_info else get_project_dir(context, prefs)
        if not target_dir:
            self.report({"ERROR"}, "Export folder is not available")
            return {"CANCELLED"}
        ensure_dir(target_dir)
        if not open_path_in_file_manager(target_dir):
            self.report({"ERROR"}, "Failed to open export folder")
            return {"CANCELLED"}
        return {"FINISHED"}


class GOB_OT_ClearCacheGlobal(Operator):
    bl_idname = "gob_sp.clear_cache_global"
    bl_label = "Clear Global Cache"

    def execute(self, context):
        prefs = get_prefs(context)
        root = get_bridge_root(prefs)
        result = clear_cache_dir(root)
        if result == "empty":
            self.report({"INFO"}, "Global cache is already empty")
            refresh_cache_sizes(context, prefs)
            return {"FINISHED"}
        if result == "error":
            self.report({"WARNING"}, "Failed to clear global cache")
            refresh_cache_sizes(context, prefs)
            return {"CANCELLED"}
        self.report({"INFO"}, "Global cache cleared")
        refresh_cache_sizes(context, prefs)
        return {"FINISHED"}

    def invoke(self, context, _event):
        return context.window_manager.invoke_confirm(self, _event)


class GOB_OT_ClearCacheLocal(Operator):
    bl_idname = "gob_sp.clear_cache_local"
    bl_label = "Clear Project Cache"

    def execute(self, context):
        prefs = get_prefs(context)
        root = get_project_dir(context, prefs)
        result = clear_cache_dir(root)
        if result == "empty":
            self.report({"INFO"}, "Project cache is already empty")
            refresh_cache_sizes(context, prefs)
            return {"FINISHED"}
        if result == "error":
            self.report({"WARNING"}, "Failed to clear project cache")
            refresh_cache_sizes(context, prefs)
            return {"CANCELLED"}
        self.report({"INFO"}, "Project cache cleared")
        refresh_cache_sizes(context, prefs)
        return {"FINISHED"}

    def invoke(self, context, _event):
        return context.window_manager.invoke_confirm(self, _event)


class GOB_OT_OpenDiscord(Operator):
    bl_idname = "gob_sp.open_discord"
    bl_label = "Join Discord"

    def execute(self, _context):
        bpy.ops.wm.url_open(url=DISCORD_INVITE_URL)
        return {"FINISHED"}


class GOB_OT_OpenBugReport(Operator):
    bl_idname = "gob_sp.open_bug_report"
    bl_label = "Report Bug"

    def execute(self, _context):
        bpy.ops.wm.url_open(url=BUG_REPORT_URL)
        return {"FINISHED"}


class GOB_OT_CheckUpdates(Operator):
    bl_idname = "gob_sp.check_updates"
    bl_label = "Check for Updates"

    def execute(self, _context):
        start_update_check(show_no_update=True, show_popup=True)
        return {"FINISHED"}


class GOB_OT_OpenUpdateURL(Operator):
    bl_idname = "gob_sp.open_update_url"
    bl_label = "Open Update Download"

    def execute(self, _context):
        if not _last_update_info or not _last_update_info.get("download_url"):
            return {"CANCELLED"}
        bpy.ops.wm.url_open(url=_last_update_info["download_url"])
        return {"FINISHED"}


class GOB_PT_Panel(Panel):
    bl_label = "GoB SP Bridge"
    bl_idname = "GOB_PT_sp_bridge"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "GoB SP"

    def draw(self, context):
        layout = self.layout
        prefs = get_prefs(context)
        scene = context.scene
        active_info = None
        project_dir = ""
        blender_file = ""
        auto_sp_project = ""
        linked_sp_project = ""
        auto_exists = False
        auto_is_temp = False
        if prefs:
            global _ui_link_cache
            now = time.time()
            blender_file = get_blender_file_path_or_temp(prefs)
            project_dir = get_project_dir_fast(context, prefs)
            cache_ok = (
                now - _ui_link_cache.get("timestamp", 0.0) < UI_LINK_CACHE_TTL
                and _ui_link_cache.get("blender_file") == blender_file
                and _ui_link_cache.get("project_dir") == str(project_dir)
            )
            if cache_ok:
                active_info = _ui_link_cache.get("active_info")
                auto_sp_project = _ui_link_cache.get("auto_sp_project", "")
                linked_sp_project = _ui_link_cache.get("linked_sp_project", "")
                auto_exists = bool(_ui_link_cache.get("auto_exists"))
                auto_is_temp = bool(_ui_link_cache.get("auto_is_temp"))
            else:
                active_info = read_active_sp_info(
                    project_meta_dir(project_dir) / ACTIVE_SP_INFO_FILENAME
                )
                if not active_info:
                    active_info = find_active_sp_project_info(prefs)
                auto_sp_project = get_linked_sp_project_path_fast(
                    project_dir,
                    active_info=active_info,
                    blender_file=blender_file,
                    prefs=prefs,
                )
                linked_sp_project = resolve_linked_sp_project_file_fast(
                    project_dir,
                    active_info=active_info,
                    blender_file=blender_file,
                    prefs=prefs,
                )
                if auto_sp_project:
                    auto_is_temp = is_temp_sp_project_file(auto_sp_project, prefs)
                    if not auto_is_temp:
                        try:
                            auto_exists = Path(auto_sp_project).is_file()
                        except OSError:
                            auto_exists = False
                _ui_link_cache = {
                    "timestamp": now,
                    "blender_file": blender_file,
                    "project_dir": str(project_dir),
                    "active_info": active_info,
                    "auto_sp_project": auto_sp_project,
                    "linked_sp_project": linked_sp_project,
                    "auto_is_temp": auto_is_temp,
                    "auto_exists": auto_exists,
                }
        show_export = getattr(scene, "gob_sp_ui_show_export_settings", True)
        row = layout.row(align=True)
        row.operator(GOB_OT_SendToSP.bl_idname, icon="EXPORT")
        row.operator(GOB_OT_ImportFromSP.bl_idname, icon="IMPORT")
        layout.operator(GOB_OT_OpenExportFolder.bl_idname, icon="FILE_FOLDER")
        if prefs:
            export_box = layout.box()
            row = export_box.row()
            icon = "TRIA_DOWN" if show_export else "TRIA_RIGHT"
            row.prop(scene, "gob_sp_ui_show_export_settings", icon=icon,
                     emboss=False, text="Send to Painter")
            if show_export:
                scope_box = export_box.box()
                scope_box.label(text="Mesh Selection")
                scope_col = scope_box.column(align=True)
                scope_col.prop(prefs, "export_selected_only", text="Only Selected")
                scope_col.prop(prefs, "export_low_poly", text="Export Low")
                scope_col.prop(prefs, "export_high_poly", text="Export High")
                scope_col.prop(
                    prefs,
                    "experimental_auto_split_selected",
                    text="Auto-split Selected (experimental)",
                )
                if prefs.export_high_poly:
                    id_box = export_box.box()
                    id_box.label(text="Low/High Identification")
                    if prefs.export_selected_only and prefs.experimental_auto_split_selected:
                        info = id_box.box()
                        info.label(text="Auto-split uses triangle counts", icon="INFO")
                        info.label(text="Lower triangle meshes export as low")
                        info.label(text="Higher triangle meshes export as high")
                    else:
                        id_col = id_box.column(align=True)
                        id_col.prop_search(
                            scene,
                            "gob_sp_low_poly_collection",
                            bpy.data,
                            "collections",
                            text="Low Collection",
                        )
                        id_col.prop_search(
                            scene,
                            "gob_sp_high_poly_collection",
                            bpy.data,
                            "collections",
                            text="High Collection",
                        )
                        and_or_row = id_col.row()
                        and_or_row.alignment = "CENTER"
                        and_or_row.label(text="AND/OR")
                        id_col.prop(prefs, "low_poly_suffixes")
                        id_col.prop(prefs, "high_poly_suffixes")
                        info = id_box.box()
                        info.label(text="Collections override suffix matching", icon="INFO")
                        info.label(text="SP bake expects _low/_high when matching by name", icon="INFO")
                elif prefs.export_low_poly:
                    id_box = export_box.box()
                    id_box.label(text="Low Identification")
                    id_col = id_box.column(align=True)
                    id_col.prop_search(
                        scene,
                        "gob_sp_low_poly_collection",
                        bpy.data,
                        "collections",
                        text="Low Collection",
                    )
                    and_or_row = id_col.row()
                    and_or_row.alignment = "CENTER"
                    and_or_row.label(text="AND/OR")
                    id_col.prop(prefs, "low_poly_suffixes")
                    info = id_box.box()
                    info.label(text="Collection overrides suffix matching", icon="INFO")
                    info.label(text="SP bake expects _low/_high when matching by name", icon="INFO")

                link_box = export_box.box()
                row = link_box.row()
                icon = "TRIA_DOWN" if prefs.ui_show_project_link else "TRIA_RIGHT"
                row.prop(prefs, "ui_show_project_link", icon=icon, emboss=False, text="Project Link")
                if prefs.ui_show_project_link:
                    if auto_sp_project and auto_is_temp:
                        link_box.label(text="Linked SP project is unsaved", icon="INFO")
                        link_box.label(text=f"Detected: {auto_sp_project}", icon="INFO")
                    elif auto_sp_project and not auto_exists:
                        link_box.label(text="Linked SP project not found", icon="INFO")
                        link_box.label(text=f"Detected: {auto_sp_project}", icon="INFO")
                    elif auto_sp_project:
                        link_box.label(text=f"Detected: {auto_sp_project}", icon="INFO")
                    else:
                        link_box.label(text="No linked SP project detected", icon="INFO")
                    link_toggle = link_box.row()
                    link_toggle.enabled = bool(linked_sp_project)
                    link_toggle.prop(prefs, "open_linked_sp_project", text="Open linked project on send")
                    link_force = link_box.row()
                    link_force.prop(
                        prefs,
                        "force_new_sp_project_on_send",
                        text="Open new Painter instance",
                    )

            fbx_box = layout.box()
            row = fbx_box.row()
            icon = "TRIA_DOWN" if prefs.ui_show_fbx_settings else "TRIA_RIGHT"
            row.prop(prefs, "ui_show_fbx_settings", icon=icon, emboss=False, text="FBX Export Settings")
            if prefs.ui_show_fbx_settings:
                col = fbx_box.column(align=True)
                col.prop(prefs, "fbx_export_scale")
                col.prop(prefs, "fbx_apply_unit_scale")
                col.prop(prefs, "fbx_export_custom_normals")
                fbx_box.label(text="Tip: if triangles too small, raise Export Scale")

            cache_box = layout.box()
            cache_refresh = bool(prefs.ui_show_cache or prefs.auto_clear_cache)
            cache_max_age = 5.0 if prefs.auto_clear_cache else 30.0
            if cache_refresh:
                global_size, local_size = get_cached_cache_sizes(
                    context,
                    prefs,
                    max_age=cache_max_age,
                )
            else:
                global_size, local_size = _cache_size_global, _cache_size_local
            warn_size = format_bytes(CACHE_WARN_BYTES)
            limit_bytes = cache_limit_bytes(prefs) if prefs.auto_clear_cache else 0
            cache_label = "Cache"
            if limit_bytes and global_size >= limit_bytes:
                cache_label = f"Cache (over {format_bytes(limit_bytes)})"
            elif max(global_size, local_size) >= CACHE_WARN_BYTES:
                cache_label = f"Cache (over {warn_size})"
            row = cache_box.row()
            icon = "TRIA_DOWN" if prefs.ui_show_cache else "TRIA_RIGHT"
            row.prop(prefs, "ui_show_cache", icon=icon, emboss=False, text=cache_label)
            if prefs.ui_show_cache:
                cache_box.label(text=f"Global cache: {format_bytes(global_size)}")
                cache_box.label(text=f"Project cache: {format_bytes(local_size)}")
                row = cache_box.row()
                row.prop(prefs, "auto_clear_cache")
                row = cache_box.row()
                row.enabled = prefs.auto_clear_cache
                row.prop(prefs, "cache_limit_gb")
                if prefs.auto_clear_cache:
                    cache_box.label(text="Auto-clear keeps the current project", icon="INFO")
                row = cache_box.row(align=True)
                row.operator(GOB_OT_ClearCacheGlobal.bl_idname, icon="TRASH")
                row.operator(GOB_OT_ClearCacheLocal.bl_idname, icon="TRASH")

            links = layout.box()
            links.label(text="Community")
            update_row = links.row(align=True)
            update_row.label(text=_update_status_text)
            update_row.operator(GOB_OT_CheckUpdates.bl_idname, text="Check")
            if _last_update_info and _last_update_info.get("download_url"):
                update_row.operator(GOB_OT_OpenUpdateURL.bl_idname, text="Download")
            link_row = links.row(align=True)
            link_row.operator(GOB_OT_OpenDiscord.bl_idname, icon="URL")
            link_row.operator(GOB_OT_OpenBugReport.bl_idname, icon="URL")


classes = (
    GOBSPPreferences,
    GOB_OT_SendToSP,
    GOB_OT_ImportFromSP,
    GOB_OT_OpenExportFolder,
    GOB_OT_ClearCacheGlobal,
    GOB_OT_ClearCacheLocal,
    GOB_OT_OpenDiscord,
    GOB_OT_OpenBugReport,
    GOB_OT_CheckUpdates,
    GOB_OT_OpenUpdateURL,
    GOB_PT_Panel,
)


def register():
    for cls in classes:
        bpy.utils.register_class(cls)
    bpy.types.Scene.gob_sp_ui_show_export_settings = BoolProperty(
        name="Show Export Settings",
        default=True,
    )
    bpy.types.Scene.gob_sp_ui_export_settings_initialized = BoolProperty(
        name="Export Settings Initialized",
        default=False,
        options={"HIDDEN"},
    )
    bpy.types.Scene.gob_sp_low_poly_collection = PointerProperty(
        name="Low Poly Collection",
        description="Collection to export as low poly (overrides suffix matching)",
        type=bpy.types.Collection,
        poll=_scene_collection_poll,
    )
    bpy.types.Scene.gob_sp_high_poly_collection = PointerProperty(
        name="High Poly Collection",
        description="Collection to export as high poly (overrides suffix matching)",
        type=bpy.types.Collection,
        poll=_scene_collection_poll,
    )
    if _init_scene_ui_prefs not in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.append(_init_scene_ui_prefs)
    if _update_active_blender_info not in bpy.app.handlers.save_post:
        bpy.app.handlers.save_post.append(_update_active_blender_info)
    if not bpy.app.timers.is_registered(_init_scene_ui_prefs):
        bpy.app.timers.register(_init_scene_ui_prefs, first_interval=0.1)
    if not bpy.app.timers.is_registered(_active_blender_heartbeat):
        bpy.app.timers.register(_active_blender_heartbeat, first_interval=1.0)
    start_update_check(show_popup=False)


def unregister():
    if _init_scene_ui_prefs in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.remove(_init_scene_ui_prefs)
    if _update_active_blender_info in bpy.app.handlers.save_post:
        bpy.app.handlers.save_post.remove(_update_active_blender_info)
    if bpy.app.timers.is_registered(_init_scene_ui_prefs):
        bpy.app.timers.unregister(_init_scene_ui_prefs)
    if bpy.app.timers.is_registered(_active_blender_heartbeat):
        bpy.app.timers.unregister(_active_blender_heartbeat)
    if hasattr(bpy.types.Scene, "gob_sp_ui_export_settings_initialized"):
        del bpy.types.Scene.gob_sp_ui_export_settings_initialized
    if hasattr(bpy.types.Scene, "gob_sp_ui_show_export_settings"):
        del bpy.types.Scene.gob_sp_ui_show_export_settings
    if hasattr(bpy.types.Scene, "gob_sp_high_poly_collection"):
        del bpy.types.Scene.gob_sp_high_poly_collection
    if hasattr(bpy.types.Scene, "gob_sp_low_poly_collection"):
        del bpy.types.Scene.gob_sp_low_poly_collection
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
