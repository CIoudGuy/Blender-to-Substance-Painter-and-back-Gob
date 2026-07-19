bl_info = {
    "name": "GoB SP Bridge",
    "author": "Cloud Guy | cloud_was_taken on Discord",
    "version": (1, 0, 0),
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
import urllib.parse
import urllib.request
from pathlib import Path

import bpy
from bpy.app.handlers import persistent
from bpy.props import BoolProperty, EnumProperty, FloatProperty, StringProperty, PointerProperty
from bpy.types import AddonPreferences, Operator, Panel
from mathutils import Vector


BRIDGE_ENV_VAR = "GOB_SP_BRIDGE_DIR"
BRIDGE_ROOT_HINT_FILENAME = "bridge_root.json"
BRIDGE_SHARED_HINT_DIRNAME = ".gob_sp_bridge"
MANIFEST_FILENAME = "bridge.json"
BLENDER_EXPORT_FILENAME = "b2sp.fbx"
BLENDER_HIGH_FILENAME = "b2sp_hi.fbx"
BLENDER_CAGE_FILENAME = "b2sp_cage.fbx"
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
    "https://files.devalt.cloud/api/app/blender-to-substance-painter-and-back-gob"
)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".tga", ".exr"}
CACHE_WARN_BYTES = 35 * 1024 ** 3
DEFAULT_CACHE_LIMIT_GB = 35.0
UI_LINK_CACHE_TTL = 3.0
OBJECT_PROJECT_KEY_PROP = "gob_sp_project_key"
_temp_session_id = None
_temp_blender_file = None
_last_blender_file = None
_blender_session_id = uuid.uuid4().hex[:12]
_bridge_conflict_info = None
_project_dir_cache = {}
_ui_link_cache = {
    "timestamp": 0.0,
    "blender_file": "",
    "project_dir": "",
    "active_info": None,
    "linked_sp_project": "",
    "sp_running": None,
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
    return str(documents_bridge_root())


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
    normalized = normalize_path(path)
    return normalized.lower() if os.name == "nt" else normalized


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
        return normalize_path_key(path_obj.parent) == normalize_path_key(bridge_temp_dir(prefs))
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
        key = normalize_path_key(root)
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
        write_json_atomic(primary, data)
    except OSError:
        return
    for path in paths[1:]:
        if not path.exists():
            continue
        try:
            write_json_atomic(path, data)
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
    return normalize_path_key(left) == normalize_path_key(right)


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
        if selected_only and selected_names is not None and obj.name not in selected_names:
            continue
        if obj.name in seen:
            continue
        seen.add(obj.name)
        results.append(obj)
    return results


def get_identify_mode(scene):
    return getattr(scene, "gob_sp_identify_mode", "SUFFIXES")


def collect_low_poly_objects(context, prefs):
    scene = context.scene
    selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
    selected_names = None
    if selected_only:
        selected_names = {
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        }
    if get_identify_mode(scene) == "COLLECTIONS":
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
    return [obj for obj in search_pool if obj.type == "MESH"]


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


def write_json_atomic(path, data):
    path = Path(path)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, ensure_ascii=True)
        os.replace(temp_path, path)
    except OSError:
        try:
            temp_path.unlink()
        except OSError:
            pass
        raise


def write_manifest(path, data):
    write_json_atomic(path, data)


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
    try:
        roots.append(Path(get_bridge_root(prefs)))
    except (TypeError, ValueError):
        pass
    docs_root = documents_bridge_root()
    if docs_root:
        roots.append(Path(docs_root))
    if project_dir:
        roots.append(project_meta_dir(project_dir))
    unique = []
    seen = set()
    for root in roots:
        key = normalize_path_key(root)
        if key in seen:
            continue
        seen.add(key)
        unique.append(root / ACTIVE_BLENDER_INFO_FILENAME)
    return unique


def write_active_blender_info(context=None, prefs=None):
    global _bridge_conflict_info
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
        "session_id": _blender_session_id,
        "pid": os.getpid(),
    }
    blender_file = get_blender_file_path_or_temp(prefs)
    if blender_file:
        info["blender_file"] = blender_file
    for path in active_blender_info_paths(prefs, project_dir):
        try:
            ensure_dir(path.parent)
            if _bridge_conflict_info is None:
                previous = None
                try:
                    if path.is_file():
                        with open(path, "r", encoding="utf-8") as handle:
                            previous = json.load(handle)
                except Exception:
                    previous = None
                if isinstance(previous, dict):
                    previous_session = previous.get("session_id")
                    previous_timestamp = previous.get("timestamp")
                    try:
                        previous_age = time.time() - float(previous_timestamp)
                    except (TypeError, ValueError):
                        previous_age = None
                    if (
                        previous_session
                        and previous_session != _blender_session_id
                        and previous_age is not None
                        and previous_age < 45.0
                    ):
                        _bridge_conflict_info = {
                            "session_id": previous_session,
                            "pid": previous.get("pid"),
                            "timestamp": previous_timestamp,
                        }
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


def sync_saved_blender_file(context=None, prefs=None, after_save=False):
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
        if after_save:
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


@persistent
def _update_active_blender_info(_context=None):
    try:
        context = bpy.context
        prefs = get_prefs(context) if context else None
        sync_saved_blender_file(context, prefs, after_save=True)
        write_active_blender_info(context, prefs)
    except Exception:
        pass
    return None


def _refresh_active_blender_info():
    try:
        context = bpy.context
        prefs = get_prefs(context) if context else None
        sync_saved_blender_file(context, prefs)
        write_active_blender_info(context, prefs)
    except Exception:
        pass
    return None


def _active_blender_heartbeat():
    _refresh_active_blender_info()
    _auto_clear_cache_tick()
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
        key = normalize_path_key(root)
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
        try:
            candidates = list(root.rglob(MANIFEST_FILENAME))
        except OSError:
            continue
        for candidate in candidates:
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
        try:
            candidates = list(root.rglob(MANIFEST_FILENAME))
        except OSError:
            continue
        for candidate in candidates:
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
        try:
            candidates = list(root.rglob(MANIFEST_FILENAME))
        except OSError:
            continue
        for candidate in candidates:
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
        try:
            candidates = list(root.rglob(MANIFEST_FILENAME))
        except OSError:
            continue
        for candidate in candidates:
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
        try:
            candidates = list(root.rglob(MANIFEST_FILENAME))
        except OSError:
            continue
        for candidate in candidates:
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


def manifest_project_keys(manifest, manifest_path=None):
    if not isinstance(manifest, dict):
        return []
    values = [
        get_manifest_sp_project_file(manifest),
        get_manifest_link_sp_project_file(manifest),
    ]
    if manifest_path:
        try:
            project_dir = project_dir_from_manifest_path(manifest_path)
        except Exception:
            project_dir = None
        if project_dir:
            values.append(str(project_dir))
    keys = []
    seen = set()
    for value in values:
        if not value:
            continue
        key = normalize_path_key(value)
        if not key or key in seen:
            continue
        seen.add(key)
        keys.append(key)
    return keys


def get_object_project_key(obj):
    if not obj:
        return ""
    try:
        value = obj.get(OBJECT_PROJECT_KEY_PROP)
    except Exception:
        return ""
    if not value:
        return ""
    try:
        return str(value)
    except Exception:
        return ""


def tag_objects_with_project_key(context, objects, project_key, clear_existing=False):
    if not context or not project_key:
        return
    mesh_objects = [obj for obj in objects if obj and obj.type == "MESH"]
    keep_names = {obj.name for obj in mesh_objects}
    if clear_existing:
        for obj in context.scene.objects:
            if obj.type != "MESH" or obj.name in keep_names:
                continue
            if get_object_project_key(obj) != project_key:
                continue
            try:
                del obj[OBJECT_PROJECT_KEY_PROP]
            except Exception:
                continue
    for obj in mesh_objects:
        try:
            obj[OBJECT_PROJECT_KEY_PROP] = project_key
        except Exception:
            continue


def find_project_tag_targets(context, project_keys):
    if not project_keys:
        return []
    key_set = {key for key in project_keys if key}
    if not key_set:
        return []
    return [
        obj for obj in context.scene.objects
        if obj.type == "MESH" and get_object_project_key(obj) in key_set
    ]


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
    if project_dir and paths_match(active_info.get("project_dir"), project_dir):
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


def looks_like_bridge_root(path):
    try:
        path = Path(path)
    except (TypeError, ValueError):
        return False
    try:
        if path.name == documents_bridge_root().name:
            return True
        if (path / BRIDGE_ROOT_HINT_FILENAME).exists():
            return True
        if (path / PROJECT_META_DIRNAME).exists():
            return True
        if (path / TEMP_DIRNAME).exists():
            return True
    except OSError:
        return False
    return False


def looks_like_bridge_project_dir(path):
    try:
        if (path / PROJECT_META_DIRNAME).is_dir():
            return True
        if (path / MANIFEST_FILENAME).is_file():
            return True
    except OSError:
        return False
    return False


def clear_cache_dir_conservative(root):
    keep_names = {
        BRIDGE_ROOT_HINT_FILENAME,
        ACTIVE_SP_INFO_FILENAME,
        ACTIVE_BLENDER_INFO_FILENAME,
        LINKS_FILENAME,
        TEMP_DIRNAME,
    }
    try:
        children = list(root.iterdir())
    except OSError:
        return "error"
    for child in children:
        if child.name in keep_names:
            continue
        try:
            if child.is_dir():
                if not looks_like_bridge_project_dir(child):
                    continue
                shutil.rmtree(child)
        except OSError:
            return "error"
    ensure_dir(root)
    return "cleared"


def clear_cache_dir(path):
    if not path.exists():
        return "empty"
    if not looks_like_bridge_root(path):
        return clear_cache_dir_conservative(path)
    try:
        shutil.rmtree(path)
    except OSError:
        return "error"
    ensure_dir(path)
    _project_dir_cache.clear()
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
            keep.add(normalize_path_key(path_obj))
    try:
        for child in root.iterdir():
            try:
                child_key = normalize_path_key(child.resolve())
            except OSError:
                child_key = normalize_path_key(child)
            if child.is_file() and child.name in {
                BRIDGE_ROOT_HINT_FILENAME,
                ACTIVE_SP_INFO_FILENAME,
                ACTIVE_BLENDER_INFO_FILENAME,
                LINKS_FILENAME,
            }:
                continue
            if child.is_dir() and child.name == TEMP_DIRNAME:
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
    _project_dir_cache.clear()
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
    local_version = local_version_string()
    url = UPDATE_URL + "?current=" + urllib.parse.quote(local_version, safe="")
    # The edge blocks urllib's default User-Agent (403); identify like the
    # Substance Painter plugin does.
    request = urllib.request.Request(
        url, headers={"User-Agent": f"GoBBridge/{local_version}"}
    )
    try:
        with urllib.request.urlopen(request, timeout=4) as response:
            data = json.load(response)
    except (OSError, json.JSONDecodeError, urllib.error.URLError) as exc:
        return {"status": "error", "error": str(exc)}
    if not isinstance(data, dict):
        return {"status": "error", "error": "Invalid update data"}
    remote_version = str(data.get("version") or "").strip()
    download_url = ""
    files = data.get("files")
    if isinstance(files, list):
        for entry in files:
            if not isinstance(entry, dict):
                continue
            name = str(entry.get("name") or "")
            if name.lower().startswith("gob_blender"):
                download_url = str(entry.get("url") or "")
                break
    if not download_url:
        download_url = str(data.get("page") or "")
    if not remote_version and not download_url:
        return {"status": "error", "error": "Missing remote version"}
    if not remote_version or not is_version_newer(remote_version, local_version):
        return {
            "status": "none",
            "local_version": local_version,
            "remote_version": remote_version,
        }
    return {
        "status": "update",
        "info": {
            "version": remote_version,
            "download_url": download_url or None,
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
    if hasattr(mat, "use_nodes"):
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
        separate = nodes.new("ShaderNodeSeparateColor")
        separate.location = (-220, -300)
        links.new(orm_node.outputs["Color"], separate.inputs["Color"])
        if ao_output is None:
            ao_output = separate.outputs["Red"]
        if roughness_output is None:
            roughness_output = separate.outputs["Green"]
        if metallic_output is None:
            metallic_output = separate.outputs["Blue"]
    if metallic_roughness_node:
        separate = nodes.new("ShaderNodeSeparateColor")
        separate.location = (-220, -120)
        links.new(metallic_roughness_node.outputs["Color"], separate.inputs["Color"])
        if roughness_output is None:
            roughness_output = separate.outputs["Green"]
        if metallic_output is None:
            metallic_output = separate.outputs["Blue"]
    if metallic_smoothness_node:
        separate = nodes.new("ShaderNodeSeparateColor")
        separate.location = (-220, -180)
        links.new(metallic_smoothness_node.outputs["Color"], separate.inputs["Color"])
        if metallic_output is None:
            metallic_output = separate.outputs["Red"]
        if roughness_output is None:
            invert = nodes.new("ShaderNodeInvert")
            invert.inputs["Fac"].default_value = 1.0
            invert.location = (-120, -200)
            links.new(metallic_smoothness_node.outputs["Alpha"], invert.inputs["Color"])
            roughness_output = invert.outputs["Color"]
    if mask_node:
        separate = nodes.new("ShaderNodeSeparateColor")
        separate.location = (-220, -240)
        links.new(mask_node.outputs["Color"], separate.inputs["Color"])
        if metallic_output is None:
            metallic_output = separate.outputs["Red"]
        if ao_output is None:
            ao_output = separate.outputs["Green"]
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
            separate = nodes.new("ShaderNodeSeparateColor")
            separate.location = (-250, -520)
            invert = nodes.new("ShaderNodeInvert")
            invert.location = (-200, -640)
            combine = nodes.new("ShaderNodeCombineColor")
            combine.location = (-100, -520)
            links.new(normal_node.outputs["Color"], separate.inputs["Color"])
            links.new(separate.outputs["Red"], combine.inputs["Red"])
            links.new(separate.outputs["Green"], invert.inputs["Color"])
            links.new(invert.outputs["Color"], combine.inputs["Green"])
            links.new(separate.outputs["Blue"], combine.inputs["Blue"])
            links.new(combine.outputs["Color"], normal_map.inputs["Color"])
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
        if hasattr(mat, "blend_method"):
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


def find_texture_targets(context, grouped, project_keys=None):
    if not grouped:
        return []
    if project_keys:
        tagged = find_project_tag_targets(context, project_keys)
        if tagged:
            return tagged
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
            try:
                was_hidden_view = obj.hide_get()
            except Exception:
                was_hidden_view = False
            obj_states.append((
                obj,
                was_hidden_view,
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
    export_ok = True
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
    except Exception:
        export_ok = False
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
        for obj, was_hidden_view, was_hidden, was_render, was_select in obj_states:
            if not object_is_valid(obj):
                continue
            try:
                obj.hide_set(was_hidden_view)
            except Exception:
                pass
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
    return export_ok


def export_selected_fbx(filepath, prefs=None, strip_uvs=False):
    return export_fbx_objects(
        filepath,
        bpy.context.selected_objects,
        prefs=prefs,
        strip_uvs=strip_uvs,
    )


def mesh_has_uvs(mesh):
    if not mesh:
        return False
    try:
        return bool(mesh.uv_layers)
    except Exception:
        return False


def object_has_uvs(obj, depsgraph=None):
    if obj.type != "MESH":
        return False
    if mesh_has_uvs(getattr(obj, "data", None)):
        return True
    try:
        if depsgraph is None:
            depsgraph = bpy.context.evaluated_depsgraph_get()
        evaluated_obj = obj.evaluated_get(depsgraph)
        mesh = evaluated_obj.to_mesh()
    except Exception:
        return False
    try:
        return mesh_has_uvs(mesh)
    finally:
        try:
            evaluated_obj.to_mesh_clear()
        except Exception:
            pass


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


_tri_count_cache = {}
_obj_info_cache = {}
_export_plan_cache = {"key": None, "plan": None}


def evaluated_triangle_count(obj, depsgraph=None):
    key = None
    try:
        key = (obj.name, obj.data.as_pointer())
    except Exception:
        key = None
    if key is not None and key in _tri_count_cache:
        return _tri_count_cache[key]
    try:
        if depsgraph is None:
            depsgraph = bpy.context.evaluated_depsgraph_get()
        evaluated_obj = obj.evaluated_get(depsgraph)
        mesh = evaluated_obj.to_mesh()
        try:
            mesh.calc_loop_triangles()
            count = len(mesh.loop_triangles)
        finally:
            try:
                evaluated_obj.to_mesh_clear()
            except Exception:
                pass
    except Exception:
        count = mesh_triangle_count(obj)
    if key is not None:
        _tri_count_cache[key] = count
    return count


def autosplit_obj_info(obj, depsgraph=None):
    key = None
    try:
        key = (obj.name, obj.data.as_pointer())
    except Exception:
        key = None
    if key is not None and key in _obj_info_cache:
        return _obj_info_cache[key]
    count = evaluated_triangle_count(obj, depsgraph)
    bbox = None
    try:
        if depsgraph is None:
            depsgraph = bpy.context.evaluated_depsgraph_get()
        evaluated_obj = obj.evaluated_get(depsgraph)
        corners = [
            evaluated_obj.matrix_world @ Vector(corner)
            for corner in evaluated_obj.bound_box
        ]
        bbox = (
            min(corner.x for corner in corners),
            min(corner.y for corner in corners),
            min(corner.z for corner in corners),
            max(corner.x for corner in corners),
            max(corner.y for corner in corners),
            max(corner.z for corner in corners),
        )
    except Exception:
        bbox = None
    has_subsurf = False
    try:
        has_subsurf = any(
            mod and mod.type in {"SUBSURF", "MULTIRES"} and mod.show_viewport
            for mod in obj.modifiers
        )
    except Exception:
        has_subsurf = False
    info = {"count": count, "bbox": bbox, "has_subsurf": has_subsurf}
    if key is not None:
        _obj_info_cache[key] = info
    return info


def matched_name_suffix(name, suffixes):
    lname = name.lower()
    for suffix in suffixes:
        if lname.endswith(suffix):
            return suffix
    return ""


def bbox_intersects(left, right):
    return (
        left[0] <= right[3] and right[0] <= left[3]
        and left[1] <= right[4] and right[1] <= left[4]
        and left[2] <= right[5] and right[2] <= left[5]
    )


def spatial_clusters(meshes, infos):
    padded = {}
    for obj in meshes:
        bbox = infos[obj.name]["bbox"]
        if not bbox:
            continue
        pad_x = max(0.02 * (bbox[3] - bbox[0]), 0.0)
        pad_y = max(0.02 * (bbox[4] - bbox[1]), 0.0)
        pad_z = max(0.02 * (bbox[5] - bbox[2]), 0.0)
        padded[obj.name] = (
            bbox[0] - pad_x,
            bbox[1] - pad_y,
            bbox[2] - pad_z,
            bbox[3] + pad_x,
            bbox[4] + pad_y,
            bbox[5] + pad_z,
        )
    names = sorted(padded)
    parent = {name: name for name in names}

    def find(name):
        while parent[name] != name:
            parent[name] = parent[parent[name]]
            name = parent[name]
        return name

    for index, name_a in enumerate(names):
        for name_b in names[index + 1:]:
            if bbox_intersects(padded[name_a], padded[name_b]):
                root_a = find(name_a)
                root_b = find(name_b)
                if root_a != root_b:
                    parent[root_b] = root_a
    obj_by_name = {obj.name: obj for obj in meshes}
    clusters = {}
    for name in names:
        clusters.setdefault(find(name), []).append(obj_by_name[name])
    return [cluster for cluster in clusters.values() if len(cluster) >= 2]


def classify_low_high(context, objects, depsgraph=None, method="SMART"):
    meshes = [obj for obj in objects if obj and obj.type == "MESH"]
    if not meshes:
        return [], [], ["No meshes selected"], {}
    reasons = {}
    if len(meshes) == 1:
        reasons[meshes[0].name] = "only mesh"
        return [meshes[0]], [], [], reasons
    if depsgraph is None:
        try:
            depsgraph = context.evaluated_depsgraph_get()
        except Exception:
            depsgraph = None
    prefs = get_prefs(context)
    low_suffixes = parse_suffixes(getattr(prefs, "low_poly_suffixes", "")) or ["_low"]
    high_suffixes = parse_suffixes(getattr(prefs, "high_poly_suffixes", "")) or ["_high"]

    def strip_base(name):
        base = re.sub(r"\.\d+$", "", name.lower())
        for suffix in low_suffixes + high_suffixes:
            if base.endswith(suffix):
                base = base[: -len(suffix)]
                break
        base = re.sub(r"_lod\d+$", "", base)
        base = re.sub(r"_\d+$", "", base)
        return base

    infos = {obj.name: autosplit_obj_info(obj, depsgraph) for obj in meshes}
    counts = {name: info["count"] for name, info in infos.items()}
    warnings = []

    def name_groups(items):
        groups = {}
        for obj in items:
            groups.setdefault(strip_base(obj.name), []).append(obj)
        return groups

    def gap_split(items):
        items = sorted(items, key=lambda obj: (counts[obj.name], obj.name))
        if len(items) < 2:
            return list(items), [], False
        sorted_counts = [counts[obj.name] for obj in items]
        if sorted_counts[0] == sorted_counts[-1]:
            return list(items), [], False
        best_index = 0
        best_gap = -1
        for idx in range(len(items) - 1):
            gap = sorted_counts[idx + 1] - sorted_counts[idx]
            if gap > best_gap:
                best_gap = gap
                best_index = idx
        count_below = sorted_counts[best_index]
        count_above = sorted_counts[best_index + 1]
        ratio = count_above / count_below if count_below > 0 else float("inf")
        if ratio < 2.0:
            return list(items), [], False
        return list(items[: best_index + 1]), list(items[best_index + 1:]), True

    if method == "NAMES":
        low = []
        high = []
        unpaired = []
        for group in name_groups(meshes).values():
            lows = [obj for obj in group if is_name_with_suffix(obj.name, low_suffixes)]
            highs = [obj for obj in group if is_name_with_suffix(obj.name, high_suffixes)]
            if len(group) == 2 and len(lows) == 1 and len(highs) == 1:
                low_obj = lows[0]
                high_obj = highs[0]
                low.append(low_obj)
                high.append(high_obj)
                reasons[low_obj.name] = f"name says {matched_name_suffix(low_obj.name, low_suffixes)}"
                reasons[high_obj.name] = f"name says {matched_name_suffix(high_obj.name, high_suffixes)}"
                if counts[low_obj.name] > counts[high_obj.name]:
                    warnings.append(f"{low_obj.name} has more triangles than {high_obj.name}")
            else:
                unpaired.extend(group)
        if unpaired:
            warnings.append("Unpaired meshes export as low")
            for obj in unpaired:
                low.append(obj)
                reasons[obj.name] = "unpaired name"
        return low, high, warnings, reasons

    if method == "TRIANGLES":
        if len({counts[obj.name] for obj in meshes}) == 1:
            for obj in meshes:
                reasons[obj.name] = "same triangle count"
            return list(meshes), [], warnings, reasons
        gap_low, gap_high, gap_found = gap_split(meshes)
        if not gap_found:
            warnings.append("No clear low/high separation, everything exports as low")
            for obj in meshes:
                reasons[obj.name] = "no clear separation"
            return list(meshes), [], warnings, reasons
        for obj in gap_low:
            reasons[obj.name] = "gap split: low"
        for obj in gap_high:
            reasons[obj.name] = "gap split: high"
        return gap_low, gap_high, warnings, reasons

    locked = {}
    for obj in meshes:
        try:
            tagged = bool(obj.get("gob_high_poly"))
        except Exception:
            tagged = False
        if tagged:
            locked[obj.name] = "high"
            reasons[obj.name] = "tagged as high poly"
    for obj in meshes:
        if obj.name in locked:
            continue
        low_suffix = matched_name_suffix(obj.name, low_suffixes)
        high_suffix = matched_name_suffix(obj.name, high_suffixes)
        if low_suffix and not high_suffix:
            locked[obj.name] = "low"
            reasons[obj.name] = f"name says {low_suffix}"
        elif high_suffix:
            locked[obj.name] = "high"
            reasons[obj.name] = f"name says {high_suffix}"
    for group in name_groups(meshes).values():
        for low_obj in [obj for obj in group if locked.get(obj.name) == "low"]:
            for high_obj in [obj for obj in group if locked.get(obj.name) == "high"]:
                if counts[low_obj.name] >= 2 * max(1, counts[high_obj.name]):
                    warnings.append(f"{low_obj.name} has more triangles than {high_obj.name}")
    remaining = [obj for obj in meshes if obj.name not in locked]
    unpaired = []
    for group in name_groups(remaining).values():
        if len(group) == 2:
            first, second = sorted(group, key=lambda obj: (counts[obj.name], obj.name))
            if counts[first.name] == counts[second.name]:
                warnings.append(
                    f"{first.name} and {second.name} have equal triangle counts, both export as low"
                )
                locked[first.name] = "low"
                locked[second.name] = "low"
                reasons[first.name] = f"paired by name with {second.name}, equal counts"
                reasons[second.name] = f"paired by name with {first.name}, equal counts"
            else:
                locked[first.name] = "low"
                locked[second.name] = "high"
                reasons[first.name] = f"paired by name with {second.name}"
                reasons[second.name] = f"paired by name with {first.name}"
        else:
            unpaired.extend(group)
    clustered_names = set()
    for cluster in spatial_clusters(unpaired, infos):
        ordered = sorted(cluster, key=lambda obj: (counts[obj.name], obj.name))
        lightest = ordered[0]
        locked[lightest.name] = "low"
        clustered_names.add(lightest.name)
        reasons[lightest.name] = f"shares space with {ordered[1].name}, lighter"
        for obj in ordered[1:]:
            locked[obj.name] = "high"
            clustered_names.add(obj.name)
            reasons[obj.name] = f"shares space with {lightest.name}, heavier"
    singles = [obj for obj in unpaired if obj.name not in clustered_names]
    leftover = []
    for obj in singles:
        if infos[obj.name]["has_subsurf"]:
            locked[obj.name] = "high"
            reasons[obj.name] = "has subdivision modifier"
        else:
            leftover.append(obj)
    if leftover:
        gap_low, gap_high, gap_found = gap_split(leftover)
        if gap_found:
            for obj in gap_low:
                locked[obj.name] = "low"
                reasons[obj.name] = "gap split: low"
            for obj in gap_high:
                locked[obj.name] = "high"
                reasons[obj.name] = "gap split: high"
        else:
            equal_counts = len({counts[obj.name] for obj in leftover}) == 1
            if not equal_counts:
                warnings.append("No clear separation, exports as low")
            for obj in leftover:
                locked[obj.name] = "low"
                reasons[obj.name] = "same triangle count" if equal_counts else "no clear separation"
    low = []
    high = []
    for obj in meshes:
        if locked.get(obj.name) == "high":
            high.append(obj)
        else:
            low.append(obj)
        reasons.setdefault(obj.name, "exports as low")
    return low, high, warnings, reasons


def build_export_plan(context, prefs, deep=False):
    scene = context.scene
    mode = get_identify_mode(scene)
    warnings = []
    reasons = {}
    depsgraph = None
    try:
        depsgraph = context.evaluated_depsgraph_get()
    except Exception:
        depsgraph = None
    if mode == "AUTO_SPLIT":
        selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
        pool = context.selected_objects if selected_only else scene.objects
        method = getattr(scene, "gob_sp_autosplit_method", "SMART")
        low, high, warnings, reasons = classify_low_high(
            context,
            pool,
            depsgraph,
            method=method,
        )
    else:
        low = collect_low_poly_objects(context, prefs)
        high = []
        if prefs and getattr(prefs, "export_high_poly", False):
            high = collect_high_poly_candidates(context, prefs)
            if high and low:
                high_names = {obj.name for obj in high}
                low = [obj for obj in low if obj.name not in high_names]
    cage = []
    if prefs and getattr(prefs, "export_cage_poly", False):
        cage = collect_cage_objects(context, prefs)
        if cage:
            cage_names = {obj.name for obj in cage}
            low = [obj for obj in low if obj.name not in cage_names]
            high = [obj for obj in high if obj.name not in cage_names]
        else:
            warnings.append("Cage export enabled but no cage meshes identified")
    if prefs and getattr(prefs, "export_low_poly", False) and not low and not warnings:
        warnings.append("No low poly meshes found")
    if (
        prefs and getattr(prefs, "export_high_poly", False)
        and not high and low and mode != "AUTO_SPLIT"
    ):
        warnings.append("No high poly meshes identified, sending low only")
    auto_unwrap = bool(prefs and getattr(prefs, "sp_auto_unwrap", False))
    if not auto_unwrap and prefs and getattr(prefs, "export_low_poly", False) and low:
        if deep:
            missing_uvs = [
                obj for obj in low if not object_has_uvs(obj, depsgraph=depsgraph)
            ]
        else:
            missing_uvs = [
                obj for obj in low if not mesh_has_uvs(getattr(obj, "data", None))
            ]
        if missing_uvs:
            warnings.append(
                f"Missing UVs on {len(missing_uvs)} low mesh(es), unwrap before export"
            )
    effective_high = bool(prefs and getattr(prefs, "export_high_poly", False)) and bool(high)
    return {
        "low": low,
        "high": high,
        "cage": cage,
        "warnings": warnings,
        "effective_high": effective_high,
        "reasons": reasons,
    }


def _export_plan_cache_key(context, prefs):
    scene = context.scene
    selected_only = bool(getattr(prefs, "export_selected_only", False))
    selected = None
    if selected_only:
        selected = tuple(sorted(
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        ))
    low_collection = getattr(scene, "gob_sp_low_poly_collection", None)
    high_collection = getattr(scene, "gob_sp_high_poly_collection", None)
    cage_collection = getattr(scene, "gob_sp_cage_collection", None)
    return (
        selected,
        len(scene.objects),
        get_identify_mode(scene),
        getattr(scene, "gob_sp_autosplit_method", "SMART"),
        selected_only,
        bool(getattr(prefs, "export_low_poly", False)),
        bool(getattr(prefs, "export_high_poly", False)),
        getattr(prefs, "low_poly_suffixes", ""),
        getattr(prefs, "high_poly_suffixes", ""),
        low_collection.name if low_collection else "",
        high_collection.name if high_collection else "",
        bool(getattr(prefs, "export_cage_poly", False)),
        getattr(prefs, "cage_poly_suffixes", ""),
        cage_collection.name if cage_collection else "",
    )


def get_cached_export_plan(context, prefs):
    key = _export_plan_cache_key(context, prefs)
    cached = _export_plan_cache
    if cached["plan"] is not None and cached["key"] == key:
        return cached["plan"]
    plan = build_export_plan(context, prefs)
    cached["key"] = key
    cached["plan"] = plan
    return plan


def _invalidate_export_caches():
    _tri_count_cache.clear()
    _obj_info_cache.clear()
    _export_plan_cache["key"] = None
    _export_plan_cache["plan"] = None


def _on_export_settings_update(self, _context):
    _invalidate_export_caches()
    _set_export_warning("")


def collect_high_poly_candidates(context, prefs):
    scene = context.scene
    objects = []
    selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
    selected_names = None
    if selected_only:
        selected_names = {
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        }
    if get_identify_mode(scene) == "COLLECTIONS":
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
    if selected_only and selected_names is not None:
        objects = [obj for obj in objects if obj.name in selected_names]
    unique = []
    seen = set()
    for obj in objects:
        if obj.name in seen:
            continue
        seen.add(obj.name)
        unique.append(obj)
    return unique


def collect_cage_objects(context, prefs):
    scene = context.scene
    objects = []
    selected_only = bool(prefs and getattr(prefs, "export_selected_only", False))
    selected_names = None
    if selected_only:
        selected_names = {
            obj.name for obj in context.selected_objects if obj.type == "MESH"
        }
    if get_identify_mode(scene) == "COLLECTIONS":
        cage_collection = getattr(scene, "gob_sp_cage_collection", None)
        if not collection_in_scene(scene, cage_collection):
            cage_collection = None
        if cage_collection:
            objects = collect_collection_meshes(
                cage_collection,
                selected_only=selected_only,
                selected_names=selected_names,
            )
            if objects:
                return objects
    suffixes = parse_suffixes(getattr(prefs, "cage_poly_suffixes", ""))
    if suffixes:
        for obj in context.scene.objects:
            if obj.type != "MESH":
                continue
            if is_name_with_suffix(obj.name, suffixes):
                objects.append(obj)
    for obj in context.scene.objects:
        if obj.type == "MESH" and obj.get("gob_cage_poly"):
            objects.append(obj)
    if selected_only and selected_names is not None:
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
    try:
        bpy.ops.import_scene.fbx(filepath=str(filepath))
    except Exception:
        return []
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


TEMPLATE_SCAN_CACHE_TTL = 5.0
_template_scan_cache = {"timestamp": 0.0, "entries": None}

COLOR_MANAGEMENT_MANIFEST_KEYS = {
    "PAINTER_DEFAULT": "painter_default",
    "OCIO_SUBSTANCE": "ocio:substance",
    "OCIO_ACES_1_0_3": "ocio:aces_1.0.3",
    "OCIO_ACES_1_2": "ocio:aces_1.2",
    "OCIO_ACES_2_0": "ocio:aces_2.0",
    "OCIO_CUSTOM": "ocio_custom",
}

PROJECT_WORKFLOW_MANIFEST_KEYS = {
    "DEFAULT": "Default",
    "UV_TILE": "UVTile",
    "TEXTURE_SET_PER_UV_TILE": "TextureSetPerUVTile",
}

NORMAL_MAP_FORMAT_MANIFEST_KEYS = {
    "OPENGL": "OpenGL",
    "DIRECTX": "DirectX",
}


def sp_install_resources_dir(sp_exe):
    if not sp_exe:
        return None
    try:
        path = Path(sp_exe).expanduser()
    except Exception:
        return None
    if path.suffix.lower() == ".app":
        return path / "Contents" / "Resources"
    for parent in path.parents:
        if parent.suffix.lower() == ".app":
            return parent / "Contents" / "Resources"
    return path.parent / "resources"


def sp_user_templates_dir():
    try:
        docs = windows_documents_dir()
    except Exception:
        docs = None
    if not docs:
        try:
            docs = Path.home() / "Documents"
        except Exception:
            return None
    return (
        Path(docs)
        / "Adobe"
        / "Adobe Substance 3D Painter"
        / "assets"
        / "templates"
    )


def _template_entries_from_dir(directory, prefix):
    entries = []
    try:
        if not directory or not directory.is_dir():
            return entries
        for path in sorted(directory.glob("*.spt"), key=lambda item: item.name.lower()):
            if not path.is_file():
                continue
            label = path.stem if prefix == "builtin" else f"{path.stem} (user)"
            entries.append({
                "key": f"{prefix}:{path.stem}",
                "label": label,
                "path": str(path),
            })
    except Exception:
        return []
    return entries


def scan_sp_templates(prefs=None):
    entries = []
    try:
        resources = sp_install_resources_dir(find_sp_exe(prefs))
        if resources:
            entries.extend(_template_entries_from_dir(
                resources / "starter_assets" / "templates",
                "builtin",
            ))
    except Exception:
        pass
    try:
        entries.extend(_template_entries_from_dir(sp_user_templates_dir(), "user"))
    except Exception:
        pass
    return entries


def get_sp_template_entries(prefs=None, force=False):
    now = time.time()
    if (
        not force
        and _template_scan_cache["entries"] is not None
        and now - _template_scan_cache["timestamp"] < TEMPLATE_SCAN_CACHE_TTL
    ):
        return _template_scan_cache["entries"]
    entries = scan_sp_templates(prefs)
    _template_scan_cache["timestamp"] = now
    _template_scan_cache["entries"] = entries
    return entries


def _refresh_template_cache():
    _template_scan_cache["timestamp"] = 0.0
    _template_scan_cache["entries"] = None


def _sp_template_picker_items(self, _context):
    items = [(
        "NONE",
        "Painter default",
        "No template, Painter's standard new project",
        0,
    )]
    try:
        for index, entry in enumerate(get_sp_template_entries(), start=1):
            items.append((entry["key"], entry["label"], "", index))
    except Exception:
        pass
    return items


def _sp_template_picker_get(self):
    current = getattr(self, "sp_project_template", "NONE") or "NONE"
    for _key, _label, _desc, number in _sp_template_picker_items(self, None):
        if _key == current:
            return number
    return 0


def _sp_template_picker_set(self, value):
    for key, _label, _desc, number in _sp_template_picker_items(self, None):
        if number == value:
            self.sp_project_template = key
            return
    self.sp_project_template = "NONE"


def resolve_sp_template_path(template_key, prefs=None):
    if not template_key or template_key == "NONE":
        return ""
    try:
        prefix, stem = template_key.split(":", 1)
    except ValueError:
        return ""
    if not stem or stem in (".", "..") or "/" in stem or "\\" in stem:
        return ""
    base = None
    if prefix == "builtin":
        try:
            resources = sp_install_resources_dir(find_sp_exe(prefs))
        except Exception:
            resources = None
        if resources:
            base = resources / "starter_assets" / "templates"
    elif prefix == "user":
        base = sp_user_templates_dir()
    if not base:
        return ""
    candidate = base / f"{stem}.spt"
    try:
        if candidate.is_file():
            return str(candidate)
    except OSError:
        return ""
    return ""


def build_sp_project_settings(prefs):
    warnings = []
    if prefs is None:
        return None, warnings
    raw_template = getattr(prefs, "sp_project_template", "NONE")
    template_key = raw_template or "NONE"
    color_mode = getattr(prefs, "sp_color_management", "PAINTER_DEFAULT") or "PAINTER_DEFAULT"
    auto_unwrap = bool(getattr(prefs, "sp_auto_unwrap", False))
    unwrap_mode = getattr(prefs, "sp_unwrap_mode", "DEFAULT") or "DEFAULT"
    resolution = getattr(prefs, "sp_default_resolution", "PAINTER_DEFAULT") or "PAINTER_DEFAULT"
    workflow = getattr(prefs, "sp_project_workflow", "PAINTER_DEFAULT") or "PAINTER_DEFAULT"
    normal_format = getattr(prefs, "sp_normal_map_format", "PAINTER_DEFAULT") or "PAINTER_DEFAULT"

    template_path = ""
    if template_key != "NONE":
        known_keys = set()
        try:
            known_keys = {entry["key"] for entry in get_sp_template_entries(prefs)}
        except Exception:
            known_keys = set()
        if template_key not in known_keys:
            warnings.append("Stored template is no longer available, using Painter default")
            template_key = "NONE"
        else:
            template_path = resolve_sp_template_path(template_key, prefs)
            if not template_path:
                warnings.append(
                    "Template not found, using Painter default: "
                    + template_key.split(":", 1)[-1]
                )

    has_custom_template = bool(template_key != "NONE" and template_path)
    if (
        not has_custom_template
        and color_mode == "PAINTER_DEFAULT"
        and not auto_unwrap
        and resolution == "PAINTER_DEFAULT"
        and workflow == "PAINTER_DEFAULT"
        and normal_format == "PAINTER_DEFAULT"
    ):
        return None, warnings

    color_key = COLOR_MANAGEMENT_MANIFEST_KEYS.get(color_mode, "painter_default")
    color_config_path = ""
    if color_mode == "OCIO_CUSTOM":
        raw_path = (getattr(prefs, "sp_color_config_path", "") or "").strip()
        candidate = normalize_path(bpy.path.abspath(raw_path)) if raw_path else ""
        if candidate and Path(candidate).is_file():
            color_config_path = candidate
        else:
            warnings.append(
                "Custom .ocio config not found, using Painter default color management"
            )
            color_key = "painter_default"

    settings = {
        "template": template_path,
        "color_management": color_key,
        "color_config_path": color_config_path,
        "auto_unwrap": auto_unwrap,
        "unwrap_mode": "hard_surface" if unwrap_mode == "HARD_SURFACE" else "default",
    }
    if resolution != "PAINTER_DEFAULT":
        settings["default_texture_resolution"] = int(resolution)
    if workflow != "PAINTER_DEFAULT":
        settings["project_workflow"] = PROJECT_WORKFLOW_MANIFEST_KEYS.get(
            workflow, "Default"
        )
    if normal_format != "PAINTER_DEFAULT":
        settings["normal_map_format"] = NORMAL_MAP_FORMAT_MANIFEST_KEYS.get(
            normal_format, "OpenGL"
        )
    return settings, warnings


def open_sp_project_file(project_file, sp_exe=None):
    if not project_file:
        return False
    try:
        if sys.platform == "darwin":
            if sp_exe and sp_exe.lower().endswith(".app"):
                subprocess.Popen(["open", "-a", sp_exe, project_file])
            elif sp_exe and Path(sp_exe).is_file():
                subprocess.Popen([sp_exe, project_file])
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
    name_tokens = ("adobe substance 3d painter", "substance 3d painter")

    def _contains_sp_name(text):
        haystack = (text or "").lower()
        return any(token in haystack for token in name_tokens)

    def _run_capture(cmd, timeout=2.0):
        kwargs = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "text": True,
            "encoding": "utf-8",
            "errors": "ignore",
            "timeout": timeout,
            "check": False,
        }
        if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        try:
            result = subprocess.run(cmd, **kwargs)
        except (OSError, ValueError, subprocess.SubprocessError):
            return ""
        return (result.stdout or "") + "\n" + (result.stderr or "")

    if os.name == "nt":
        for cmd in (["tasklist", "/FO", "CSV", "/NH"], ["tasklist", "/FO", "CSV"]):
            if _contains_sp_name(_run_capture(cmd)):
                return True
        ps_cmd = (
            "$ErrorActionPreference='SilentlyContinue'; "
            "Get-Process | Where-Object { "
            "$_.ProcessName -like '*Substance*Painter*' -or "
            "$_.ProcessName -like '*Adobe*Substance*Painter*' "
            "} | Select-Object -ExpandProperty ProcessName"
        )
        return _contains_sp_name(
            _run_capture(
                [
                    "powershell",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_cmd,
                ]
            )
        )

    # macOS/Linux fallback when Windows process tools are unavailable.
    if _contains_sp_name(_run_capture(["pgrep", "-fl", "Substance 3D Painter"])):
        return True
    return _contains_sp_name(_run_capture(["ps", "-A", "-o", "comm="]))


_update_check_in_progress = False
_update_check_result = None
_update_check_show_no_update = False
_update_check_show_popup = False
_last_update_info = None
_update_status_kind = "idle"
_update_status_text = "Update: not checked yet"
_update_status_time = 0.0
_last_export_warning = ""
_cache_size_check_time = 0.0
_cache_size_global = 0
_cache_size_local = 0
_cache_size_project_root = None
_last_auto_clear_time = 0.0
_sp_running_cache = {"timestamp": 0.0, "running": None}
_sp_running_probe_thread = None
_import_probe_cache = {"timestamp": 0.0, "key": "", "available": False}
SP_RUNNING_CACHE_TTL = 10.0
SP_RUNNING_SEND_MAX_AGE = 15.0
IMPORT_PROBE_CACHE_TTL = 3.0


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
    try:
        _update_check_result = check_for_updates()
    except Exception as exc:
        _update_check_result = {"status": "error", "error": str(exc)}


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


def elide_path(text, limit=45):
    text = str(text or "")
    if len(text) <= limit:
        return text
    return "…" + text[-(limit - 1):]


def _sp_running_probe():
    global _sp_running_probe_thread
    try:
        running = bool(is_sp_running())
    except Exception:
        running = False
    _sp_running_cache["running"] = running
    _sp_running_cache["timestamp"] = time.time()
    _sp_running_probe_thread = None


def get_cached_sp_running():
    global _sp_running_probe_thread
    now = time.time()
    if now - _sp_running_cache["timestamp"] >= SP_RUNNING_CACHE_TTL:
        if _sp_running_probe_thread is None or not _sp_running_probe_thread.is_alive():
            _sp_running_probe_thread = threading.Thread(
                target=_sp_running_probe,
                daemon=True,
            )
            _sp_running_probe_thread.start()
    return _sp_running_cache["running"]


def sp_running_for_send():
    now = time.time()
    cached = _sp_running_cache["running"]
    if cached is not None and now - _sp_running_cache["timestamp"] < SP_RUNNING_SEND_MAX_AGE:
        return cached
    running = bool(is_sp_running())
    _sp_running_cache["running"] = running
    _sp_running_cache["timestamp"] = time.time()
    return running


def get_ui_link_status(context, prefs):
    now = time.time()
    cache = _ui_link_cache
    blender_file = get_blender_file_path_or_temp(prefs)
    if now - cache["timestamp"] < UI_LINK_CACHE_TTL and cache["blender_file"] == blender_file:
        return cache
    project_dir = None
    active_info = None
    try:
        project_dir = get_project_dir_fast(context, prefs)
    except Exception:
        project_dir = None
    try:
        if project_dir:
            active_info = read_active_sp_info(
                project_meta_dir(project_dir) / ACTIVE_SP_INFO_FILENAME
            )
    except Exception:
        active_info = None
    linked_sp_project = ""
    try:
        linked_sp_project = resolve_linked_sp_project_file_fast(
            project_dir,
            active_info=active_info,
            blender_file=blender_file,
            prefs=prefs,
        )
    except Exception:
        linked_sp_project = ""
    cache["timestamp"] = now
    cache["blender_file"] = blender_file
    cache["project_dir"] = str(project_dir) if project_dir else ""
    cache["active_info"] = active_info
    cache["linked_sp_project"] = str(linked_sp_project) if linked_sp_project else ""
    cache["sp_running"] = get_cached_sp_running()
    return cache


def import_available(context, prefs):
    now = time.time()
    blender_file = get_blender_file_path_or_temp(prefs)
    key = normalize_path_key(blender_file)
    if (
        key == _import_probe_cache["key"]
        and now - _import_probe_cache["timestamp"] < IMPORT_PROBE_CACHE_TTL
    ):
        return _import_probe_cache["available"]
    available = False
    try:
        project_dir = get_project_dir_fast(context, prefs)
        if project_dir:
            manifest_path = find_project_manifest_path(project_dir)
            if manifest_path and manifest_path.exists():
                available = True
        if not available and find_active_sp_project_info(prefs):
            available = True
    except Exception:
        available = False
    _import_probe_cache["key"] = key
    _import_probe_cache["timestamp"] = now
    _import_probe_cache["available"] = available
    return available


def _on_bridge_dir_update(self, _context):
    _project_dir_cache.clear()


def _on_auto_clear_cache_update(self, _context):
    if not self.auto_clear_cache:
        return
    limit = getattr(self, "cache_limit_gb", DEFAULT_CACHE_LIMIT_GB)
    message = (
        "Warning: auto clear removes cached projects (keeps the current project) "
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
    bpy.app.timers.register(_update_poll, first_interval=0.5, persistent=True)


def get_cached_cache_sizes(context, prefs, max_age=30.0):
    global _cache_size_check_time
    global _cache_size_global
    global _cache_size_local
    global _cache_size_project_root
    now = time.time()
    max_age = max(30.0, max_age)
    project_dir = str(get_project_dir(context, prefs))
    if (
        _cache_size_project_root != project_dir
        or now - _cache_size_check_time > max_age
    ):
        _cache_size_project_root = project_dir
        _cache_size_global = bridge_cache_size_bytes(prefs)
        _cache_size_local = project_cache_size_bytes(context, prefs)
        _cache_size_check_time = now
    return _cache_size_global, _cache_size_local


def _auto_clear_cache_tick():
    global _last_auto_clear_time
    try:
        context = bpy.context
        if context is None:
            return
        prefs = get_prefs(context)
        if not prefs or not getattr(prefs, "auto_clear_cache", False):
            return
        now = time.time()
        if now - _last_auto_clear_time < 30.0:
            return
        _last_auto_clear_time = now
        limit_bytes = cache_limit_bytes(prefs)
        if not limit_bytes:
            return
        if bridge_cache_size_bytes(prefs) <= limit_bytes:
            return
        keep_paths = [get_project_dir(context, prefs)]
        result = clear_cache_dir_except(get_bridge_root(prefs), keep_paths=keep_paths)
        if result == "cleared":
            refresh_cache_sizes(context, prefs)
    except Exception:
        pass


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


@persistent
def _init_scene_ui_prefs(_context=None):
    _refresh_active_blender_info()
    return None


class GOBSPPreferences(AddonPreferences):
    bl_idname = __name__

    bridge_dir: StringProperty(
        name="Bridge Folder",
        description=(
            "Shared cache folder for the Blender/Painter bridge "
            "(the GOB_SP_BRIDGE_DIR environment variable overrides this)"
        ),
        subtype="DIR_PATH",
        default=default_bridge_dir(),
        update=_on_bridge_dir_update,
    )
    auto_launch_sp: BoolProperty(
        name="Auto launch Substance Painter",
        default=True,
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
        update=_on_export_settings_update,
    )
    export_low_poly: BoolProperty(
        name="Export Low Poly",
        default=True,
        update=_on_export_settings_update,
    )
    export_cage_poly: BoolProperty(
        name="Export Cage",
        description="Export a custom cage mesh for baking in Painter",
        default=False,
        update=_on_export_settings_update,
    )
    export_selected_only: BoolProperty(
        name="Only Selected Meshes",
        description="Limit low/high exports to the current selection",
        default=False,
        update=_on_export_settings_update,
    )
    low_poly_suffixes: StringProperty(
        name="Low Poly Suffixes",
        description="Comma separated suffixes for low poly objects (must be at end)",
        default="_low",
        update=_on_export_settings_update,
    )
    high_poly_suffixes: StringProperty(
        name="High Poly Suffixes",
        description="Comma separated suffixes for high poly objects (must be at end)",
        default="_high",
        update=_on_export_settings_update,
    )
    cage_poly_suffixes: StringProperty(
        name="Cage Poly Suffixes",
        description="Comma separated suffixes for cage objects (must be at end)",
        default="_cage",
        update=_on_export_settings_update,
    )
    fbx_export_scale: FloatProperty(
        name="FBX Export Scale",
        description="If triangles come out too small in Painter, raise Export Scale",
        default=1.0,
        min=0.001,
        max=1000.0,
    )
    fbx_apply_unit_scale: BoolProperty(
        name="Apply Unit Scale",
        default=True,
    )
    auto_clear_cache: BoolProperty(
        name="Auto clear Global Cache",
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
    sp_project_template: StringProperty(
        name="Template Key",
        description="Stored project template key (use the Template dropdown to change it)",
        default="NONE",
    )
    sp_project_template_picker: EnumProperty(
        name="Template",
        description="Painter project template used when a new Painter project is created",
        items=_sp_template_picker_items,
        get=_sp_template_picker_get,
        set=_sp_template_picker_set,
    )
    sp_color_management: EnumProperty(
        name="Color Management",
        description="Color management used when a new Painter project is created",
        items=(
            ("PAINTER_DEFAULT", "Painter default (Legacy)", "Let Painter pick its standard Legacy color management"),
            ("OCIO_SUBSTANCE", "OpenColorIO: Substance (sRGB/Rec709)", "Use Painter's bundled Substance OpenColorIO config"),
            ("OCIO_ACES_1_0_3", "OpenColorIO: ACES 1.0.3", "Use Painter's bundled ACES 1.0.3 OpenColorIO config"),
            ("OCIO_ACES_1_2", "OpenColorIO: ACES 1.2", "Use Painter's bundled ACES 1.2 OpenColorIO config"),
            ("OCIO_ACES_2_0", "OpenColorIO: ACES 2.0", "Use Painter's bundled ACES 2.0 OpenColorIO config"),
            ("OCIO_CUSTOM", "Custom .ocio file", "Use your own .ocio config file"),
        ),
        default="PAINTER_DEFAULT",
    )
    sp_color_config_path: StringProperty(
        name="Custom .ocio Config",
        description="Path to your custom .ocio config file",
        subtype="FILE_PATH",
        default="",
    )
    sp_auto_unwrap: BoolProperty(
        name="Unwrap UVs in Painter",
        description="Let Painter generate the UVs when a new Painter project is created",
        default=False,
    )
    sp_unwrap_mode: EnumProperty(
        name="Unwrap Mode",
        description="Hard surface unwrap needs Painter 12.1 or newer",
        items=(
            ("DEFAULT", "Default", "Painter's standard auto unwrap"),
            ("HARD_SURFACE", "Hard surface", "Hard surface unwrap (needs Painter 12.1+)"),
        ),
        default="DEFAULT",
    )
    sp_default_resolution: EnumProperty(
        name="Default Texture Resolution",
        description="Texture resolution used when a new Painter project is created",
        items=(
            ("PAINTER_DEFAULT", "Painter default", "Let Painter pick the texture resolution"),
            ("512", "512", "512 pixels"),
            ("1024", "1024", "1024 pixels"),
            ("2048", "2048", "2048 pixels"),
            ("4096", "4096", "4096 pixels"),
            ("8192", "8192", "8192 pixels"),
        ),
        default="PAINTER_DEFAULT",
    )
    sp_project_workflow: EnumProperty(
        name="Project Workflow",
        description="Project workflow used when a new Painter project is created",
        items=(
            ("PAINTER_DEFAULT", "Painter default", "Let Painter pick the project workflow"),
            ("DEFAULT", "Default", "One texture set per mesh material"),
            ("UV_TILE", "UV Tiles (UDIM)", "UV tiles across texture sets"),
            ("TEXTURE_SET_PER_UV_TILE", "Texture Set per UV Tile (legacy)", "Legacy UV tile workflow"),
        ),
        default="PAINTER_DEFAULT",
    )
    sp_normal_map_format: EnumProperty(
        name="Normal Map Format",
        description="Normal map format used when a new Painter project is created",
        items=(
            ("PAINTER_DEFAULT", "Painter default", "Let Painter pick the normal map format"),
            ("OPENGL", "OpenGL", "OpenGL style normal maps"),
            ("DIRECTX", "DirectX", "DirectX style normal maps"),
        ),
        default="PAINTER_DEFAULT",
    )

    def draw(self, _context):
        layout = self.layout
        layout.prop(self, "bridge_dir")
        layout.prop(self, "auto_launch_sp")
        layout.separator()
        box = layout.box()
        box.label(text="Cache")
        box.prop(self, "auto_clear_cache")
        row = box.row()
        row.enabled = self.auto_clear_cache
        row.prop(self, "cache_limit_gb")
        layout.separator()
        layout.label(text=f"Version {local_version_string()}")
        layout.operator(GOB_OT_CheckUpdates.bl_idname, text="Check for Updates")


class GOB_OT_SendToSP(Operator):
    bl_idname = "gob_sp.send_to_substance_painter"
    bl_label = "Send to Substance Painter"
    bl_description = "Export the low/high poly meshes as FBX and hand them to Substance Painter"

    @classmethod
    def poll(cls, context):
        prefs = get_prefs(context)
        if not prefs:
            return False
        if getattr(prefs, "export_selected_only", False):
            return any(obj.type == "MESH" for obj in context.selected_objects)
        scene = context.scene
        return bool(scene) and any(obj.type == "MESH" for obj in scene.objects)

    def execute(self, context):
        prefs = get_prefs(context)
        write_active_blender_info(context, prefs)
        if _bridge_conflict_info:
            self.report({"WARNING"}, "Another Blender instance is using this bridge")
        blender_file = get_blender_file_path_or_temp(prefs)
        force_new_project = bool(prefs and prefs.force_new_sp_project_on_send)
        auto_unwrap = bool(prefs and getattr(prefs, "sp_auto_unwrap", False))
        _invalidate_export_caches()
        plan = build_export_plan(context, prefs, deep=True)
        low_objects = plan["low"]
        high_candidates = plan["high"]
        cage_objects = plan.get("cage") or []
        if plan["warnings"]:
            _set_export_warning(plan["warnings"][0])
            for message in plan["warnings"]:
                self.report({"WARNING"}, message)
        else:
            _set_export_warning("")
        if not low_objects and (not prefs or prefs.export_low_poly):
            self.report({"ERROR"}, "Select or name at least one low poly mesh")
            return {"CANCELLED"}

        depsgraph = None
        try:
            depsgraph = context.evaluated_depsgraph_get()
        except Exception:
            depsgraph = None
        if (
            low_objects
            and not auto_unwrap
            and any(not object_has_uvs(obj, depsgraph=depsgraph) for obj in low_objects)
        ):
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
            if active_info and not paths_match(active_info.get("project_dir"), project_dir):
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
            strip_uvs = auto_unwrap
            exported = export_fbx_objects(
                export_path,
                low_objects,
                prefs=prefs,
                strip_uvs=strip_uvs,
            )
            if not exported or not export_path.exists():
                self.report({"ERROR"}, "Low poly export failed or produced no FBX")
                return {"CANCELLED"}
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

        cage_export_path = None
        if prefs and getattr(prefs, "export_cage_poly", False):
            if cage_objects:
                cage_export_path = project_dir / BLENDER_CAGE_FILENAME
                exported = export_fbx_objects(cage_export_path, cage_objects, prefs=prefs)
                if not exported or not cage_export_path.exists():
                    self.report({"WARNING"}, "Cage export failed or produced no FBX")
                    cage_export_path = None

        force_new_token = ""
        if force_new_project:
            force_new_token = uuid.uuid4().hex

        manifest_path = project_manifest_path(project_dir)
        if manifest_path:
            ensure_dir(manifest_path.parent)
        try:
            sp_running = sp_running_for_send()
        except Exception:
            sp_running = False
        manifest = {
            "version": 1,
            "source": "blender",
            "project": get_project_name(context),
            "mesh_fbx": str(export_path) if (not prefs or prefs.export_low_poly) else old_mesh,
            "timestamp": time.time(),
        }
        if mesh_signature:
            manifest["mesh_signature"] = mesh_signature
        if prefs and not prefs.export_low_poly and old_manifest:
            old_signature = old_manifest.get("mesh_signature")
            if old_signature:
                manifest["mesh_signature"] = old_signature
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
        if cage_export_path:
            manifest["cage_mesh_fbx"] = str(cage_export_path)
        if prefs and getattr(prefs, "export_cage_poly", False):
            manifest["cage_mesh_exported"] = bool(cage_export_path)
        sp_project_settings, sp_project_warnings = build_sp_project_settings(prefs)
        for message in sp_project_warnings:
            self.report({"WARNING"}, message)
        if sp_project_settings:
            manifest["sp_project_settings"] = sp_project_settings
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
            if linked_sp_project and not already_open and should_force_open:
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
    bl_description = "Import the mesh and textures Substance Painter exported back into Blender"
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        prefs = get_prefs(context)
        if not prefs:
            return False
        return import_available(context, prefs)

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
        project_dir = project_dir_from_manifest_path(manifest_path)
        project_keys = manifest_project_keys(manifest, manifest_path=manifest_path)
        primary_project_key = project_keys[0] if project_keys else ""
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
            if primary_project_key and new_objects:
                tag_objects_with_project_key(
                    context,
                    new_objects,
                    primary_project_key,
                    clear_existing=True,
                )

        texture_paths = gather_texture_paths(manifest)
        targets = list(new_objects)
        project_targets = find_project_tag_targets(context, project_keys)
        if project_targets:
            existing = {obj.name for obj in targets}
            for obj in project_targets:
                if obj.name not in existing:
                    targets.append(obj)
                    existing.add(obj.name)
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
            matched_targets = find_texture_targets(context, grouped, project_keys=project_keys)
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
            targets = find_texture_targets(context, grouped, project_keys=project_keys)
            if not targets:
                all_meshes = [obj for obj in context.scene.objects if obj.type == "MESH"]
                if len(all_meshes) == 1:
                    targets = all_meshes
                    strict = True
        if not targets and grouped:
            self.report(
                {"WARNING"},
                "No mesh targets found; match material or object names to texture sets",
            )
        if targets and primary_project_key:
            tag_objects_with_project_key(
                context,
                targets,
                primary_project_key,
                clear_existing=bool(new_objects),
            )
        if texture_paths and targets:
            apply_textures_to_objects(targets, grouped, manifest=manifest, strict=strict)

        self.report({"INFO"}, "Imported assets from Substance Painter")
        return {"FINISHED"}


class GOB_OT_OpenExportFolder(Operator):
    bl_idname = "gob_sp.open_export_folder"
    bl_label = "Open Export Folder"
    bl_description = "Open the bridge project folder used to exchange files with Substance Painter"

    @classmethod
    def poll(cls, context):
        prefs = get_prefs(context)
        if not prefs:
            return False
        try:
            return bool(get_project_dir_fast(context, prefs))
        except Exception:
            return False

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
    bl_description = "Delete the whole bridge cache folder contents (keeps link/state files)"

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
    bl_description = "Delete the cache folder of the current bridge project only"

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
    bl_description = "Open the GoB SP Bridge Discord invite in your browser"

    def execute(self, _context):
        bpy.ops.wm.url_open(url=DISCORD_INVITE_URL)
        return {"FINISHED"}


class GOB_OT_OpenBugReport(Operator):
    bl_idname = "gob_sp.open_bug_report"
    bl_label = "Report Bug"
    bl_description = "Open the GoB SP Bridge issue tracker in your browser"

    def execute(self, _context):
        bpy.ops.wm.url_open(url=BUG_REPORT_URL)
        return {"FINISHED"}


class GOB_OT_CheckUpdates(Operator):
    bl_idname = "gob_sp.check_updates"
    bl_label = "Check for Updates"
    bl_description = "Check whether a newer GoB SP Bridge version is available"

    def execute(self, _context):
        start_update_check(show_no_update=True, show_popup=True)
        return {"FINISHED"}


class GOB_OT_OpenUpdateURL(Operator):
    bl_idname = "gob_sp.open_update_url"
    bl_label = "Open Update Download"
    bl_description = "Open the download page for the latest GoB SP Bridge version"

    def execute(self, _context):
        if not _last_update_info or not _last_update_info.get("download_url"):
            return {"CANCELLED"}
        bpy.ops.wm.url_open(url=_last_update_info["download_url"])
        return {"FINISHED"}


AUTOSPLIT_METHOD_LABELS = {
    "SMART": "Smart",
    "TRIANGLES": "Triangles only",
    "NAMES": "Names only",
}


class GOB_OT_PreviewSplit(Operator):
    bl_idname = "gob_sp.preview_split"
    bl_label = "Preview Auto Split"
    bl_description = "Preview how the selected meshes split into low and high poly"

    @classmethod
    def poll(cls, context):
        scene = context.scene
        if get_identify_mode(scene) != "AUTO_SPLIT":
            return False
        return any(obj.type == "MESH" for obj in context.selected_objects)

    def invoke(self, context, _event):
        scene = context.scene
        self._method = getattr(scene, "gob_sp_autosplit_method", "SMART")
        try:
            depsgraph = context.evaluated_depsgraph_get()
        except Exception:
            depsgraph = None
        low, high, warnings, reasons = classify_low_high(
            context,
            context.selected_objects,
            depsgraph,
            method=self._method,
        )
        self._entries = [
            (
                obj.name,
                evaluated_triangle_count(obj, depsgraph),
                "LOW",
                reasons.get(obj.name, ""),
            )
            for obj in low
        ]
        self._entries += [
            (
                obj.name,
                evaluated_triangle_count(obj, depsgraph),
                "HIGH",
                reasons.get(obj.name, ""),
            )
            for obj in high
        ]
        self._low_count = len(low)
        self._warnings = list(warnings)
        return context.window_manager.invoke_props_dialog(self, width=520)

    def draw(self, _context):
        layout = self.layout
        method_label = AUTOSPLIT_METHOD_LABELS.get(self._method, self._method)
        layout.label(text=f"Method: {method_label}")
        for message in self._warnings:
            row = layout.row()
            row.alert = True
            row.label(text=message, icon="ERROR")
        row = layout.row()
        low_col = row.column()
        low_col.label(text=f"LOW ({self._low_count})")
        high_col = row.column()
        high_col.label(text=f"HIGH ({len(self._entries) - self._low_count})")
        for name, count, side, reason in self._entries:
            col = low_col if side == "LOW" else high_col
            col.label(text=f"{name} · {count:,} tris · {side} · {reason}")

    def execute(self, _context):
        return {"FINISHED"}


class GOB_OT_RefreshTemplates(Operator):
    bl_idname = "gob_sp.refresh_templates"
    bl_label = "Refresh Templates"
    bl_description = "Rescan the Painter builtin and user template folders"

    def execute(self, _context):
        _refresh_template_cache()
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
        if prefs is None:
            layout.label(text="Enable the addon and restart Blender", icon="INFO")
            return
        status = get_ui_link_status(context, prefs)
        status_box = layout.box()
        linked_sp_project = status.get("linked_sp_project") or ""
        row = status_box.row()
        if linked_sp_project:
            row.label(text=Path(linked_sp_project).name, icon="LINKED")
        else:
            row.label(text="No linked Painter project", icon="UNLINKED")
        sp_running = status.get("sp_running")
        if sp_running is None:
            running_text = "Painter: unknown"
        elif sp_running:
            running_text = "Painter: running"
        else:
            running_text = "Painter: not running"
        status_box.label(text=running_text)
        status_box.label(
            text=elide_path(status.get("project_dir") or ""),
            icon="FILE_FOLDER",
        )
        if _bridge_conflict_info:
            conflict_row = status_box.row()
            conflict_row.alert = True
            conflict_row.label(
                text="Another Blender instance is using this bridge",
                icon="ERROR",
            )
        row = layout.row(align=True)
        row.scale_y = 1.4
        row.operator(GOB_OT_SendToSP.bl_idname, icon="EXPORT")
        row.operator(GOB_OT_ImportFromSP.bl_idname, icon="IMPORT")
        layout.operator(GOB_OT_OpenExportFolder.bl_idname, icon="FILE_FOLDER")
        scene = context.scene
        layout.prop(scene, "gob_sp_ui_tab", expand=True)
        tab = getattr(scene, "gob_sp_ui_tab", "SEND")
        if tab == "SETUP":
            self.draw_setup_tab(layout, prefs)
        elif tab == "CACHE":
            self.draw_cache_tab(context, layout, prefs)
        elif tab == "ABOUT":
            self.draw_about_tab(layout)
        else:
            self.draw_send_tab(context, layout, prefs)

    def draw_send_tab(self, context, layout, prefs):
        scene = context.scene
        col = layout.column(align=True)
        col.prop(prefs, "export_low_poly", text="Export Low")
        col.prop(prefs, "export_high_poly", text="Export High")
        col.prop(prefs, "export_cage_poly", text="Export Cage")
        col.prop(prefs, "export_selected_only", text="Only Selected")
        if not prefs.export_low_poly and not prefs.export_high_poly:
            layout.label(
                text="Enable Export Low or Export High to identify meshes",
                icon="INFO",
            )
        layout.label(text="Identify by:")
        layout.prop(scene, "gob_sp_identify_mode", expand=True)
        mode = get_identify_mode(scene)
        if mode == "COLLECTIONS":
            col = layout.column(align=True)
            col.prop_search(
                scene,
                "gob_sp_low_poly_collection",
                bpy.data,
                "collections",
                text="Low Collection",
            )
            col.prop_search(
                scene,
                "gob_sp_high_poly_collection",
                bpy.data,
                "collections",
                text="High Collection",
            )
            if getattr(prefs, "export_cage_poly", False):
                col.prop_search(
                    scene,
                    "gob_sp_cage_collection",
                    bpy.data,
                    "collections",
                    text="Cage Collection",
                )
        elif mode == "AUTO_SPLIT":
            layout.prop(scene, "gob_sp_autosplit_method", expand=True)
            layout.operator(GOB_OT_PreviewSplit.bl_idname, icon="VIEWZOOM")
        else:
            col = layout.column(align=True)
            col.prop(prefs, "low_poly_suffixes")
            col.prop(prefs, "high_poly_suffixes")
            if getattr(prefs, "export_cage_poly", False):
                col.prop(prefs, "cage_poly_suffixes")
            layout.label(
                text="Painter's baker matches meshes by _low/_high names",
                icon="INFO",
            )
        plan = get_cached_export_plan(context, prefs)
        if getattr(prefs, "export_selected_only", False):
            summary = f"{len(plan['low'])} low · {len(plan['high'])} high selected"
        else:
            summary = f"{len(plan['low']) + len(plan['high'])} meshes (scene)"
        layout.label(text=summary, icon="MESH_DATA")
        warnings = [message for message in plan["warnings"] if message]
        if not warnings:
            if _last_export_warning:
                _set_export_warning("")
        elif _last_export_warning and _last_export_warning not in warnings:
            warnings.append(_last_export_warning)
        for message in warnings:
            row = layout.row()
            row.alert = True
            row.label(text=message, icon="ERROR")

    def draw_setup_tab(self, layout, prefs):
        row = layout.row(align=True)
        row.prop(prefs, "sp_project_template_picker", text="Template")
        row.operator(GOB_OT_RefreshTemplates.bl_idname, text="", icon="FILE_REFRESH")
        layout.prop(prefs, "sp_color_management", text="Color Management")
        if prefs.sp_color_management == "OCIO_CUSTOM":
            layout.prop(prefs, "sp_color_config_path", text="")
        template_key = getattr(prefs, "sp_project_template", "NONE") or "NONE"
        if template_key != "NONE" and prefs.sp_color_management != "PAINTER_DEFAULT":
            layout.label(
                text="The template may override the color management choice",
                icon="INFO",
            )
        layout.prop(prefs, "sp_default_resolution", text="Default Resolution")
        layout.prop(prefs, "sp_project_workflow", text="Project Workflow")
        layout.prop(prefs, "sp_normal_map_format", text="Normal Map Format")
        layout.prop(prefs, "sp_auto_unwrap")
        if prefs.sp_auto_unwrap:
            layout.prop(prefs, "sp_unwrap_mode", text="Unwrap Mode")
        layout.label(text="Used when a new Painter project is created", icon="INFO")
        layout.separator()
        col = layout.column(align=True)
        col.prop(prefs, "fbx_export_scale")
        col.prop(prefs, "fbx_apply_unit_scale")
        layout.prop(
            prefs,
            "force_new_sp_project_on_send",
            text="Always create a new Painter project",
        )

    def draw_cache_tab(self, context, layout, prefs):
        global_size, local_size = get_cached_cache_sizes(context, prefs)
        layout.label(text=f"Global cache: {format_bytes(global_size)}")
        layout.label(text=f"Project cache: {format_bytes(local_size)}")
        layout.prop(prefs, "auto_clear_cache")
        row = layout.row()
        row.enabled = prefs.auto_clear_cache
        row.prop(prefs, "cache_limit_gb")
        if prefs.auto_clear_cache:
            layout.label(text="Auto clear keeps the current project", icon="INFO")
        row = layout.row(align=True)
        row.operator(GOB_OT_ClearCacheGlobal.bl_idname, icon="TRASH")
        row.operator(GOB_OT_ClearCacheLocal.bl_idname, icon="TRASH")

    def draw_about_tab(self, layout):
        layout.label(text=f"Version {local_version_string()}")
        update_row = layout.row(align=True)
        update_row.label(text=_update_status_text)
        update_row.operator(GOB_OT_CheckUpdates.bl_idname, text="Check")
        if _last_update_info and _last_update_info.get("download_url"):
            update_row.operator(GOB_OT_OpenUpdateURL.bl_idname, text="Download")
        link_row = layout.row(align=True)
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
    GOB_OT_PreviewSplit,
    GOB_OT_RefreshTemplates,
    GOB_PT_Panel,
)


def register():
    for cls in classes:
        bpy.utils.register_class(cls)
    bpy.types.Scene.gob_sp_ui_tab = EnumProperty(
        name="Panel Tab",
        description="Which section of the GoB SP panel to show",
        items=(
            ("SEND", "Send", "Export meshes and send them to Substance Painter"),
            ("SETUP", "Project Setup", "Painter project creation settings"),
            ("CACHE", "Cache", "Bridge cache settings"),
            ("ABOUT", "About", "Version, updates and community links"),
        ),
        default="SEND",
    )
    bpy.types.Scene.gob_sp_identify_mode = EnumProperty(
        name="Low/High Identification",
        description="How to tell low poly meshes apart from high poly meshes",
        items=(
            ("SUFFIXES", "By Name", "Match meshes by name suffixes"),
            ("COLLECTIONS", "By Collection", "Use the low/high collections (suffix matching as fallback)"),
            ("AUTO_SPLIT", "Auto Split (Beta)", "Split meshes automatically (beta)"),
        ),
        default="SUFFIXES",
        update=_on_export_settings_update,
    )
    bpy.types.Scene.gob_sp_autosplit_method = EnumProperty(
        name="Auto Split Method",
        description="Which signals Auto Split uses to tell low and high poly meshes apart",
        items=(
            ("SMART", "Smart", "Use names, tags, modifiers, positions and triangle counts"),
            ("TRIANGLES", "Triangles only", "Split by evaluated triangle counts only"),
            ("NAMES", "Names only", "Split by _low/_high name pairs only"),
        ),
        default="SMART",
        update=_on_export_settings_update,
    )
    bpy.types.Scene.gob_sp_low_poly_collection = PointerProperty(
        name="Low Poly Collection",
        description="Collection to export as low poly (overrides suffix matching)",
        type=bpy.types.Collection,
        poll=_scene_collection_poll,
        update=_on_export_settings_update,
    )
    bpy.types.Scene.gob_sp_high_poly_collection = PointerProperty(
        name="High Poly Collection",
        description="Collection to export as high poly (overrides suffix matching)",
        type=bpy.types.Collection,
        poll=_scene_collection_poll,
        update=_on_export_settings_update,
    )
    bpy.types.Scene.gob_sp_cage_collection = PointerProperty(
        name="Cage Collection",
        description="Collection to export as cage (overrides suffix matching)",
        type=bpy.types.Collection,
        poll=_scene_collection_poll,
        update=_on_export_settings_update,
    )
    if _init_scene_ui_prefs not in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.append(_init_scene_ui_prefs)
    if _update_active_blender_info not in bpy.app.handlers.save_post:
        bpy.app.handlers.save_post.append(_update_active_blender_info)
    if not bpy.app.timers.is_registered(_init_scene_ui_prefs):
        bpy.app.timers.register(_init_scene_ui_prefs, first_interval=0.1, persistent=True)
    if not bpy.app.timers.is_registered(_active_blender_heartbeat):
        bpy.app.timers.register(_active_blender_heartbeat, first_interval=1.0, persistent=True)
    start_update_check(show_popup=False)


def unregister():
    global _update_check_in_progress
    global _update_check_result
    if _init_scene_ui_prefs in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.remove(_init_scene_ui_prefs)
    if _update_active_blender_info in bpy.app.handlers.save_post:
        bpy.app.handlers.save_post.remove(_update_active_blender_info)
    if bpy.app.timers.is_registered(_init_scene_ui_prefs):
        bpy.app.timers.unregister(_init_scene_ui_prefs)
    if bpy.app.timers.is_registered(_active_blender_heartbeat):
        bpy.app.timers.unregister(_active_blender_heartbeat)
    if bpy.app.timers.is_registered(_update_poll):
        bpy.app.timers.unregister(_update_poll)
    _update_check_in_progress = False
    _update_check_result = None
    if hasattr(bpy.types.Scene, "gob_sp_identify_mode"):
        del bpy.types.Scene.gob_sp_identify_mode
    if hasattr(bpy.types.Scene, "gob_sp_ui_tab"):
        del bpy.types.Scene.gob_sp_ui_tab
    if hasattr(bpy.types.Scene, "gob_sp_autosplit_method"):
        del bpy.types.Scene.gob_sp_autosplit_method
    if hasattr(bpy.types.Scene, "gob_sp_cage_collection"):
        del bpy.types.Scene.gob_sp_cage_collection
    if hasattr(bpy.types.Scene, "gob_sp_high_poly_collection"):
        del bpy.types.Scene.gob_sp_high_poly_collection
    if hasattr(bpy.types.Scene, "gob_sp_low_poly_collection"):
        del bpy.types.Scene.gob_sp_low_poly_collection
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
