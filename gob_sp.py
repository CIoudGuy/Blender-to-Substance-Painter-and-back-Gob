import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets, QtNetwork

import substance_painter as sp
import substance_painter.event as sp_event


BRIDGE_ENV_VAR = "GOB_SP_BRIDGE_DIR"
BRIDGE_ROOT_HINT_FILENAME = "bridge_root.json"
BRIDGE_SHARED_HINT_DIRNAME = ".gob_sp_bridge"
MANIFEST_FILENAME = "bridge.json"
BLENDER_EXPORT_FILENAME = "b2sp.fbx"
SP_EXPORT_FILENAME = "sp2b.fbx"
LOG_FILENAME = "sp_export_log.txt"
PROJECT_META_DIRNAME = ".gob_meta"
ACTIVE_SP_INFO_FILENAME = "active_sp.json"
ACTIVE_BLENDER_INFO_FILENAME = "active_blender.json"
LINKS_FILENAME = "project_links.json"
TEMP_DIRNAME = ".gob_temp"
TEMP_SP_PREFIX = "gob_unsaved_sp_"
TEMP_BLENDER_PREFIX = "gob_unsaved_bl_"
TEMP_SP_SUFFIX = ".spp"
TEMP_BLENDER_SUFFIX = ".blend"
HIGH_POLY_RETRY_DELAY_MS = 800
HIGH_POLY_RETRY_COUNT = 60
ACTIVE_BLENDER_INFO_MAX_AGE = 120.0
FORCE_NEW_TOKEN_ENV = "GOB_SP_FORCE_NEW_TOKEN"
FORCE_NEW_TOKEN_ARG_PREFIXES = ("--gob-force-new-token=", "--gob-force-new=")
UPDATE_URL = (
    "https://raw.githubusercontent.com/CIoudGuy/Blender-to-Substance-Painter-and-back-Gob/"
    "refs/heads/main/version.json"
)
BUG_REPORT_URL = (
    "https://github.com/CIoudGuy/Blender-to-Substance-Painter-and-back-Gob/issues"
)
PLUGIN_VERSION = "0.2.0"

EXPORT_FORMATS = [
    ("png", "PNG"),
    ("tga", "TGA"),
    ("tiff", "TIFF"),
    ("exr", "EXR"),
]
EXPORT_BIT_DEPTHS = [
    ("8", "8-bit"),
    ("16", "16-bit"),
]
EXPORT_RESOLUTIONS = [
    (7, "128"),
    (8, "256"),
    (9, "512"),
    (10, "1024"),
    (11, "2048"),
    (12, "4096"),
    (13, "8192"),
]
PADDING_ALGORITHMS = [
    ("passthrough", "Passthrough"),
    ("color", "Color"),
    ("transparent", "Transparent"),
    ("diffusion", "Diffusion"),
    ("infinite", "Infinite"),
]
DEFAULT_EXPORT_SETTINGS = {
    "file_format": "png",
    "bit_depth": "8",
    "size_log2": 11,
    "padding_algorithm": "infinite",
    "dilation_distance": 16,
    "dithering": False,
}

CUSTOM_EXPORT_PRESETS = []
SETTINGS_FILENAME = "gob_sp_settings.json"
SETTINGS_VERSION = 1
PROJECT_SETTINGS_FILENAME = "gob_sp_project_settings.json"
DEFAULT_USER_PRESET_NAME = "Default"
UPDATE_IGNORE_VERSION_KEY = "update_ignore_version"
_temp_session_id = None
_temp_sp_project_file = None
_temp_blender_file = None
_last_sp_project_file = None
_project_dir_cache = {}
_force_new_token = ""


def _rgb_channels(src_type, src_name):
    return [
        {
            "destChannel": "R",
            "srcChannel": "R",
            "srcMapType": src_type,
            "srcMapName": src_name,
        },
        {
            "destChannel": "G",
            "srcChannel": "G",
            "srcMapType": src_type,
            "srcMapName": src_name,
        },
        {
            "destChannel": "B",
            "srcChannel": "B",
            "srcMapType": src_type,
            "srcMapName": src_name,
        },
    ]


def _gray_channels(src_type, src_name):
    return [
        {
            "destChannel": "L",
            "srcChannel": "L",
            "srcMapType": src_type,
            "srcMapName": src_name,
        }
    ]


def _map_params():
    return {
        "fileFormat": "png",
        "bitDepth": "8",
        "dithering": False,
        "sizeLog2": 11,
        "paddingAlgorithm": "infinite",
        "dilationDistance": 16,
    }


CUSTOM_EXPORT_PRESETS.extend([
    {
        "name": "Roblox PBR (OpenGL)",
        "maps": [
            {
                "fileName": "$textureSet_Color",
                "channels": _rgb_channels("documentMap", "basecolor"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Metalness",
                "channels": _gray_channels("documentMap", "metallic"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Roughness",
                "channels": _gray_channels("documentMap", "roughness"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Normal",
                "channels": _rgb_channels("virtualMap", "Normal_OpenGL"),
                "parameters": _map_params(),
            },
        ],
    },
    {
        "name": "Roblox PBR (DirectX)",
        "maps": [
            {
                "fileName": "$textureSet_Color",
                "channels": _rgb_channels("documentMap", "basecolor"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Metalness",
                "channels": _gray_channels("documentMap", "metallic"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Roughness",
                "channels": _gray_channels("documentMap", "roughness"),
                "parameters": _map_params(),
            },
            {
                "fileName": "$textureSet_Normal",
                "channels": _rgb_channels("virtualMap", "Normal_DirectX"),
                "parameters": _map_params(),
            },
        ],
    },
])


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
    docs = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.DocumentsLocation)
    if docs:
        return os.path.join(docs, "GoB_SP_Bridge")
    win_docs = windows_documents_dir()
    if win_docs:
        return os.path.join(win_docs, "GoB_SP_Bridge")
    return os.path.join(os.path.expanduser("~"), "Documents", "GoB_SP_Bridge")


def documents_bridge_root():
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
    docs = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.DocumentsLocation)
    if docs:
        return Path(docs) / "GoB_SP_Bridge"
    win_docs = windows_documents_dir()
    if win_docs:
        return Path(win_docs) / "GoB_SP_Bridge"
    return Path(os.path.expanduser("~")) / "Documents" / "GoB_SP_Bridge"


def bridge_root_hint_path():
    return Path(default_bridge_dir()).expanduser() / BRIDGE_ROOT_HINT_FILENAME


def shared_bridge_root_hint_path():
    return Path.home() / BRIDGE_SHARED_HINT_DIRNAME / BRIDGE_ROOT_HINT_FILENAME


def read_bridge_root_hint():
    for path in (bridge_root_hint_path(), shared_bridge_root_hint_path()):
        if not path.exists():
            continue
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(data, dict):
            continue
        root = data.get("bridge_root")
        if root:
            return Path(root).expanduser()
    return None


def write_bridge_root_hint(root_path):
    if not root_path:
        return
    payload = {"bridge_root": str(Path(root_path).expanduser())}
    for path in (bridge_root_hint_path(), shared_bridge_root_hint_path()):
        try:
            ensure_dir(path.parent)
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, indent=2, ensure_ascii=True)
        except OSError:
            continue


def settings_path():
    base = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.AppDataLocation)
    if not base:
        base = os.path.join(os.path.expanduser("~"), ".gob_sp_bridge")
    return Path(base) / SETTINGS_FILENAME


def load_settings():
    path = settings_path()
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_settings(data):
    path = settings_path()
    ensure_dir(path.parent)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, ensure_ascii=True)
    except OSError:
        return


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


def project_settings_path(project_dir=None):
    if project_dir is None:
        try:
            project_dir = get_project_dir()
        except Exception:
            return None
    if not project_dir:
        return None
    return project_meta_dir(project_dir) / PROJECT_SETTINGS_FILENAME


def load_project_settings(project_dir=None):
    base_dir = project_dir
    if base_dir is None:
        try:
            base_dir = get_project_dir()
        except Exception:
            return {}
    if not base_dir:
        return {}
    path = project_meta_dir(base_dir) / PROJECT_SETTINGS_FILENAME
    if not path.exists():
        legacy_path = Path(base_dir) / PROJECT_SETTINGS_FILENAME
        if legacy_path.exists():
            path = legacy_path
        else:
            return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_project_settings(data, project_dir=None):
    path = project_settings_path(project_dir)
    if not path:
        return
    ensure_dir(path.parent)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, ensure_ascii=True)
    except OSError:
        return


def update_project_settings(update_data, project_dir=None):
    if not update_data:
        return
    data = load_project_settings(project_dir)
    if not isinstance(data, dict):
        data = {}
    data.update(update_data)
    save_project_settings(data, project_dir=project_dir)


def get_update_ignore_version():
    data = load_settings()
    value = data.get(UPDATE_IGNORE_VERSION_KEY)
    return str(value) if value else ""


def set_update_ignore_version(version):
    if not version:
        return
    data = load_settings()
    data[UPDATE_IGNORE_VERSION_KEY] = str(version)
    save_settings(data)


def load_persistent_state(project_dir=None):
    data = load_settings()
    project_data = load_project_settings(project_dir)
    last_settings = project_data.get("last_settings")
    if last_settings is None:
        last_settings = data.get("last_settings", {})
    return {
        "version": data.get("version", SETTINGS_VERSION),
        "last_settings": last_settings,
        "user_presets": data.get("user_presets", []),
    }


def save_persistent_state(last_settings=None, user_presets=None, project_dir=None):
    data = load_settings()
    if not isinstance(data, dict):
        data = {}
    if last_settings is not None:
        data["last_settings"] = last_settings
    if user_presets is not None:
        data["user_presets"] = user_presets
    data["version"] = SETTINGS_VERSION
    save_settings(data)
    if last_settings is not None:
        project_data = load_project_settings(project_dir)
        if not isinstance(project_data, dict):
            project_data = {}
        project_data["version"] = SETTINGS_VERSION
        project_data["last_settings"] = last_settings
        save_project_settings(project_data, project_dir=project_dir)


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


def get_sp_name(obj):
    if obj is None:
        return ""
    try:
        value = getattr(obj, "name", None)
        if value is not None and not callable(value):
            return str(value)
    except Exception:
        pass
    try:
        name_fn = getattr(obj, "name", None)
        if callable(name_fn):
            return str(name_fn())
    except Exception:
        pass
    return ""


def get_all_texture_sets():
    try:
        attr = getattr(sp.textureset, "all_texture_sets", None)
        if callable(attr):
            return list(attr())
        if attr is None:
            return []
        return list(attr)
    except Exception:
        return []


def get_all_stacks(texset):
    if not texset:
        return []
    try:
        attr = getattr(texset, "all_stacks", None)
        if callable(attr):
            return list(attr())
        if attr is None:
            return []
        return list(attr)
    except Exception:
        return []


def get_bridge_root():
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        return Path(env_path).expanduser()
    hint = read_bridge_root_hint()
    if hint:
        return hint
    return Path(default_bridge_dir()).expanduser()


def get_project_name():
    try:
        name = get_sp_name(sp.project)
    except Exception:
        name = ""
    if not name:
        name = "untitled"
    return sanitize_name(name)


def get_project_dir():
    base_dir = get_bridge_root() / get_project_name()
    sp_project_file = ""
    try:
        if sp.project.is_open():
            sp_project_file = get_sp_project_file_path_or_temp()
    except Exception:
        sp_project_file = ""
    return resolve_project_dir_for_sp(sp_project_file, base_dir)


def get_sp_project_file_path():
    candidates = (
        "file_path",
        "filePath",
        "project_file_path",
        "project_file",
        "projectFile",
        "project_filepath",
        "projectFilePath",
        "projectPath",
        "project_path",
        "filepath",
        "path",
    )
    for name in candidates:
        attr = getattr(sp.project, name, None)
        if attr is None:
            continue
        try:
            value = attr() if callable(attr) else attr
        except Exception:
            continue
        if value:
            return str(value)
    return ""


def temp_session_id():
    global _temp_session_id
    if _temp_session_id is None:
        _temp_session_id = f"{os.getpid()}_{int(time.time())}"
    return _temp_session_id


def bridge_temp_dir():
    return get_bridge_root() / TEMP_DIRNAME


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


def temp_sp_project_file_path():
    global _temp_sp_project_file
    if _temp_sp_project_file:
        return _temp_sp_project_file
    temp_dir = bridge_temp_dir()
    filename = f"{TEMP_SP_PREFIX}{temp_session_id()}{TEMP_SP_SUFFIX}"
    temp_path = temp_dir / filename
    ensure_placeholder_file(temp_path)
    _temp_sp_project_file = str(temp_path)
    return _temp_sp_project_file


def temp_blender_file_path():
    global _temp_blender_file
    if _temp_blender_file:
        return _temp_blender_file
    temp_dir = bridge_temp_dir()
    filename = f"{TEMP_BLENDER_PREFIX}{temp_session_id()}{TEMP_BLENDER_SUFFIX}"
    temp_path = temp_dir / filename
    ensure_placeholder_file(temp_path)
    _temp_blender_file = str(temp_path)
    return _temp_blender_file


def get_sp_project_file_path_or_temp():
    sp_path = get_sp_project_file_path()
    if sp_path:
        return sp_path
    return temp_sp_project_file_path()


def is_temp_file(path, prefix, suffix):
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
        return normalize_path(path_obj.parent).lower() == normalize_path(bridge_temp_dir()).lower()
    except Exception:
        return False


def is_temp_sp_project_file(path):
    return is_temp_file(path, TEMP_SP_PREFIX, TEMP_SP_SUFFIX)


def is_temp_blender_file(path):
    return is_temp_file(path, TEMP_BLENDER_PREFIX, TEMP_BLENDER_SUFFIX)


def project_meta_dir(project_dir):
    return Path(project_dir) / PROJECT_META_DIRNAME


def project_dir_cache_key(sp_project_file):
    if not sp_project_file:
        return ""
    return normalize_path_key(sp_project_file)


def cached_project_dir(sp_project_file):
    key = project_dir_cache_key(sp_project_file)
    if not key:
        return None
    cached = _project_dir_cache.get(key)
    return Path(cached) if cached else None


def set_cached_project_dir(sp_project_file, project_dir):
    key = project_dir_cache_key(sp_project_file)
    if not key or not project_dir:
        return
    _project_dir_cache[key] = str(project_dir)


def manifest_matches_sp_project_file(manifest, sp_project_file):
    if not manifest or not sp_project_file:
        return False
    manifest_sp = manifest.get("sp_project_file") or manifest.get("sp_project_path")
    return bool(manifest_sp and paths_match(manifest_sp, sp_project_file))


def resolve_project_dir_for_sp(sp_project_file, base_dir):
    cached = cached_project_dir(sp_project_file)
    if cached:
        return cached
    if sp_project_file and base_dir.exists():
        manifest_path = find_project_manifest_path(base_dir)
        manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
        if manifest_matches_sp_project_file(manifest, sp_project_file):
            set_cached_project_dir(sp_project_file, base_dir)
            return base_dir
    if sp_project_file:
        manifest_path = find_manifest_for_sp_project(
            get_candidate_bridge_roots(),
            sp_project_file,
        )
        if manifest_path:
            project_dir = project_dir_from_manifest_path(manifest_path)
            set_cached_project_dir(sp_project_file, project_dir)
            return project_dir
        project_dir = unique_project_dir(base_dir, sp_project_file)
        set_cached_project_dir(sp_project_file, project_dir)
        return project_dir
    return base_dir


def unique_project_dir(base_dir, sp_project_file):
    if not base_dir.exists():
        return base_dir
    if sp_project_file:
        manifest_path = find_project_manifest_path(base_dir)
        manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
        if manifest_matches_sp_project_file(manifest, sp_project_file):
            return base_dir
    root = base_dir.parent
    base_name = base_dir.name
    index = 1
    while True:
        candidate = root / f"{base_name}{index}"
        if candidate.exists():
            if sp_project_file:
                manifest_path = find_project_manifest_path(candidate)
                manifest = read_manifest(manifest_path) if manifest_path and manifest_path.exists() else None
                if manifest_matches_sp_project_file(manifest, sp_project_file):
                    return candidate
            index += 1
            continue
        return candidate


def project_dir_for_send(sp_project_file):
    base_dir = get_bridge_root() / get_project_name()
    if sp_project_file:
        cached = cached_project_dir(sp_project_file)
        if cached:
            return cached
        manifest_path = find_manifest_for_sp_project(
            get_candidate_bridge_roots(),
            sp_project_file,
        )
        if manifest_path:
            project_dir = project_dir_from_manifest_path(manifest_path)
            set_cached_project_dir(sp_project_file, project_dir)
            return project_dir
    project_dir = unique_project_dir(base_dir, sp_project_file)
    if sp_project_file:
        set_cached_project_dir(sp_project_file, project_dir)
    return project_dir


def normalize_normal_map_format(value):
    if value is None:
        return None
    text = str(value).lower()
    if "directx" in text or "d3d" in text or text == "dx":
        return "directx"
    if "opengl" in text or "ogl" in text or text == "gl":
        return "opengl"
    return None


def get_sp_normal_map_format():
    candidates = ("normal_map_format", "normalMapFormat", "normal_format", "normalFormat")
    containers = [sp.project]
    for name in (
        "project_settings",
        "projectSettings",
        "settings",
        "get_project_settings",
        "getProjectSettings",
        "get_settings",
        "getSettings",
    ):
        attr = getattr(sp.project, name, None)
        if attr is None:
            continue
        try:
            value = attr() if callable(attr) else attr
        except Exception:
            continue
        if value is not None:
            containers.append(value)
    for container in containers:
        for name in candidates:
            attr = getattr(container, name, None)
            if attr is None:
                continue
            try:
                value = attr() if callable(attr) else attr
            except Exception:
                continue
            normalized = normalize_normal_map_format(value)
            if normalized:
                return normalized
    return None


def normalize_path(path):
    if not path:
        return ""
    try:
        return os.path.abspath(os.path.expanduser(str(path)))
    except Exception:
        return str(path)


def normalize_path_key(path):
    return normalize_path(path).lower()


def link_registry_paths():
    roots = []
    docs_root = documents_bridge_root()
    if docs_root:
        roots.append(Path(docs_root))
    for root in get_candidate_bridge_roots():
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


def load_link_registry():
    for path in link_registry_paths():
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


def save_link_registry(data):
    paths = link_registry_paths()
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


def update_link_registry(sp_project_file=None, blender_file=None, update_blender_link=True):
    if not sp_project_file or not blender_file:
        return
    data = load_link_registry()
    sp_key = normalize_path_key(sp_project_file)
    bl_key = normalize_path_key(blender_file)
    sp_map = data.get("sp_to_blender")
    if not isinstance(sp_map, dict):
        sp_map = {}
    bl_map = data.get("blender_to_sp")
    if not isinstance(bl_map, dict):
        bl_map = {}
    sp_map[sp_key] = str(blender_file)
    if update_blender_link:
        bl_map[bl_key] = str(sp_project_file)
    data["sp_to_blender"] = sp_map
    data["blender_to_sp"] = bl_map
    save_link_registry(data)


def is_force_new_project_dir(project_dir):
    if not project_dir:
        return False
    manifest_path = find_project_manifest_path(project_dir)
    if not manifest_path or not manifest_path.exists():
        return False
    manifest = read_manifest(manifest_path)
    return bool(isinstance(manifest, dict) and manifest.get("force_new_project"))


def paths_match(left, right):
    if not left or not right:
        return False
    return normalize_path(left).lower() == normalize_path(right).lower()


def find_manifest_for_sp_project(bridge_roots, sp_project_file, source=None):
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


def find_mesh_in_roots(bridge_roots, project_name, filename):
    if not project_name or not filename:
        return None
    for root in bridge_roots:
        if not root:
            continue
        candidate = Path(root) / project_name / filename
        try:
            if candidate.is_file():
                return str(candidate)
        except OSError:
            continue
    return None


def read_linked_blender_file(project_dir):
    blender_file = ""
    sp_project_file = get_sp_project_file_path_or_temp()
    if sp_project_file:
        registry = load_link_registry()
        blender_file = registry.get("sp_to_blender", {}).get(
            normalize_path_key(sp_project_file)
        )
        if blender_file:
            return str(blender_file)
        manifest_path = find_manifest_for_sp_project(
            get_candidate_bridge_roots(),
            sp_project_file,
        )
        if manifest_path:
            manifest = read_manifest(manifest_path)
            if manifest:
                blender_file = manifest.get("blender_file") or ""
                if blender_file:
                    return str(blender_file)
    if project_dir:
        manifest_path = find_project_manifest_path(project_dir)
        if manifest_path and manifest_path.exists():
            manifest = read_manifest(manifest_path)
            if manifest:
                manifest_sp = manifest.get("sp_project_file") or manifest.get("sp_project_path")
                if manifest_sp and sp_project_file and not paths_match(manifest_sp, sp_project_file):
                    pass
                else:
                    blender_file = manifest.get("blender_file") or ""
                    if blender_file:
                        return str(blender_file)
        project_data = load_project_settings(project_dir)
        if isinstance(project_data, dict):
            blender_file = project_data.get("linked_blender_file") or ""
            if blender_file:
                return str(blender_file)
    return ""


def resolve_primary_sp_project_for_blender(blender_file, current_sp_project_file):
    if not blender_file:
        return ""
    registry = load_link_registry()
    candidate = registry.get("blender_to_sp", {}).get(normalize_path_key(blender_file))
    if not candidate:
        return ""
    if current_sp_project_file and paths_match(candidate, current_sp_project_file):
        return ""
    return str(candidate)


def update_manifest_sp_project_file(old_sp_project_file, new_sp_project_file):
    if not old_sp_project_file or not new_sp_project_file:
        return
    manifest_path = find_manifest_for_sp_project(
        get_candidate_bridge_roots(),
        old_sp_project_file,
    )
    if not manifest_path:
        return
    manifest = read_manifest(manifest_path)
    if not isinstance(manifest, dict):
        return
    manifest["sp_project_file"] = str(new_sp_project_file)
    project_dir = project_dir_from_manifest_path(manifest_path)
    target_path = project_manifest_path(project_dir)
    if target_path:
        ensure_dir(target_path.parent)
        write_manifest(target_path, manifest)


def write_manifest_sp_project_file(manifest, project_dir, sp_project_file):
    if not isinstance(manifest, dict) or not project_dir or not sp_project_file:
        return
    manifest["sp_project_file"] = str(sp_project_file)
    target_path = project_manifest_path(project_dir)
    if not target_path:
        return
    ensure_dir(target_path.parent)
    write_manifest(target_path, manifest)


def sync_saved_sp_project_file():
    global _last_sp_project_file
    if not sp.project.is_open():
        _last_sp_project_file = None
        return
    current_real = get_sp_project_file_path()
    current = current_real or temp_sp_project_file_path()
    if _last_sp_project_file and current_real and not paths_match(current_real, _last_sp_project_file):
        project_dir = get_project_dir()
        linked_blender_file = read_linked_blender_file(project_dir)
        force_new = is_force_new_project_dir(project_dir)
        if linked_blender_file:
            update_link_registry(
                sp_project_file=current_real,
                blender_file=linked_blender_file,
                update_blender_link=not force_new,
            )
        update_manifest_sp_project_file(_last_sp_project_file, current_real)
        set_cached_project_dir(current_real, get_project_dir())
    _last_sp_project_file = current


def find_blender_exe():
    for env_var in ("BLENDER_EXE", "BLENDER_EXECUTABLE", "BLENDER_PATH"):
        env_path = os.environ.get(env_var)
        if not env_path:
            continue
        env_candidate = Path(env_path).expanduser()
        if env_candidate.is_file():
            return str(env_candidate)
        if sys.platform == "darwin" and env_candidate.suffix.lower() == ".app" and env_candidate.is_dir():
            return str(env_candidate)

    if os.name == "nt":
        program_files = os.environ.get("ProgramFiles", r"C:\\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)")
        for base in (program_files, program_files_x86):
            if not base:
                continue
            base_path = Path(base) / "Blender Foundation"
            if not base_path.exists():
                continue
            direct = base_path / "Blender" / "blender.exe"
            if direct.is_file():
                return str(direct)
            for candidate in sorted(base_path.glob("Blender*/blender.exe"), reverse=True):
                if candidate.is_file():
                    return str(candidate)
    elif sys.platform == "darwin":
        app_candidates = [
            Path("/Applications/Blender.app"),
            Path.home() / "Applications" / "Blender.app",
        ]
        for candidate in app_candidates:
            if candidate.is_dir():
                return str(candidate)
        for root in (Path("/Applications"), Path.home() / "Applications"):
            if not root.exists():
                continue
            for app in sorted(root.glob("Blender*.app")):
                if app.is_dir():
                    return str(app)
    else:
        blender_bin = shutil.which("blender")
        if blender_bin:
            return blender_bin
    return ""


def open_linked_blender_file(path, project_dir=None):
    if not path:
        return False
    if blender_project_is_open(path, project_dir=project_dir):
        return True
    blender_exe = find_blender_exe()
    try:
        if sys.platform == "darwin":
            if blender_exe:
                if blender_exe.lower().endswith(".app") and Path(blender_exe).is_dir():
                    subprocess.Popen(["open", "-a", blender_exe, str(path)])
                    return True
                if Path(blender_exe).is_file():
                    subprocess.Popen([blender_exe, str(path)])
                    return True
            subprocess.Popen(["open", str(path)])
            return True
        if os.name == "nt":
            if blender_exe and Path(blender_exe).is_file():
                subprocess.Popen([blender_exe, str(path)])
                return True
            if hasattr(os, "startfile"):
                os.startfile(str(path))
                return True
            subprocess.Popen(["cmd", "/c", "start", "", str(path)])
            return True
        if blender_exe and Path(blender_exe).is_file():
            subprocess.Popen([blender_exe, str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
        return True
    except Exception:
        try:
            return QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(path)))
        except Exception:
            return False


def active_sp_info_paths(project_dir=None):
    paths = [get_bridge_root() / ACTIVE_SP_INFO_FILENAME]
    docs_root = documents_bridge_root()
    if docs_root:
        paths.append(Path(docs_root) / ACTIVE_SP_INFO_FILENAME)
    if project_dir:
        paths.append(project_meta_dir(project_dir) / ACTIVE_SP_INFO_FILENAME)
    unique = []
    seen = set()
    for path in paths:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique


def write_active_sp_info():
    try:
        sync_saved_sp_project_file()
    except Exception:
        pass
    info = {
        "timestamp": time.time(),
        "project_open": False,
    }
    project_dir = None
    try:
        if sp.project.is_open():
            info["project_open"] = True
            info["project_name"] = get_project_name()
            project_dir = get_project_dir()
            info["project_dir"] = str(project_dir)
            sp_project_file = get_sp_project_file_path_or_temp()
            if sp_project_file:
                info["sp_project_file"] = sp_project_file
            blender_file = read_linked_blender_file(project_dir)
            if blender_file:
                info["blender_file"] = blender_file
    except Exception:
        pass
    for path in active_sp_info_paths(project_dir):
        ensure_dir(path.parent)
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(info, handle, indent=2, ensure_ascii=True)
        except OSError:
            continue


def active_blender_info_paths(project_dir=None):
    paths = [get_bridge_root() / ACTIVE_BLENDER_INFO_FILENAME]
    docs_root = documents_bridge_root()
    if docs_root:
        paths.append(Path(docs_root) / ACTIVE_BLENDER_INFO_FILENAME)
    if project_dir:
        paths.append(project_meta_dir(project_dir) / ACTIVE_BLENDER_INFO_FILENAME)
    unique = []
    seen = set()
    for path in paths:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique


def read_active_blender_info(path):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
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
        "timestamp": timestamp,
        "blender_file": data.get("blender_file"),
        "project_dir": data.get("project_dir"),
        "project_name": data.get("project_name"),
    }


def find_active_blender_info(project_dir=None, max_age=ACTIVE_BLENDER_INFO_MAX_AGE):
    now = time.time()
    best = None
    best_time = 0.0
    for path in active_blender_info_paths(project_dir):
        if not path.exists():
            continue
        info = read_active_blender_info(path)
        if not info:
            continue
        ts = info.get("timestamp", 0.0) or 0.0
        if max_age and ts and now - ts > max_age:
            continue
        if ts > best_time:
            best_time = ts
            best = info
    return best


def blender_project_is_open(path, project_dir=None):
    if not path:
        return False
    info = find_active_blender_info(project_dir)
    if not info:
        return False
    active_path = info.get("blender_file") or ""
    return bool(active_path and paths_match(active_path, path))


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


def get_candidate_bridge_roots():
    roots = []
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        roots.append(Path(env_path))
    hint = read_bridge_root_hint()
    if hint:
        roots.append(hint)
    docs = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.DocumentsLocation)
    if docs:
        roots.append(Path(docs) / "GoB_SP_Bridge")
    win_docs = windows_documents_dir()
    if win_docs:
        roots.append(Path(win_docs) / "GoB_SP_Bridge")
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
    sp_info = data.get("substance_painter") or {}
    if not isinstance(sp_info, dict):
        return {"status": "error", "error": "Missing Substance Painter update data"}
    remote_version = str(sp_info.get("version") or "").strip()
    if not remote_version:
        return {"status": "error", "error": "Missing remote version"}
    if not is_version_newer(remote_version, PLUGIN_VERSION):
        return {
            "status": "none",
            "local_version": PLUGIN_VERSION,
            "remote_version": remote_version,
        }
    return {
        "status": "update",
        "info": {
            "version": remote_version,
            "download_url": sp_info.get("download_url"),
            "notes": data.get("notes"),
            "local_version": PLUGIN_VERSION,
        },
    }


def parse_update_data(data):
    if not isinstance(data, dict):
        return {"status": "error", "error": "Invalid update data"}
    sp_info = data.get("substance_painter") or {}
    if not isinstance(sp_info, dict):
        return {"status": "error", "error": "Missing Substance Painter update data"}
    remote_version = str(sp_info.get("version") or "").strip()
    if not remote_version:
        return {"status": "error", "error": "Missing remote version"}
    if not is_version_newer(remote_version, PLUGIN_VERSION):
        return {
            "status": "none",
            "local_version": PLUGIN_VERSION,
            "remote_version": remote_version,
        }
    return {
        "status": "update",
        "info": {
            "version": remote_version,
            "download_url": sp_info.get("download_url"),
            "notes": data.get("notes"),
            "local_version": PLUGIN_VERSION,
        },
    }


def show_update_dialog(info):
    if not info:
        return "cancel"
    version = info.get("version")
    box = QtWidgets.QMessageBox()
    box.setIcon(QtWidgets.QMessageBox.Information)
    box.setWindowTitle("GoB Bridge Update")
    box.setText(
        f"Update available: {info['version']} (current {info['local_version']})"
    )
    notes = info.get("notes")
    if notes:
        box.setInformativeText(str(notes))
    download_button = None
    if info.get("download_url"):
        download_button = box.addButton("Download", QtWidgets.QMessageBox.AcceptRole)
    dont_show_button = box.addButton("Don't show again", QtWidgets.QMessageBox.DestructiveRole)
    cancel_button = box.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
    box.exec()
    clicked = box.clickedButton()
    if download_button and clicked == download_button:
        if info.get("download_url"):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl(info["download_url"]))
        return "download"
    if clicked == dont_show_button:
        confirm = QtWidgets.QMessageBox.question(
            box,
            "GoB Bridge Update",
            "Stop showing update notifications for this version?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        if confirm == QtWidgets.QMessageBox.Yes:
            if version:
                set_update_ignore_version(version)
            return "dont_show"
        return "cancel"
    if clicked == cancel_button:
        return "cancel"
    return "cancel"


def show_update_result(result, show_no_update=False, force_prompt=False, auto_prompt=False):
    status = result.get("status") if result else None
    if status == "update":
        info = result.get("info") or {}
        version = info.get("version")
        if not force_prompt and version:
            ignored = get_update_ignore_version()
            if ignored and ignored == version:
                return
        if auto_prompt and version:
            global _update_prompted_version
            if _update_prompted_version == version:
                return
            _update_prompted_version = version
        show_update_dialog(info)
        return
    if not show_no_update:
        return
    if status == "none":
        box = QtWidgets.QMessageBox()
        box.setIcon(QtWidgets.QMessageBox.Information)
        box.setWindowTitle("GoB Bridge Update")
        box.setText(f"You're up to date ({PLUGIN_VERSION}).")
        box.exec()
        return
    error = result.get("error") if result else "Update check failed."
    box = QtWidgets.QMessageBox()
    box.setIcon(QtWidgets.QMessageBox.Warning)
    box.setWindowTitle("GoB Bridge Update")
    box.setText(str(error))
    box.exec()


_update_check_in_progress = False
_update_check_started_at = 0.0
_update_check_id = 0
_update_check_active_id = None
_update_check_timeout = 8.0
_update_check_result = None
_update_poll_timer = None
_update_check_show_no_update = False
_update_check_force_prompt = False
_update_check_auto_prompt = False
_update_net_manager = None
_update_reply = None
_update_timeout_timer = None
_update_status_kind = "idle"
_update_status_text = "Update: not checked yet"
_last_update_info = None
_update_prompted_version = None
_update_status_callbacks = set()


def _notify_update_listeners():
    for callback in list(_update_status_callbacks):
        try:
            callback()
        except Exception:
            _update_status_callbacks.discard(callback)


def _set_update_status(kind, text, info=None):
    global _update_status_kind
    global _update_status_text
    global _last_update_info
    _update_status_kind = kind
    _update_status_text = text
    if info:
        _last_update_info = info
    elif kind != "update":
        _last_update_info = None
    _notify_update_listeners()


def add_update_listener(callback):
    if callback:
        _update_status_callbacks.add(callback)


def remove_update_listener(callback):
    _update_status_callbacks.discard(callback)


def start_update_check(show_no_update=False, force_prompt=False, auto_prompt=False):
    global _update_check_in_progress
    global _update_check_started_at
    global _update_check_show_no_update
    global _update_check_force_prompt
    global _update_check_auto_prompt
    global _update_check_result
    global _update_poll_timer
    global _update_net_manager
    global _update_reply
    global _update_timeout_timer
    if _update_check_in_progress:
        if time.time() - _update_check_started_at < _update_check_timeout:
            return
        if _update_reply is not None:
            try:
                _update_reply.abort()
            except Exception:
                pass
        _update_check_in_progress = False
    _update_check_in_progress = True
    _update_check_started_at = time.time()
    _update_check_show_no_update = show_no_update
    _update_check_force_prompt = force_prompt
    _update_check_auto_prompt = auto_prompt
    _set_update_status("checking", "Update: checking...")
    _update_check_result = None

    if _update_net_manager is None:
        _update_net_manager = QtNetwork.QNetworkAccessManager()
    request = QtNetwork.QNetworkRequest(QtCore.QUrl(UPDATE_URL))
    request.setHeader(QtNetwork.QNetworkRequest.UserAgentHeader, f"GoBBridge/{PLUGIN_VERSION}")
    _update_reply = _update_net_manager.get(request)

    def _finish_update_result(result):
        global _update_check_in_progress
        global _update_check_result
        global _update_reply
        global _update_timeout_timer
        global _update_check_show_no_update
        global _update_check_force_prompt
        global _update_check_auto_prompt
        if _update_timeout_timer is not None:
            try:
                _update_timeout_timer.stop()
            except Exception:
                pass
            _update_timeout_timer = None
        if _update_reply is not None:
            try:
                _update_reply.deleteLater()
            except Exception:
                pass
            _update_reply = None
        global _update_poll_timer
        _update_check_in_progress = False
        if result.get("status") == "update":
            info = result.get("info")
            _set_update_status("update", f"Update available: {info['version']}", info=info)
        elif result.get("status") == "none":
            _set_update_status("up_to_date", f"Up to date ({PLUGIN_VERSION})")
        else:
            error = result.get("error") if result else "Update check failed."
            _set_update_status("error", f"Update check failed: {error}")
        show_update_result(
            result,
            show_no_update=_update_check_show_no_update,
            force_prompt=_update_check_force_prompt,
            auto_prompt=_update_check_auto_prompt,
        )
        _update_check_show_no_update = False
        _update_check_force_prompt = False
        _update_check_auto_prompt = False

    def _handle_reply_finished():
        global _update_check_result
        if _update_reply is None:
            return
        if _update_reply.error() != QtNetwork.QNetworkReply.NoError:
            error = _update_reply.errorString()
            _finish_update_result({"status": "error", "error": error})
            return
        try:
            raw = bytes(_update_reply.readAll())
            data = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            _finish_update_result({"status": "error", "error": str(exc)})
            return
        result = parse_update_data(data)
        _finish_update_result(result)

    _update_reply.finished.connect(_handle_reply_finished)

    if _update_timeout_timer is not None:
        try:
            _update_timeout_timer.stop()
        except Exception:
            pass
        _update_timeout_timer = None
    _update_timeout_timer = QtCore.QTimer()
    _update_timeout_timer.setSingleShot(True)

    def _on_timeout():
        if not _update_check_in_progress:
            return
        if _update_reply is not None:
            try:
                _update_reply.abort()
            except Exception:
                pass
        _finish_update_result({"status": "error", "error": "Update check timed out"})

    _update_timeout_timer.timeout.connect(_on_timeout)
    _update_timeout_timer.start(int(_update_check_timeout * 1000))


def show_message(title, text, icon=QtWidgets.QMessageBox.Information):
    box = QtWidgets.QMessageBox()
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setText(text)
    box.exec()


def show_warning_dialog(title, summary, details=None):
    box = QtWidgets.QMessageBox()
    box.setIcon(QtWidgets.QMessageBox.Information)
    box.setWindowTitle(title)
    box.setText(summary)
    box.setSizeGripEnabled(True)
    box.setWindowFlag(QtCore.Qt.WindowType.WindowMaximizeButtonHint, True)
    layout = box.layout()
    if layout:
        layout.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)
    if details:
        box.setInformativeText(
            "Suggestion: you can ignore these if the missing maps aren't needed.\n"
            "Details are available below."
        )
        box.setDetailedText(details)
        details_edit = box.findChild(QtWidgets.QTextEdit)
        if details_edit:
            details_edit.setMinimumSize(640, 320)
            details_edit.setSizePolicy(
                QtWidgets.QSizePolicy.Expanding,
                QtWidgets.QSizePolicy.Expanding,
            )
    box.setMinimumSize(520, 220)
    box.exec()


def append_log(project_dir, message, data=None):
    try:
        log_path = project_meta_dir(project_dir) / LOG_FILENAME
        ensure_dir(log_path.parent)
        with open(log_path, "a", encoding="utf-8") as handle:
            handle.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
            if data is not None:
                handle.write(json.dumps(data, indent=2, ensure_ascii=True))
                handle.write("\n")
    except OSError:
        return


def high_poly_url(path):
    return QtCore.QUrl.fromLocalFile(path).toString()


def apply_high_poly_mesh(high_path):
    if not high_path or not Path(high_path).is_file():
        return False
    try:
        import substance_painter.baking as baking
        hipoly = high_poly_url(high_path)
        texsets = list(get_all_texture_sets())
        if not texsets:
            return False
        applied = 0
        for texset in texsets:
            params = baking.BakingParameters.from_texture_set(texset)
            common = params.common()
            hipoly_prop = common.get("HipolyMesh")
            if hipoly_prop:
                baking.BakingParameters.set({hipoly_prop: hipoly})
                applied += 1
        return applied > 0
    except Exception:
        return False


def clear_high_poly_mesh():
    try:
        import substance_painter.baking as baking
        texsets = list(get_all_texture_sets())
        if not texsets:
            return False
        applied = 0
        for texset in texsets:
            params = baking.BakingParameters.from_texture_set(texset)
            common = params.common()
            hipoly_prop = common.get("HipolyMesh")
            if hipoly_prop:
                baking.BakingParameters.set({hipoly_prop: ""})
                applied += 1
        return applied > 0
    except Exception:
        return False


_high_poly_retry_timer = None
_high_poly_retry_path = None
_high_poly_retry_remaining = 0


def _stop_high_poly_retry():
    global _high_poly_retry_timer
    global _high_poly_retry_path
    global _high_poly_retry_remaining
    if _high_poly_retry_timer is not None:
        try:
            _high_poly_retry_timer.stop()
        except Exception:
            pass
    _high_poly_retry_path = None
    _high_poly_retry_remaining = 0


def _queue_high_poly_retry(high_path, retries=HIGH_POLY_RETRY_COUNT):
    global _high_poly_retry_timer
    global _high_poly_retry_path
    global _high_poly_retry_remaining
    if not high_path or retries <= 0:
        return
    _high_poly_retry_path = high_path
    _high_poly_retry_remaining = max(_high_poly_retry_remaining, retries)
    if _high_poly_retry_timer is None:
        _high_poly_retry_timer = QtCore.QTimer()
        _high_poly_retry_timer.setInterval(HIGH_POLY_RETRY_DELAY_MS)

        def _on_retry():
            global _high_poly_retry_remaining
            if not _high_poly_retry_path or _high_poly_retry_remaining <= 0:
                _stop_high_poly_retry()
                return
            if apply_high_poly_mesh(_high_poly_retry_path):
                _stop_high_poly_retry()
                return
            _high_poly_retry_remaining -= 1
            if _high_poly_retry_remaining <= 0:
                _stop_high_poly_retry()

        _high_poly_retry_timer.timeout.connect(_on_retry)
    try:
        _high_poly_retry_timer.start()
    except Exception:
        pass


def apply_high_poly_when_ready(high_path):
    if not high_path:
        return
    _queue_high_poly_retry(high_path)
    if sp.project.is_open() and sp.project.is_in_edition_state():
        if not apply_high_poly_mesh(high_path):
            _queue_high_poly_retry(high_path)
        return

    def _on_enter(_event):
        sp_event.DISPATCHER.disconnect(sp_event.ProjectEditionEntered, _on_enter)
        if not apply_high_poly_mesh(high_path):
            _queue_high_poly_retry(high_path)

    sp_event.DISPATCHER.connect(sp_event.ProjectEditionEntered, _on_enter)


def clear_high_poly_when_ready():
    _stop_high_poly_retry()
    if sp.project.is_open() and sp.project.is_in_edition_state():
        clear_high_poly_mesh()
        return

    def _on_enter(_event):
        sp_event.DISPATCHER.disconnect(sp_event.ProjectEditionEntered, _on_enter)
        clear_high_poly_mesh()

    sp_event.DISPATCHER.connect(sp_event.ProjectEditionEntered, _on_enter)


def resource_id_url(resource_id):
    try:
        return resource_id.url()
    except TypeError:
        return resource_id.url


def preset_url(preset):
    try:
        return preset.url
    except TypeError:
        return preset.url()


def collect_export_presets():
    presets = []
    seen = set()

    def add_preset(preset, kind, label_suffix=None):
        name = None
        url = None
        if hasattr(preset, "resource_id"):
            name = preset.resource_id.name
            url = resource_id_url(preset.resource_id)
        else:
            name = getattr(preset, "name", None) or getattr(preset, "label", None)
            url = getattr(preset, "url", None)
            if callable(url):
                try:
                    url = url()
                except Exception:
                    url = None
        if not name or not url:
            return
        key = name.lower()
        if key in seen:
            return
        seen.add(key)
        label = name if not label_suffix else f"{name} ({label_suffix})"
        presets.append({
            "kind": kind,
            "label": label,
            "name": name,
            "url": url,
        })

    try:
        resource_presets = sp.export.list_resource_export_presets()
    except Exception:
        resource_presets = []
    for preset in resource_presets:
        add_preset(preset, "resource")

    try:
        user_presets = sp.export.list_user_export_presets()
    except Exception:
        user_presets = []
    for preset in user_presets:
        add_preset(preset, "user", "User")

    try:
        predefined_presets = sp.export.list_predefined_export_presets()
    except Exception:
        predefined_presets = []
    for preset in predefined_presets:
        add_preset(preset, "predefined", "Predefined")

    for custom in CUSTOM_EXPORT_PRESETS:
        name = custom["name"]
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        presets.append({
            "kind": "custom",
            "label": f"{name} (Custom)",
            "name": name,
            "definition": custom,
        })

    return presets


def pick_export_preset():
    presets = collect_export_presets()
    preferred = [
        "PBR Metallic Roughness",
        "PBR Metal Rough",
        "PBR Metallic/Roughness",
    ]
    for name in preferred:
        for preset in presets:
            if preset["name"].lower() == name.lower() and preset["kind"] != "custom":
                return preset
    if presets:
        return presets[0]
    return None


def build_export_list(output_maps=None):
    export_list = []
    for texset in get_all_texture_sets():
        stacks = get_all_stacks(texset)
        if stacks:
            for stack in stacks:
                texset_name = get_sp_name(texset)
                stack_name = get_sp_name(stack)
                root = texset_name
                if stack_name:
                    root = f"{root}/{stack_name}"
                entry = {"rootPath": root}
                if output_maps:
                    entry["filter"] = {"outputMaps": output_maps}
                export_list.append(entry)
        else:
            entry = {"rootPath": get_sp_name(texset)}
            if output_maps:
                entry["filter"] = {"outputMaps": output_maps}
            export_list.append(entry)
    return export_list


def build_export_config(export_path, preset_value, output_maps=None, export_params=None, export_list=None):
    export_list = export_list or build_export_list(output_maps)
    if not export_list:
        return None
    config = {
        "exportPath": str(export_path),
        "exportShaderParams": False,
        "defaultExportPreset": preset_value,
        "exportList": export_list,
    }
    if export_params:
        config["exportParameters"] = [{
            "filter": {},
            "parameters": export_params,
        }]
    return config


def build_custom_export_config(export_path, preset_definition, output_maps=None, export_params=None, export_list=None):
    export_list = export_list or build_export_list(output_maps)
    if not export_list:
        return None
    config = {
        "exportPath": str(export_path),
        "exportShaderParams": False,
        "defaultExportPreset": preset_definition["name"],
        "exportPresets": [preset_definition],
        "exportList": export_list,
    }
    if export_params:
        config["exportParameters"] = [{
            "filter": {},
            "parameters": export_params,
        }]
    return config


def friendly_map_label(map_name):
    label = map_name
    for prefix in ("$textureSet_", "$mesh_", "$sceneMaterial_"):
        if label.startswith(prefix):
            label = label[len(prefix):]
            break
    if label.lower().startswith("textureset "):
        label = label[len("textureset "):]
    if label.lower().startswith("material "):
        label = label[len("material "):]
    cleaned = []
    depth = 0
    for ch in label:
        if ch == "(":
            depth += 1
            continue
        if ch == ")":
            if depth:
                depth -= 1
            continue
        if depth == 0:
            cleaned.append(ch)
    label = "".join(cleaned)
    label = label.replace("_", " ").replace("-", " ").strip()
    label = " ".join(label.split())
    return label or map_name


def extract_output_map_names(maps):
    output_maps = []
    if not maps:
        return output_maps
    for item in maps:
        if isinstance(item, str):
            output_maps.append(item)
            continue
        if isinstance(item, dict):
            file_name = item.get("fileName")
            if file_name:
                output_maps.append(file_name)
            continue
        file_name = getattr(item, "fileName", None) or getattr(item, "file_name", None)
        if file_name:
            output_maps.append(file_name)
    return output_maps


def get_stack_for_textureset(texset, stack_name):
    try:
        return sp.textureset.Stack.from_name(get_sp_name(texset), stack_name)
    except Exception:
        return None


def get_output_map_definitions(preset_info, stack=None):
    if not preset_info:
        return []
    if preset_info["kind"] == "custom":
        return preset_info["definition"]["maps"]
    if preset_info["kind"] in ("resource", "user"):
        try:
            resource_presets = sp.export.list_resource_export_presets()
        except Exception:
            resource_presets = []
        for preset in resource_presets:
            if preset.resource_id.name == preset_info["name"]:
                try:
                    return preset.list_output_maps()
                except Exception:
                    return []
        if preset_info["kind"] == "user":
            try:
                user_presets = sp.export.list_user_export_presets()
            except Exception:
                user_presets = []
            for preset in user_presets:
                if getattr(preset, "name", None) == preset_info["name"]:
                    try:
                        return preset.list_output_maps()
                    except Exception:
                        return []
        return []
    if preset_info["kind"] == "predefined":
        try:
            for preset in sp.export.list_predefined_export_presets():
                if preset.name == preset_info["name"]:
                    target_stack = stack
                    if not target_stack:
                        target_stack = sp.textureset.get_active_stack()
                    if not target_stack:
                        stacks = []
                        for texset in get_all_texture_sets():
                            stacks.extend(get_all_stacks(texset))
                        target_stack = stacks[0] if stacks else None
                    if not target_stack:
                        return []
                    return preset.list_output_maps(target_stack)
        except Exception:
            return []
    return []


def get_output_map_names(preset_info):
    return extract_output_map_names(get_output_map_definitions(preset_info))


def normalize_map_key(value):
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


BASECOLOR_HINTS = (
    "basecolor",
    "base_color",
    "basecol",
    "basec",
    "basecolour",
    "base_map",
    "basemap",
    "albedo",
    "diffuse",
    "color",
)
OPACITY_HINTS = ("opacity", "alpha", "transparency", "transparent", "cutout")


def map_name_is_basecolor(map_name):
    key = normalize_map_key(map_name)
    if not key:
        return False
    for hint in BASECOLOR_HINTS:
        if hint in key:
            return True
    return False


def map_def_has_opacity_channel(map_def):
    channels = map_def.get("channels") or []
    for channel in channels:
        name = str(channel.get("srcMapName") or "")
        if not name:
            continue
        key = normalize_map_key(name)
        if key in OPACITY_HINTS:
            return True
        lower = name.lower()
        if any(hint in lower for hint in OPACITY_HINTS):
            return True
    return False


def preset_basecolor_has_opacity(preset_info, stacks=None):
    map_defs = []
    if preset_info:
        stack = stacks[0] if stacks else None
        map_defs = get_output_map_definitions(preset_info, stack=stack)
    for map_def in map_defs:
        if isinstance(map_def, str):
            continue
        map_dict = map_def_to_dict(map_def)
        if not map_dict:
            continue
        file_name = map_dict.get("fileName") or map_dict.get("file_name") or ""
        if not map_name_is_basecolor(file_name):
            continue
        if map_def_has_opacity_channel(map_dict):
            return True
    return False


DOC_MAP_TO_CHANNELS = {
    "basecolor": ["BaseColor", "Diffuse", "Color"],
    "base_color": ["BaseColor", "Diffuse", "Color"],
    "basemap": ["BaseColor", "Diffuse", "Color"],
    "base map": ["BaseColor", "Diffuse", "Color"],
    "base": ["BaseColor", "Diffuse", "Color"],
    "albedo": ["BaseColor", "Diffuse", "Color"],
    "diffuse": ["Diffuse", "BaseColor", "Color"],
    "color": ["Color", "BaseColor", "Diffuse"],
    "roughness": ["Roughness"],
    "glossiness": ["Glossiness"],
    "metallic": ["Metallic"],
    "metalness": ["Metallic"],
    "normal": ["Normal"],
    "height": ["Height"],
    "displacement": ["Height"],
    "opacity": ["Opacity"],
    "alpha": ["Opacity"],
    "emissive": ["Emissive"],
    "emission": ["Emissive"],
    "specular": ["Specular", "SpecularLevel"],
    "specularlevel": ["SpecularLevel", "Specular"],
    "ao": ["AmbientOcclusion", "Occlusion", "AO"],
    "ambientocclusion": ["AmbientOcclusion", "Occlusion", "AO"],
    "occlusion": ["Occlusion", "AmbientOcclusion", "AO"],
    "user0": ["User0"],
    "user1": ["User1"],
    "user2": ["User2"],
    "user3": ["User3"],
    "blendingmask": ["BlendingMask"],
    "blending mask": ["BlendingMask"],
}


def resolve_channel_names(doc_map_name):
    if not doc_map_name:
        return []
    raw = str(doc_map_name).strip().lower()
    key = normalize_map_key(raw)
    names = []
    direct = DOC_MAP_TO_CHANNELS.get(raw)
    if direct:
        names.extend(direct)
    normalized = DOC_MAP_TO_CHANNELS.get(key)
    if normalized:
        names.extend(normalized)
    members = getattr(sp.textureset.ChannelType, "__members__", {})
    for channel_name in members:
        if normalize_map_key(channel_name) == key:
            names.append(channel_name)
    seen = set()
    unique = []
    for name in names:
        if name in seen:
            continue
        seen.add(name)
        unique.append(name)
    return unique


def stack_has_doc_map(stack, doc_map_name):
    if not doc_map_name:
        return True
    key = normalize_map_key(doc_map_name)
    if key in {"diffuse", "diffusecolor"}:
        return (
            stack_has_channel_type(stack, "Diffuse")
            or stack_has_channel_type(stack, "DiffuseColor")
        )
    lookup = resolve_channel_names(doc_map_name)
    if not lookup:
        return True
    found_type = False
    for channel_name in lookup:
        channel_type = sp.textureset.ChannelType.__members__.get(channel_name)
        if not channel_type:
            continue
        found_type = True
        if stack.has_channel(channel_type):
            return True
    return not found_type


def get_required_map_groups(map_def):
    required_doc = set()
    required_input = set()
    if not isinstance(map_def, dict):
        return required_doc, required_input
    channels = map_def.get("channels") or []
    for channel in channels:
        src_type = channel.get("srcMapType")
        name = channel.get("srcMapName")
        if not name:
            continue
        if src_type == "documentMap":
            required_doc.add(name.lower())
        elif src_type == "inputMap":
            required_input.add(name.lower())
    return required_doc, required_input


def collect_selected_stacks(selected_texture_sets=None, stack_roots=None):
    roots = stack_roots if stack_roots is not None else collect_stack_roots(selected_texture_sets)
    return [stack for _, stack in roots]


def collect_stack_roots(selected_texture_sets=None):
    roots = []
    selected_sets = None
    if selected_texture_sets:
        selected_sets = {name.lower() for name in selected_texture_sets if name}
    for texset in get_all_texture_sets():
        texset_name = get_sp_name(texset)
        if selected_sets and texset_name.lower() not in selected_sets:
            continue
        tex_stacks = get_all_stacks(texset)
        if not tex_stacks:
            stack = get_stack_for_textureset(texset, "")
            tex_stacks = [stack] if stack else []
        for stack in tex_stacks:
            if not stack:
                continue
            root = texset_name
            stack_name = get_sp_name(stack)
            if stack_name:
                root = f"{texset_name}/{stack_name}"
            roots.append((root, stack))
    return roots


def channel_display_name(doc_map_name):
    if not doc_map_name:
        return ""
    lookup = resolve_channel_names(doc_map_name)
    if lookup:
        return lookup[0]
    return str(doc_map_name)


def ensure_stack_channel(stack, doc_map_name):
    if not stack or not doc_map_name:
        return False
    if stack_has_doc_map(stack, doc_map_name):
        return True
    lookup = resolve_channel_names(doc_map_name)
    for channel_name in lookup:
        channel_type = sp.textureset.ChannelType.__members__.get(channel_name)
        if not channel_type:
            continue
        try:
            stack.add_channel(channel_type)
        except Exception:
            continue
        try:
            if stack.has_channel(channel_type):
                return True
        except Exception:
            continue
    return stack_has_doc_map(stack, doc_map_name)


def collect_missing_map_channels(
    preset_info,
    selected_output_maps,
    selected_texture_sets=None,
    stack_roots=None,
):
    missing_by_map = {}
    if not selected_output_maps:
        return missing_by_map
    selected = {str(name) for name in selected_output_maps if name}
    roots = stack_roots if stack_roots is not None else collect_stack_roots(selected_texture_sets)
    static_map_defs = None
    if preset_info and preset_info.get("kind") in ("custom", "resource", "user"):
        static_map_defs = get_output_map_definitions(preset_info)
    for root, stack in roots:
        map_defs = static_map_defs or get_output_map_definitions(preset_info, stack)
        if not map_defs:
            continue
        for map_def in map_defs:
            if isinstance(map_def, str):
                map_name = map_def
                if map_name not in selected:
                    continue
                continue
            map_dict = map_def_to_dict(map_def)
            if not map_dict:
                continue
            map_name = map_dict.get("fileName")
            if not map_name or map_name not in selected:
                continue
            required_doc, required_input = get_required_map_groups(map_dict)
            if not required_doc and not required_input:
                continue
            missing = []
            for req in sorted(required_doc):
                if not stack_has_doc_map(stack, req):
                    missing.append(req)
            for req in sorted(required_input):
                if not stack_has_doc_map(stack, req):
                    missing.append(req)
            if missing:
                missing_by_map.setdefault(map_name, {})[root] = missing
    return missing_by_map


def auto_enable_missing_channels(missing_by_map, selected_texture_sets=None, stack_roots=None):
    if not missing_by_map:
        return {}
    roots = stack_roots if stack_roots is not None else collect_stack_roots(selected_texture_sets)
    stack_lookup = {root: stack for root, stack in roots}
    enabled = {}
    for _, root_info in missing_by_map.items():
        for root, missing in root_info.items():
            stack = stack_lookup.get(root)
            if not stack:
                continue
            for name in missing:
                if ensure_stack_channel(stack, name):
                    enabled.setdefault(root, set()).add(name)
    return enabled


def channel_is_available(stack, channel):
    if not stack or not channel:
        return True
    src_type = channel.get("srcMapType")
    name = channel.get("srcMapName")
    if not name or not src_type:
        return True
    if src_type in ("documentMap", "inputMap"):
        return stack_has_doc_map(stack, name)
    return True


def channel_available_any_stack(channel, stacks):
    if not stacks:
        return True
    for stack in stacks:
        if channel_is_available(stack, channel):
            return True
    return False


def channel_available_all_stacks(channel, stacks):
    if not stacks:
        return True
    for stack in stacks:
        if not channel_is_available(stack, channel):
            return False
    return True


def stack_has_channel_type(stack, channel_name):
    if not stack or not channel_name:
        return False
    channel_type = sp.textureset.ChannelType.__members__.get(channel_name)
    if not channel_type:
        return False
    try:
        return bool(stack.has_channel(channel_type))
    except Exception:
        return False


def stacks_have_channel(stacks, channel_name):
    if not stacks:
        return False
    for stack in stacks:
        if stack_has_channel_type(stack, channel_name):
            return True
    return False


def channel_is_opacity(channel):
    if not channel:
        return False
    name = str(channel.get("srcMapName") or "")
    if not name:
        return False
    key = normalize_map_key(name)
    if key in OPACITY_HINTS:
        return True
    lower = name.lower()
    return any(hint in lower for hint in OPACITY_HINTS)


def channel_is_diffuse(channel):
    if not channel:
        return False
    name = str(channel.get("srcMapName") or "")
    if not name:
        return False
    key = normalize_map_key(name)
    return key in {"diffuse", "diffusecolor"}


def _strip_output_prefix(map_name):
    label = str(map_name or "")
    for prefix in ("$textureSet_", "$mesh_", "$sceneMaterial_"):
        if label.lower().startswith(prefix.lower()):
            label = label[len(prefix):]
            break
    return label


def _packed_doc_channels(pairs):
    channels = []
    for dest, src in pairs:
        channels.append({
            "destChannel": dest,
            "srcChannel": "L",
            "srcMapType": "documentMap",
            "srcMapName": src,
        })
    return channels


def _auto_map_type(map_name):
    label = _strip_output_prefix(map_name)
    key = normalize_map_key(label)
    if not key:
        return None
    if "occlusionroughnessmetallic" in key or "orm" in key or "arm" in key:
        return "orm"
    if "materialparams" in key or "materialparam" in key or "maskmap" in key:
        return "orm"
    if "normal" in key:
        return "normal"
    if "roughness" in key:
        return "roughness"
    if "gloss" in key or "smoothness" in key:
        return "glossiness"
    if "metallic" in key or "metalness" in key or "metal" == key:
        return "metallic"
    if "specular" in key or "reflection" in key:
        return "specular"
    if "height" in key or "displacement" in key:
        return "height"
    if "emissive" in key or "emission" in key:
        return "emission"
    if any(hint in key for hint in OPACITY_HINTS):
        return "opacity"
    if map_name_is_basecolor(label) or key in {"base", "color", "diffuse", "diffusecolor", "albedo"}:
        return "basecolor"
    if "occlusion" in key or key == "ao":
        return "ao"
    return None


def _auto_map_definition(map_name):
    map_type = _auto_map_type(map_name)
    if not map_type:
        return None
    if map_type == "orm":
        channels = _packed_doc_channels([
            ("R", "ambientocclusion"),
            ("G", "roughness"),
            ("B", "metallic"),
        ])
    elif map_type == "basecolor":
        channels = _rgb_channels("documentMap", "basecolor")
    elif map_type == "normal":
        channels = _rgb_channels("documentMap", "normal")
    elif map_type == "emission":
        channels = _rgb_channels("documentMap", "emissive")
    elif map_type == "ao":
        channels = _gray_channels("documentMap", "ambientocclusion")
    elif map_type == "height":
        channels = _gray_channels("documentMap", "height")
    elif map_type == "glossiness":
        channels = _gray_channels("documentMap", "glossiness")
    elif map_type == "specular":
        channels = _gray_channels("documentMap", "specular")
    elif map_type == "opacity":
        channels = _gray_channels("documentMap", "opacity")
    else:
        channels = _gray_channels("documentMap", map_type)
    return {
        "fileName": str(map_name),
        "channels": channels,
    }


def _map_def_has_channels(map_def):
    if isinstance(map_def, dict):
        return bool(map_def.get("channels"))
    channels = getattr(map_def, "channels", None)
    if channels is None:
        return False
    try:
        return bool(list(channels))
    except TypeError:
        return bool(channels)


def _should_force_basecolor_for_diffuse(map_names, stacks):
    if not map_names or not stacks:
        return False
    wants_diffuse = False
    for name in map_names:
        key = normalize_map_key(name)
        if "diffuse" in key:
            wants_diffuse = True
            break
    if not wants_diffuse:
        return False
    if stacks_have_channel(stacks, "Diffuse"):
        return False
    return stacks_have_channel(stacks, "BaseColor")


def channel_to_dict(channel):
    if isinstance(channel, dict):
        return dict(channel)
    dest = getattr(channel, "destChannel", None) or getattr(channel, "dest_channel", None)
    src = getattr(channel, "srcChannel", None) or getattr(channel, "src_channel", None)
    src_type = getattr(channel, "srcMapType", None) or getattr(channel, "src_map_type", None)
    src_name = getattr(channel, "srcMapName", None) or getattr(channel, "src_map_name", None)
    data = {}
    if dest is not None:
        data["destChannel"] = dest
    if src is not None:
        data["srcChannel"] = src
    if src_type is not None:
        data["srcMapType"] = src_type
    if src_name is not None:
        data["srcMapName"] = src_name
    return data or None


def map_def_to_dict(map_def):
    if isinstance(map_def, dict):
        result = dict(map_def)
        channels = result.get("channels")
        if isinstance(channels, list):
            converted = []
            for channel in channels:
                channel_dict = channel_to_dict(channel)
                if channel_dict:
                    converted.append(channel_dict)
            if channels and not converted:
                return None
            if converted:
                result["channels"] = converted
        return result
    file_name = getattr(map_def, "fileName", None) or getattr(map_def, "file_name", None)
    if not file_name:
        return None
    result = {"fileName": file_name}
    channels = getattr(map_def, "channels", None)
    if channels is not None:
        converted = []
        for channel in channels:
            channel_dict = channel_to_dict(channel)
            if channel_dict:
                converted.append(channel_dict)
        if channels and not converted:
            return None
        if converted:
            result["channels"] = converted
    parameters = getattr(map_def, "parameters", None) or getattr(map_def, "params", None)
    if parameters is not None:
        result["parameters"] = parameters
    return result


def sanitize_map_definitions(preset_info, selected_texture_sets=None, stacks=None):
    map_defs = get_output_map_definitions(preset_info)
    if not map_defs:
        return None, False, {}
    stacks = stacks if stacks is not None else collect_selected_stacks(selected_texture_sets)
    map_names = extract_output_map_names(map_defs)
    if map_defs and not any(_map_def_has_channels(item) for item in map_defs):
        if _should_force_basecolor_for_diffuse(map_names, stacks):
            auto_defs = []
            for name in map_names:
                auto_def = _auto_map_definition(name)
                if not auto_def:
                    auto_defs = []
                    break
                auto_defs.append(auto_def)
            if auto_defs:
                return auto_defs, True, {}
    sanitized = []
    changed = False
    removed = {}
    basecolor_present = stacks_have_channel(stacks, "BaseColor")
    if not basecolor_present:
        basecolor_present = any(stack_has_doc_map(stack, "basecolor") for stack in stacks)
    can_sub_basecolor = basecolor_present and not stacks_have_channel(stacks, "Diffuse")
    for map_def in map_defs:
        if isinstance(map_def, str):
            auto_def = _auto_map_definition(map_def)
            if auto_def:
                sanitized.append(auto_def)
                changed = True
            else:
                sanitized.append(map_def)
            continue
        map_dict = map_def_to_dict(map_def)
        if not map_dict:
            return None, False, {}
        map_name = map_dict.get("fileName")
        channels = map_dict.get("channels")
        if not channels:
            auto_def = _auto_map_definition(map_name)
            if auto_def:
                sanitized.append(auto_def)
                changed = True
                continue
        if channels:
            kept = []
            removed_channels = []
            for channel in channels:
                if can_sub_basecolor and channel_is_diffuse(channel):
                    channel = dict(channel)
                    channel["srcMapName"] = "basecolor"
                    changed = True
                require_all = False
                if map_name and map_name_is_basecolor(map_name) and channel_is_opacity(channel):
                    require_all = True
                if (channel_available_all_stacks(channel, stacks) if require_all
                        else channel_available_any_stack(channel, stacks)):
                    kept.append(channel)
                else:
                    name = channel.get("srcMapName")
                    if name:
                        removed_channels.append(str(name).lower())
                    changed = True
            if not kept:
                changed = True
                if removed_channels:
                    if map_name:
                        removed.setdefault(map_name, set()).update(removed_channels)
                continue
            if len(kept) != len(channels):
                map_dict["channels"] = kept
                if removed_channels:
                    if map_name:
                        removed.setdefault(map_name, set()).update(removed_channels)
        sanitized.append(map_dict)
    return sanitized, changed, removed


def infer_normal_map_format_from_preset(preset_info):
    if not preset_info:
        return None
    name = str(preset_info.get("name") or "").lower()
    if "directx" in name or "d3d" in name:
        return "directx"
    if "opengl" in name or "ogl" in name:
        return "opengl"
    map_defs = get_output_map_definitions(preset_info)
    for map_def in map_defs:
        if not isinstance(map_def, dict):
            continue
        channels = map_def.get("channels") or []
        for channel in channels:
            if channel.get("srcMapType") != "virtualMap":
                continue
            src_name = str(channel.get("srcMapName") or "").lower()
            if "directx" in src_name or "d3d" in src_name:
                return "directx"
            if "opengl" in src_name or "ogl" in src_name:
                return "opengl"
    return None


def build_export_list_for_preset(
    preset_info,
    selected_output_maps,
    selected_texture_sets=None,
    stack_roots=None,
):
    export_list = []
    selected_map_set = {name for name in selected_output_maps if name}
    if stack_roots is None:
        selected_sets = None
        if selected_texture_sets:
            selected_sets = {name.lower() for name in selected_texture_sets if name}
        stack_roots = []
        for texset in get_all_texture_sets():
            texset_name = get_sp_name(texset)
            if selected_sets and texset_name.lower() not in selected_sets:
                continue
            stacks = get_all_stacks(texset)
            if not stacks:
                stack = get_stack_for_textureset(texset, "")
                stacks = [stack] if stack else []
            for stack in stacks:
                if not stack:
                    continue
                root = texset_name
                stack_name = get_sp_name(stack)
                if stack_name:
                    root = f"{root}/{stack_name}"
                stack_roots.append((root, stack))
    static_map_defs = None
    if preset_info and preset_info.get("kind") in ("custom", "resource", "user"):
        static_map_defs = get_output_map_definitions(preset_info)
    for root, stack in stack_roots:
        if not stack:
            continue
        map_defs = static_map_defs or get_output_map_definitions(preset_info, stack)
        if not map_defs:
            continue
        valid_maps = []
        available_maps = []
        for map_def in map_defs:
            if isinstance(map_def, dict):
                map_name = map_def.get("fileName")
                if map_name:
                    available_maps.append(map_name)
                if not map_name or map_name not in selected_map_set:
                    continue
                required_doc, required_input = get_required_map_groups(map_def)
                missing_doc = [
                    req for req in required_doc if not stack_has_doc_map(stack, req)
                ]
                missing_input = [
                    req for req in required_input if not stack_has_doc_map(stack, req)
                ]
                if missing_doc or missing_input:
                    if map_name_is_basecolor(map_name):
                        valid_maps.append(map_name)
                    continue
                valid_maps.append(map_name)
            elif isinstance(map_def, str):
                available_maps.append(map_def)
                if map_def in selected_map_set:
                    valid_maps.append(map_def)
            else:
                map_name = getattr(map_def, "fileName", None) or getattr(map_def, "file_name", None)
                if map_name:
                    available_maps.append(map_name)
                if map_name and map_name in selected_map_set:
                    valid_maps.append(map_name)
        if not valid_maps and selected_output_maps:
            if available_maps:
                valid_maps = [name for name in selected_output_maps if name in available_maps]
            if not valid_maps:
                valid_maps = list(selected_output_maps)
        if not valid_maps:
            continue
        export_list.append({
            "rootPath": root,
            "filter": {"outputMaps": valid_maps},
        })
    return export_list


def build_export_parameters(settings):
    merged = dict(DEFAULT_EXPORT_SETTINGS)
    if settings:
        merged.update(settings)
    valid_formats = {value for value, _ in EXPORT_FORMATS}
    valid_bit_depths = {value for value, _ in EXPORT_BIT_DEPTHS}
    valid_resolutions = {value for value, _ in EXPORT_RESOLUTIONS}
    valid_padding = {value for value, _ in PADDING_ALGORITHMS}
    file_format = merged.get("file_format")
    if file_format not in valid_formats:
        file_format = DEFAULT_EXPORT_SETTINGS["file_format"]
    bit_depth = merged.get("bit_depth")
    if bit_depth not in valid_bit_depths:
        bit_depth = DEFAULT_EXPORT_SETTINGS["bit_depth"]
    size_log2 = merged.get("size_log2")
    if size_log2 not in valid_resolutions:
        size_log2 = DEFAULT_EXPORT_SETTINGS["size_log2"]
    padding = merged.get("padding_algorithm")
    if padding not in valid_padding:
        padding = DEFAULT_EXPORT_SETTINGS["padding_algorithm"]
    try:
        dilation = int(merged.get("dilation_distance"))
    except (TypeError, ValueError):
        dilation = DEFAULT_EXPORT_SETTINGS["dilation_distance"]
    dithering = bool(merged.get("dithering"))
    return {
        "fileFormat": file_format,
        "bitDepth": bit_depth,
        "dithering": dithering,
        "sizeLog2": size_log2,
        "paddingAlgorithm": padding,
        "dilationDistance": dilation,
    }


def mesh_option_key(option):
    try:
        return option.name
    except AttributeError:
        return str(option)


def resolve_enum_member(enum_obj, key_hint):
    if not enum_obj or not key_hint:
        return None
    members = getattr(enum_obj, "__members__", None)
    if members:
        for key, value in members.items():
            if key.lower() == key_hint.lower():
                return value
        for key, value in members.items():
            if key_hint.lower() in key.lower():
                return value
    if hasattr(enum_obj, key_hint):
        return getattr(enum_obj, key_hint)
    for key in dir(enum_obj):
        if key.lower() == key_hint.lower():
            return getattr(enum_obj, key)
    return None


def resolve_enum_by_hints(enum_obj, hints):
    for hint in hints:
        value = resolve_enum_member(enum_obj, hint)
        if value is not None:
            return value
    return None


def resolve_texture_size(size_value):
    if size_value is None:
        return None
    try:
        size = int(size_value)
    except (TypeError, ValueError):
        return None
    enum_cls = getattr(sp.project, "TextureSize", None)
    if enum_cls:
        value = resolve_enum_by_hints(enum_cls, [str(size)])
        if value is not None:
            return value
        try:
            return enum_cls(size, size)
        except Exception:
            pass
    return size


def try_set_attr(target, names, value):
    for name in names:
        if hasattr(target, name):
            try:
                setattr(target, name, value)
                return True
            except Exception:
                continue
    return False


def try_set_attr_contains(target, token, value):
    token = token.lower()
    for name in dir(target):
        if token not in name.lower():
            continue
        try:
            setattr(target, name, value)
            return True
        except Exception:
            continue
    return False


def set_attr_if_present(target, name, value):
    if not hasattr(target, name):
        return False
    try:
        current = getattr(target, name)
        if callable(current):
            return False
    except Exception:
        current = None
    try:
        setattr(target, name, value)
        return True
    except Exception:
        return False


def set_auto_unwrap_flags(target):
    if not target:
        return False
    handled = False
    candidates = [
        "auto_unwrap",
        "auto_unwrap_enabled",
        "auto_unwrap_uvs",
        "auto_unwrap_uv",
        "unwrap",
        "unwrap_uvs",
        "unwrap_uv",
        "generate_uvs",
        "generate_uv",
        "create_uvs",
        "create_uv",
        "auto_uv",
        "auto_uvs",
        "force_unwrap",
        "force_auto_unwrap",
    ]
    for name in candidates:
        if set_attr_if_present(target, name, True):
            handled = True
    if not handled:
        try_set_attr_contains(target, "unwrap", True)
        try_set_attr_contains(target, "uv", True)
        handled = True
    return handled


def build_project_settings(settings_dict):
    if not settings_dict or not isinstance(settings_dict, dict):
        return None
    settings_cls = getattr(sp.project, "ProjectSettings", None)
    if not settings_cls:
        return None
    try:
        settings = settings_cls()
    except Exception:
        return None

    size_value = resolve_texture_size(settings_dict.get("document_resolution"))
    if size_value is not None:
        try_set_attr(
            settings,
            [
                "texture_size",
                "texture_set_size",
                "default_texture_set_size",
                "default_texture_size",
                "texture_resolution",
            ],
            size_value,
        )

    normal_format = settings_dict.get("normal_map_format")
    if normal_format:
        enum_cls = getattr(sp.project, "NormalMapFormat", None)
        normal_val = resolve_enum_member(enum_cls, str(normal_format))
        if normal_val is None:
            normal_val = resolve_enum_member(enum_cls, str(normal_format).title())
        if normal_val is None:
            normal_val = normal_format
        try_set_attr(
            settings,
            ["normal_map_format", "normalMapFormat", "normal_format"],
            normal_val,
        )

    tangent = settings_dict.get("tangent_space_per_fragment")
    if tangent is not None:
        try_set_attr(
            settings,
            ["tangent_space_per_fragment", "compute_tangent_space_per_fragment"],
            bool(tangent),
        )

    use_uv_tiles = settings_dict.get("use_uv_tiles")
    if use_uv_tiles is not None:
        use_uv_tiles = bool(use_uv_tiles)
        try_set_attr(settings, ["use_uv_tiles", "use_uv_tile_workflow"], use_uv_tiles)
        enum_cls = getattr(sp.project, "UVTileWorkflow", None) or getattr(sp.project, "UvTileWorkflow", None)
        if enum_cls:
            if use_uv_tiles:
                uv_val = resolve_enum_by_hints(enum_cls, ["UDIM", "UV", "Tile"])
            else:
                uv_val = resolve_enum_by_hints(enum_cls, ["None", "Single", "Disabled", "Off"])
            if uv_val is not None:
                try_set_attr(settings, ["uv_tile_workflow", "uv_tiles_workflow"], uv_val)

    import_cameras = settings_dict.get("import_cameras")
    if import_cameras is not None:
        try_set_attr(settings, ["import_cameras", "import_camera"], bool(import_cameras))

    return settings


def build_import_settings(auto_unwrap=False):
    if not auto_unwrap:
        return None
    candidates = [
        "MeshImportSettings",
        "ImportSettings",
        "ProjectImportSettings",
    ]
    settings_obj = None
    for name in candidates:
        cls = getattr(sp.project, name, None)
        if cls is None:
            continue
        try:
            settings_obj = cls()
            break
        except Exception:
            continue
    if not settings_obj:
        return None
    set_auto_unwrap_flags(settings_obj)
    try_set_attr(settings_obj, ["compute_tangent_space_per_fragment"], True)
    return settings_obj


def build_reload_settings(auto_unwrap=False):
    try:
        settings = sp.project.MeshReloadingSettings(
            import_cameras=False,
            preserve_strokes=True,
        )
    except Exception:
        return None
    if auto_unwrap:
        try_set_attr(settings, ["compute_tangent_space_per_fragment"], True)
        handled = try_set_attr(
            settings,
            [
                "auto_unwrap",
                "auto_unwrap_enabled",
                "auto_unwrap_uvs",
                "unwrap",
            ],
            True,
        )
        if not handled:
            import_settings = getattr(settings, "import_settings", None) or getattr(settings, "importSettings", None)
            if import_settings:
                try_set_attr(import_settings, ["compute_tangent_space_per_fragment"], True)
                handled = try_set_attr(
                    import_settings,
                    [
                        "auto_unwrap",
                        "auto_unwrap_enabled",
                        "auto_unwrap_uvs",
                        "unwrap",
                    ],
                    True,
                )
        if not handled:
            try_set_attr_contains(settings, "unwrap", True)
    return settings


def ensure_uv_channel():
    try:
        stack = sp.textureset.get_active_stack()
    except Exception:
        stack = None
    if not stack:
        try:
            texsets = get_all_texture_sets()
            if texsets:
                stacks = get_all_stacks(texsets[0])
                stack = stacks[0] if stacks else None
        except Exception:
            stack = None
    if not stack:
        return
    try:
        uv_channel = sp.textureset.ChannelType.__members__.get("UV")
    except Exception:
        uv_channel = None
    if not uv_channel:
        return
    try:
        if not stack.has_channel(uv_channel):
            stack.add_channel(uv_channel)
    except Exception:
        return


class _NeutralCheckedItemDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        opt = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        if opt.state & QtWidgets.QStyle.State_On:
            opt.backgroundBrush = QtGui.QBrush(QtCore.Qt.transparent)
        super().paint(painter, opt, index)


class _DragCheckListWidget(QtWidgets.QListWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._drag_check_active = False
        self._drag_check_state = None
        self._drag_last_row = None
        self._drag_start_row = None
        self._drag_start_pos = None
        self._drag_initial_state = None

    def _reset_drag_state(self):
        self._drag_check_active = False
        self._drag_check_state = None
        self._drag_last_row = None
        self._drag_start_row = None
        self._drag_start_pos = None
        self._drag_initial_state = None

    def _apply_drag_range(self, start_row, end_row, state):
        if start_row is None or end_row is None:
            return
        if end_row < start_row:
            start_row, end_row = end_row, start_row
        for row in range(start_row, end_row + 1):
            item = self.item(row)
            if not item:
                continue
            if item.flags() & QtCore.Qt.ItemIsUserCheckable:
                item.setCheckState(state)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            item = self.itemAt(event.pos())
            if item and (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                self._drag_start_row = self.row(item)
                self._drag_start_pos = event.pos()
                self._drag_initial_state = item.checkState()
            else:
                self._reset_drag_state()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (event.buttons() & QtCore.Qt.LeftButton) and self._drag_start_row is not None:
            if not self._drag_check_active:
                if self._drag_start_pos is None:
                    self._drag_start_pos = event.pos()
                distance = (event.pos() - self._drag_start_pos).manhattanLength()
                if distance < QtWidgets.QApplication.startDragDistance():
                    super().mouseMoveEvent(event)
                    return
                self._drag_check_active = True
                self._drag_check_state = (
                    QtCore.Qt.Unchecked
                    if self._drag_initial_state == QtCore.Qt.Checked
                    else QtCore.Qt.Checked
                )
                self._drag_last_row = self._drag_start_row
                start_item = self.item(self._drag_start_row)
                if start_item and (start_item.flags() & QtCore.Qt.ItemIsUserCheckable):
                    start_item.setCheckState(self._drag_check_state)
            if self._drag_check_active:
                item = self.itemAt(event.pos())
                if item and (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                    row = self.row(item)
                    last_row = self._drag_last_row
                    if last_row is None:
                        last_row = row
                    if row != last_row:
                        self._apply_drag_range(last_row, row, self._drag_check_state)
                        self._drag_last_row = row
                    else:
                        item.setCheckState(self._drag_check_state)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            was_drag = self._drag_check_active
            self._reset_drag_state()
            if was_drag:
                event.accept()
                return
        super().mouseReleaseEvent(event)


class ExportDialog(QtWidgets.QDialog):
    def __init__(self):
        super().__init__(QtWidgets.QApplication.activeWindow())
        self.setWindowTitle("GoB Bridge - Send to Blender")
        self._presets = []
        self._all_presets = []
        self._user_presets = []
        self._pending_map_selection = None
        self._pending_texture_sets = None
        self._default_preset_options = None
        self._loading = True

        self._project_dir = get_project_dir()
        self._linked_blender_file = read_linked_blender_file(self._project_dir)

        state = load_persistent_state(self._project_dir)
        self._last_state = state.get("last_settings", {})
        self._user_presets = [
            preset for preset in state.get("user_presets", [])
            if (isinstance(preset, dict) and preset.get("name") and
                preset.get("name").lower() != DEFAULT_USER_PRESET_NAME.lower())
        ]
        self._pending_map_selection = self._last_state.get("output_maps")
        self._pending_texture_sets = self._last_state.get("texture_sets")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        content = QtWidgets.QWidget()
        content_layout = QtWidgets.QVBoxLayout(content)
        content_layout.setSpacing(10)
        content_layout.setContentsMargins(0, 0, 0, 0)

        header = QtWidgets.QLabel("GoB Bridge Export")
        header.setStyleSheet("font-weight: 600; font-size: 14px;")
        content_layout.addWidget(header)

        update_bar = QtWidgets.QHBoxLayout()
        self.update_status_label = QtWidgets.QLabel()
        self.update_status_label.setText(_update_status_text)
        self.update_status_label.setStyleSheet("font-weight: 600;")
        self.update_check_btn = QtWidgets.QPushButton("Check Updates")
        self.update_download_btn = QtWidgets.QPushButton("Download")
        self.update_download_btn.setEnabled(False)
        self.report_bug_btn = QtWidgets.QPushButton("Report Bug")
        update_bar.addWidget(self.update_status_label)
        update_bar.addStretch()
        update_bar.addWidget(self.update_check_btn)
        update_bar.addWidget(self.update_download_btn)
        update_bar.addWidget(self.report_bug_btn)
        update_widget = QtWidgets.QWidget()
        update_widget.setLayout(update_bar)
        content_layout.addWidget(update_widget)

        preset_bar = QtWidgets.QHBoxLayout()
        preset_bar.addWidget(QtWidgets.QLabel("Bridge preset"))
        self.user_preset_combo = QtWidgets.QComboBox()
        preset_bar.addWidget(self.user_preset_combo, 1)
        self.save_preset_btn = QtWidgets.QPushButton("Save")
        self.delete_preset_btn = QtWidgets.QPushButton("Delete")
        preset_bar.addWidget(self.save_preset_btn)
        preset_bar.addWidget(self.delete_preset_btn)
        preset_widget = QtWidgets.QWidget()
        preset_widget.setLayout(preset_bar)
        content_layout.addWidget(preset_widget)

        self.link_group = QtWidgets.QGroupBox("Linked Blender Project")
        link_layout = QtWidgets.QVBoxLayout(self.link_group)
        link_layout.setContentsMargins(8, 6, 8, 6)
        link_layout.setSpacing(4)
        self.detected_blender_label = QtWidgets.QLabel()
        self.detected_blender_label.setWordWrap(True)
        self.detected_blender_label.setStyleSheet("color: #666;")
        link_layout.addWidget(self.detected_blender_label)
        self.open_blender_cb = QtWidgets.QCheckBox("Open linked project on export")
        link_layout.addWidget(self.open_blender_cb)
        self.open_temp_blender_cb = QtWidgets.QCheckBox("Allow opening temp linked project")
        self.open_temp_blender_cb.setToolTip(
            "Enable opening a linked unsaved Blender file from the bridge temp folder."
        )
        link_layout.addWidget(self.open_temp_blender_cb)
        self.open_temp_blender_cb.toggled.connect(self._refresh_linked_blender_state)
        content_layout.addWidget(self.link_group)

        self.mesh_group = QtWidgets.QGroupBox("Mesh Export")
        mesh_layout = QtWidgets.QVBoxLayout(self.mesh_group)
        self.mesh_cb = QtWidgets.QCheckBox("Export mesh (FBX)")
        self.mesh_cb.setChecked(not (self._project_dir / SP_EXPORT_FILENAME).exists())
        mesh_layout.addWidget(self.mesh_cb)
        mesh_form = QtWidgets.QFormLayout()
        self.mesh_combo = QtWidgets.QComboBox()
        self.mesh_combo.addItem("Base Mesh", sp.export.MeshExportOption.BaseMesh)
        self.mesh_combo.addItem("Triangulated Mesh", sp.export.MeshExportOption.TriangulatedMesh)
        self.mesh_combo.addItem(
            "Tessellation Normals Base Mesh",
            sp.export.MeshExportOption.TessellationNormalsBaseMesh,
        )
        mesh_form.addRow("Mesh option", self.mesh_combo)
        self.mesh_options_widget = QtWidgets.QWidget()
        self.mesh_options_widget.setLayout(mesh_form)
        mesh_layout.addWidget(self.mesh_options_widget)
        content_layout.addWidget(self.mesh_group)

        self.texture_group = QtWidgets.QGroupBox("Texture Export")
        texture_layout = QtWidgets.QVBoxLayout(self.texture_group)
        self.textures_cb = QtWidgets.QCheckBox("Export textures")
        self.textures_cb.setChecked(False)
        texture_layout.addWidget(self.textures_cb)

        self.texture_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        texture_layout.addWidget(self.texture_splitter, 1)

        self.texture_sets_group = QtWidgets.QGroupBox("Texture Sets")
        texset_layout = QtWidgets.QVBoxLayout(self.texture_sets_group)
        texset_layout.setContentsMargins(8, 6, 8, 6)
        texset_layout.setSpacing(4)
        self.texture_sets_group.setSizePolicy(
            QtWidgets.QSizePolicy.Preferred,
            QtWidgets.QSizePolicy.Fixed,
        )
        texset_toolbar = QtWidgets.QHBoxLayout()
        self.texset_status = QtWidgets.QLabel("Selected: 0/0")
        texset_toolbar.addWidget(self.texset_status)
        texset_toolbar.addStretch()
        self.texset_all_btn = QtWidgets.QPushButton("Include All")
        self.texset_none_btn = QtWidgets.QPushButton("None")
        self.texset_refresh_btn = QtWidgets.QPushButton("Refresh")
        texset_toolbar.addWidget(self.texset_all_btn)
        texset_toolbar.addWidget(self.texset_none_btn)
        texset_toolbar.addWidget(self.texset_refresh_btn)
        texset_layout.addLayout(texset_toolbar)
        self.texset_list = _DragCheckListWidget()
        self.texset_list.setObjectName("gob_texset_list")
        self.texset_list.setMinimumWidth(280)
        self.texset_list.setMinimumHeight(120)
        self.texset_list.setFixedHeight(120)
        self.texset_list.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding,
            QtWidgets.QSizePolicy.Fixed,
        )
        self.texset_list.setStyleSheet(
            "#gob_texset_list::item:checked { background: transparent; }"
            "#gob_texset_list::item:selected { background: rgba(255, 255, 255, 0.08); }"
            "#gob_texset_list::item { color: palette(text); }"
        )
        self.texset_list.setAlternatingRowColors(False)
        self.texset_list.setUniformItemSizes(True)
        self._texset_delegate = _NeutralCheckedItemDelegate(self.texset_list)
        self.texset_list.setItemDelegate(self._texset_delegate)
        self.texset_list.itemChanged.connect(self._update_texture_set_count)
        texset_layout.addWidget(self.texset_list)

        self.map_group = QtWidgets.QGroupBox("List of Exports")
        maps_layout = QtWidgets.QVBoxLayout(self.map_group)
        maps_layout.setContentsMargins(8, 6, 8, 6)
        maps_layout.setSpacing(4)
        self.map_group.setSizePolicy(
            QtWidgets.QSizePolicy.Preferred,
            QtWidgets.QSizePolicy.Fixed,
        )
        map_toolbar = QtWidgets.QHBoxLayout()
        self.map_status = QtWidgets.QLabel("Selected: 0/0")
        map_toolbar.addWidget(self.map_status)
        map_toolbar.addStretch()
        self.select_all_btn = QtWidgets.QPushButton("Include All")
        self.select_none_btn = QtWidgets.QPushButton("None")
        self.refresh_maps_btn = QtWidgets.QPushButton("Refresh")
        map_toolbar.addWidget(self.select_all_btn)
        map_toolbar.addWidget(self.select_none_btn)
        map_toolbar.addWidget(self.refresh_maps_btn)
        maps_layout.addLayout(map_toolbar)
        self.map_list = _DragCheckListWidget()
        self.map_list.setObjectName("gob_map_list")
        self.map_list.setMinimumHeight(160)
        self.map_list.setFixedHeight(160)
        self.map_list.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding,
            QtWidgets.QSizePolicy.Fixed,
        )
        self.map_list.setStyleSheet(
            "#gob_map_list::item:checked { background: transparent; }"
            "#gob_map_list::item:selected { background: rgba(255, 255, 255, 0.08); }"
            "#gob_map_list::item { color: palette(text); }"
        )
        self.map_list.setAlternatingRowColors(False)
        self.map_list.setUniformItemSizes(True)
        self._map_delegate = _NeutralCheckedItemDelegate(self.map_list)
        self.map_list.setItemDelegate(self._map_delegate)
        self.map_list.itemChanged.connect(self._update_map_count)
        maps_layout.addWidget(self.map_list)

        left_panel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(2)
        left_layout.addWidget(self.texture_sets_group)
        left_layout.addWidget(self.map_group)
        left_layout.addStretch(1)

        self.texture_params_group = QtWidgets.QGroupBox("General Export Parameters")
        params_form = QtWidgets.QFormLayout(self.texture_params_group)
        self.output_dir_edit = QtWidgets.QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setText(str(self._project_dir / "textures"))
        self.open_output_btn = QtWidgets.QPushButton("Open")
        output_row = QtWidgets.QHBoxLayout()
        output_row.addWidget(self.output_dir_edit)
        output_row.addWidget(self.open_output_btn)
        output_widget = QtWidgets.QWidget()
        output_widget.setLayout(output_row)
        params_form.addRow("Output directory", output_widget)

        self.preset_search = QtWidgets.QLineEdit()
        self.preset_search.setPlaceholderText("Search presets...")
        self.preset_search.setClearButtonEnabled(True)
        self.preset_combo = QtWidgets.QComboBox()
        preset_box = QtWidgets.QWidget()
        preset_box_layout = QtWidgets.QVBoxLayout(preset_box)
        preset_box_layout.setContentsMargins(0, 0, 0, 0)
        preset_box_layout.addWidget(self.preset_search)
        preset_box_layout.addWidget(self.preset_combo)
        params_form.addRow("Output template", preset_box)

        self.format_combo = QtWidgets.QComboBox()
        for value, label in EXPORT_FORMATS:
            self.format_combo.addItem(label, value)
        fmt_index = self.format_combo.findData(DEFAULT_EXPORT_SETTINGS["file_format"])
        if fmt_index >= 0:
            self.format_combo.setCurrentIndex(fmt_index)

        self.bitdepth_combo = QtWidgets.QComboBox()
        for value, label in EXPORT_BIT_DEPTHS:
            self.bitdepth_combo.addItem(label, value)
        depth_index = self.bitdepth_combo.findData(DEFAULT_EXPORT_SETTINGS["bit_depth"])
        if depth_index >= 0:
            self.bitdepth_combo.setCurrentIndex(depth_index)

        filetype_layout = QtWidgets.QHBoxLayout()
        filetype_layout.addWidget(self.format_combo)
        filetype_layout.addWidget(self.bitdepth_combo)
        filetype_widget = QtWidgets.QWidget()
        filetype_widget.setLayout(filetype_layout)
        params_form.addRow("File type", filetype_widget)

        self.res_combo = QtWidgets.QComboBox()
        for value, label in EXPORT_RESOLUTIONS:
            self.res_combo.addItem(label, value)
        res_index = self.res_combo.findData(DEFAULT_EXPORT_SETTINGS["size_log2"])
        if res_index >= 0:
            self.res_combo.setCurrentIndex(res_index)
        params_form.addRow("Size", self.res_combo)

        self.padding_combo = QtWidgets.QComboBox()
        for value, label in PADDING_ALGORITHMS:
            self.padding_combo.addItem(label, value)
        pad_index = self.padding_combo.findData(DEFAULT_EXPORT_SETTINGS["padding_algorithm"])
        if pad_index >= 0:
            self.padding_combo.setCurrentIndex(pad_index)

        self.dilation_spin = QtWidgets.QSpinBox()
        self.dilation_spin.setRange(0, 256)
        self.dilation_spin.setValue(DEFAULT_EXPORT_SETTINGS["dilation_distance"])

        padding_layout = QtWidgets.QHBoxLayout()
        padding_layout.addWidget(self.padding_combo)
        padding_layout.addWidget(self.dilation_spin)
        padding_widget = QtWidgets.QWidget()
        padding_widget.setLayout(padding_layout)
        params_form.addRow("Padding", padding_widget)

        self.dither_cb = QtWidgets.QCheckBox("Dithering")
        self.dither_cb.setChecked(DEFAULT_EXPORT_SETTINGS["dithering"])
        params_form.addRow(self.dither_cb)

        right_panel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.addWidget(self.texture_params_group)

        self.texture_splitter.addWidget(left_panel)
        self.texture_splitter.addWidget(right_panel)
        self.texture_splitter.setStretchFactor(0, 3)
        self.texture_splitter.setStretchFactor(1, 2)
        self.texture_splitter.setSizes([380, 520])

        content_layout.addWidget(self.texture_group)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setWidget(content)
        layout.addWidget(scroll, 1)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.textures_cb.toggled.connect(self._on_textures_toggle)
        self.mesh_cb.toggled.connect(self._on_mesh_toggle)
        self.preset_search.textChanged.connect(self._filter_presets)
        self.preset_combo.currentIndexChanged.connect(self._refresh_map_list)
        self.select_all_btn.clicked.connect(lambda: self._set_all_map_checks(True))
        self.select_none_btn.clicked.connect(lambda: self._set_all_map_checks(False))
        self.refresh_maps_btn.clicked.connect(self._refresh_map_list)
        self.texset_all_btn.clicked.connect(lambda: self._set_all_texture_set_checks(True))
        self.texset_none_btn.clicked.connect(lambda: self._set_all_texture_set_checks(False))
        self.texset_refresh_btn.clicked.connect(self._populate_texture_sets)
        self.open_output_btn.clicked.connect(self._open_output_dir)
        self.save_preset_btn.clicked.connect(self._save_user_preset)
        self.delete_preset_btn.clicked.connect(self._delete_user_preset)
        self.user_preset_combo.currentIndexChanged.connect(self._apply_user_preset_selection)
        self.update_check_btn.clicked.connect(lambda: start_update_check(show_no_update=True, force_prompt=True))
        self.update_download_btn.clicked.connect(self._open_update_download)
        self.report_bug_btn.clicked.connect(self._open_bug_report)
        add_update_listener(self._refresh_update_status)

        self._populate_texture_sets()
        self._all_presets = collect_export_presets()
        self._filter_presets("")
        if not self._last_state or not self._last_state.get("preset"):
            default_preset = self._find_default_export_preset()
            if default_preset:
                self._select_preset_by_ref({
                    "kind": default_preset.get("kind"),
                    "name": default_preset.get("name"),
                })
        self._default_preset_options = self._build_default_preset_options()
        self._reload_user_presets()
        self._apply_saved_state(self._last_state)
        self._on_textures_toggle(self.textures_cb.isChecked())
        self._on_mesh_toggle(self.mesh_cb.isChecked())
        self._loading = False
        self._refresh_update_status()
        self._refresh_linked_blender_state()
        self._apply_initial_size()
        self._center_on_screen()

    def closeEvent(self, event):
        remove_update_listener(self._refresh_update_status)
        try:
            options = self.get_options()
            state = self._serialize_options(options)
            save_persistent_state(
                last_settings=state,
                user_presets=self._user_presets,
                project_dir=self._project_dir,
            )
        except Exception:
            pass
        super().closeEvent(event)

    def _refresh_update_status(self):
        self.update_status_label.setText(_update_status_text)
        info = _last_update_info if _update_status_kind == "update" else None
        enabled = bool(info and info.get("download_url"))
        self.update_download_btn.setEnabled(enabled)
        if enabled:
            self.update_download_btn.setToolTip(info.get("download_url"))
        else:
            self.update_download_btn.setToolTip("")

    def _center_on_screen(self):
        screen = self.screen()
        if screen is None:
            screen = QtGui.QGuiApplication.primaryScreen()
        if screen is None:
            return
        rect = screen.availableGeometry()
        frame = self.frameGeometry()
        frame.moveCenter(rect.center())
        self.move(frame.topLeft())

    def _open_update_download(self):
        if _last_update_info and _last_update_info.get("download_url"):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl(_last_update_info["download_url"]))

    def _open_bug_report(self):
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(BUG_REPORT_URL))

    def _apply_initial_size(self):
        screen = self.screen() or QtGui.QGuiApplication.primaryScreen()
        width = 900
        height = 700
        if screen:
            rect = screen.availableGeometry()
            width = min(width, int(rect.width() * 0.95))
            height = min(height, int(rect.height() * 0.85))
        self.setMinimumSize(860, 620)
        self.resize(width, height)

    def _on_textures_toggle(self, enabled):
        active = enabled and self.preset_combo.count() > 0
        self.preset_combo.setEnabled(active)
        self.preset_search.setEnabled(enabled)
        self.texture_sets_group.setEnabled(enabled)
        self.texture_params_group.setEnabled(enabled)
        self.map_group.setEnabled(enabled)
        self.texture_splitter.setVisible(enabled)
        if self._loading:
            return
        if not enabled:
            self._capture_texture_selection()
            return
        self._apply_pending_texture_sets()
        self._refresh_map_list()

    def _on_mesh_toggle(self, enabled):
        self.mesh_combo.setEnabled(enabled)
        self.mesh_options_widget.setVisible(enabled)

    def _set_all_texture_set_checks(self, checked):
        for i in range(self.texset_list.count()):
            item = self.texset_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
        self._update_texture_set_count()

    def _update_texture_set_count(self, _item=None):
        total = 0
        selected = 0
        for i in range(self.texset_list.count()):
            item = self.texset_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            total += 1
            if item.checkState() == QtCore.Qt.Checked:
                selected += 1
        self.texset_status.setText(f"Selected: {selected}/{total}")

    def _apply_pending_texture_sets(self):
        if self._pending_texture_sets is None:
            return
        selected = {name.lower() for name in self._pending_texture_sets if name}
        for i in range(self.texset_list.count()):
            item = self.texset_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            name = item.data(QtCore.Qt.UserRole) or item.text()
            item.setCheckState(
                QtCore.Qt.Checked if name.lower() in selected else QtCore.Qt.Unchecked
            )
        self._pending_texture_sets = None
        self._update_texture_set_count()

    def _populate_texture_sets(self):
        self.texset_list.blockSignals(True)
        self.texset_list.clear()
        texsets = get_all_texture_sets()
        if not texsets:
            item = QtWidgets.QListWidgetItem("No texture sets found")
            item.setFlags(QtCore.Qt.NoItemFlags)
            self.texset_list.addItem(item)
            self.texset_list.blockSignals(False)
            self._update_texture_set_count()
            return
        pending = self._pending_texture_sets
        pending_set = {name.lower() for name in pending or []}
        use_pending = pending is not None
        for texset in texsets:
            name = get_sp_name(texset)
            item = QtWidgets.QListWidgetItem(name)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            item.setForeground(self.texset_list.palette().color(QtGui.QPalette.Text))
            if use_pending:
                item.setCheckState(
                    QtCore.Qt.Checked if name.lower() in pending_set else QtCore.Qt.Unchecked
                )
            else:
                item.setCheckState(QtCore.Qt.Checked)
            item.setData(QtCore.Qt.UserRole, name)
            self.texset_list.addItem(item)
        self._pending_texture_sets = None
        self.texset_list.blockSignals(False)
        self._update_texture_set_count()

    def _set_all_map_checks(self, checked):
        for i in range(self.map_list.count()):
            item = self.map_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
        self._update_map_count()

    def _update_map_count(self, _item=None):
        total = 0
        selected = 0
        for i in range(self.map_list.count()):
            item = self.map_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            total += 1
            if item.checkState() == QtCore.Qt.Checked:
                selected += 1
        self.map_status.setText(f"Selected: {selected}/{total}")

    def _filter_presets(self, text):
        query = text.strip().lower()
        current_label = self.preset_combo.currentText()
        self.preset_combo.blockSignals(True)
        self.preset_combo.clear()
        self._presets = []
        selected_index = -1
        for preset in self._all_presets:
            label = preset["label"]
            name = preset.get("name", "")
            if query and query not in label.lower() and query not in name.lower():
                continue
            self.preset_combo.addItem(label)
            self._presets.append(preset)
            if current_label and label == current_label:
                selected_index = self.preset_combo.count() - 1
        if selected_index >= 0:
            self.preset_combo.setCurrentIndex(selected_index)
        self.preset_combo.blockSignals(False)
        self.preset_combo.setEnabled(self.textures_cb.isChecked() and self.preset_combo.count() > 0)
        self._refresh_map_list()

    def _capture_texture_selection(self):
        maps = []
        map_count = 0
        for i in range(self.map_list.count()):
            item = self.map_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            map_count += 1
            if item.checkState() == QtCore.Qt.Checked:
                maps.append(item.data(QtCore.Qt.UserRole))
        sets = []
        set_count = 0
        for i in range(self.texset_list.count()):
            item = self.texset_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            set_count += 1
            if item.checkState() == QtCore.Qt.Checked:
                sets.append(item.data(QtCore.Qt.UserRole))
        if map_count:
            self._pending_map_selection = maps
        if set_count:
            self._pending_texture_sets = sets

    def _refresh_map_list(self):
        self.map_list.clear()
        if (not self.textures_cb.isChecked() or self.preset_combo.count() == 0 or
                self.preset_combo.currentIndex() < 0 or not self._presets):
            self._update_map_count()
            return
        preset = self._presets[self.preset_combo.currentIndex()]
        output_maps = get_output_map_names(preset)
        if not output_maps:
            item = QtWidgets.QListWidgetItem("No output maps found for preset")
            item.setFlags(QtCore.Qt.NoItemFlags)
            self.map_list.addItem(item)
            self._update_map_count()
            return
        selected = None
        if self._pending_map_selection is not None:
            selected = {name.lower() for name in self._pending_map_selection if name}
        for map_name in output_maps:
            label = friendly_map_label(map_name)
            item = QtWidgets.QListWidgetItem(label)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            if selected is None or map_name.lower() in selected:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            item.setData(QtCore.Qt.UserRole, map_name)
            self.map_list.addItem(item)
        self._pending_map_selection = None
        self._update_map_count()

    def _open_output_dir(self):
        path = self.output_dir_edit.text().strip()
        if not path:
            return
        ensure_dir(path)
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))

    def _refresh_linked_blender_state(self):
        self._linked_blender_file = read_linked_blender_file(self._project_dir)
        detected_path = self._linked_blender_file or ""
        allow_temp_open = self.open_temp_blender_cb.isChecked()
        if detected_path and is_temp_blender_file(detected_path):
            self.detected_blender_label.setText(
                "Linked Blender project is unsaved.\n"
                f"Detected: {detected_path}"
            )
            self.detected_blender_label.setVisible(True)
            exists = False
            try:
                exists = Path(detected_path).is_file()
            except OSError:
                exists = False
            if allow_temp_open and exists:
                self.open_blender_cb.setEnabled(True)
                self.open_blender_cb.setToolTip(detected_path)
            else:
                self.open_blender_cb.setEnabled(False)
                self.open_blender_cb.setChecked(False)
                if allow_temp_open:
                    self.open_blender_cb.setToolTip("Temp linked Blender project not found.")
                else:
                    self.open_blender_cb.setToolTip("Linked Blender project is unsaved.")
            return
        exists = False
        if detected_path:
            try:
                exists = Path(detected_path).is_file()
            except OSError:
                exists = False
        if detected_path and not exists:
            self.detected_blender_label.setText(
                "Linked Blender project not found.\n"
                f"Detected: {detected_path}"
            )
        elif detected_path:
            self.detected_blender_label.setText(f"Detected: {detected_path}")
        else:
            self.detected_blender_label.setText("No linked Blender project detected.")
        self.detected_blender_label.setVisible(True)
        self.open_blender_cb.setEnabled(bool(exists))
        if not exists:
            self.open_blender_cb.setChecked(False)
            self.open_blender_cb.setToolTip("No linked Blender project found.")
        else:
            self.open_blender_cb.setToolTip(detected_path)

    def _select_preset_by_ref(self, preset_ref):
        if not preset_ref:
            return False
        target_name = preset_ref.get("name")
        target_kind = preset_ref.get("kind")
        if not target_name:
            return False
        selected_index = -1
        for i, preset in enumerate(self._presets):
            if (preset["name"].lower() == target_name.lower() and
                    (not target_kind or preset["kind"] == target_kind)):
                selected_index = i
                break
        if selected_index < 0:
            for i, preset in enumerate(self._presets):
                if preset["name"].lower() == target_name.lower():
                    selected_index = i
                    break
        if selected_index >= 0:
            self.preset_combo.blockSignals(True)
            self.preset_combo.setCurrentIndex(selected_index)
            self.preset_combo.blockSignals(False)
            return True
        return False

    def _apply_saved_state(self, state):
        if not state:
            return
        if "export_mesh" in state:
            self.mesh_cb.setChecked(bool(state["export_mesh"]))
        if "export_textures" in state:
            self.textures_cb.setChecked(bool(state["export_textures"]))
        if "open_temp_blender_project" in state:
            self.open_temp_blender_cb.setChecked(bool(state["open_temp_blender_project"]))
        if "open_blender_project" in state:
            self.open_blender_cb.setChecked(bool(state["open_blender_project"]))
        mesh_key = state.get("mesh_option")
        if mesh_key:
            for i in range(self.mesh_combo.count()):
                data = self.mesh_combo.itemData(i)
                if mesh_option_key(data) == mesh_key:
                    self.mesh_combo.setCurrentIndex(i)
                    break
        export_settings = state.get("export_settings") or {}
        if "file_format" in export_settings:
            idx = self.format_combo.findData(export_settings["file_format"])
            if idx >= 0:
                self.format_combo.setCurrentIndex(idx)
        if "bit_depth" in export_settings:
            idx = self.bitdepth_combo.findData(export_settings["bit_depth"])
            if idx >= 0:
                self.bitdepth_combo.setCurrentIndex(idx)
        if "size_log2" in export_settings:
            idx = self.res_combo.findData(export_settings["size_log2"])
            if idx >= 0:
                self.res_combo.setCurrentIndex(idx)
        if "padding_algorithm" in export_settings:
            idx = self.padding_combo.findData(export_settings["padding_algorithm"])
            if idx >= 0:
                self.padding_combo.setCurrentIndex(idx)
        if "dilation_distance" in export_settings:
            try:
                self.dilation_spin.setValue(int(export_settings["dilation_distance"]))
            except (TypeError, ValueError):
                pass
        if "dithering" in export_settings:
            self.dither_cb.setChecked(bool(export_settings["dithering"]))
        splitter_sizes = state.get("texture_splitter_sizes")
        if isinstance(splitter_sizes, list) and len(splitter_sizes) == 2:
            try:
                self.texture_splitter.setSizes([int(splitter_sizes[0]), int(splitter_sizes[1])])
            except (TypeError, ValueError):
                pass
        preset_ref = state.get("preset")
        if preset_ref:
            self._select_preset_by_ref(preset_ref)
        self._pending_map_selection = state.get("output_maps")
        if "texture_sets" in state:
            texture_sets = state.get("texture_sets")
            if texture_sets is None:
                self._pending_texture_sets = None
                self._set_all_texture_set_checks(True)
            else:
                self._pending_texture_sets = texture_sets
                self._apply_pending_texture_sets()
        self._refresh_map_list()
        self._refresh_linked_blender_state()

    def _serialize_options(self, options):
        preset = options.get("preset")
        preset_ref = None
        if preset:
            preset_ref = {"kind": preset.get("kind"), "name": preset.get("name")}
        mesh_option = options.get("mesh_option")
        mesh_key = mesh_option_key(mesh_option) if mesh_option is not None else None
        texture_sets = options.get("texture_sets", [])
        texture_sets_value = texture_sets
        try:
            all_sets = [get_sp_name(texset) for texset in get_all_texture_sets()]
        except Exception:
            all_sets = None
        if all_sets:
            all_set_names = {name.lower() for name in all_sets if name}
            if texture_sets and len(texture_sets) == len(all_sets):
                if all(name.lower() in all_set_names for name in texture_sets):
                    texture_sets_value = None
        return {
            "export_mesh": options.get("export_mesh", True),
            "export_textures": options.get("export_textures", True),
            "open_blender_project": options.get("open_blender_project", False),
            "open_temp_blender_project": options.get("open_temp_blender_project", False),
            "mesh_option": mesh_key,
            "preset": preset_ref,
            "output_maps": options.get("output_maps", []),
            "export_settings": options.get("export_settings", {}),
            "texture_sets": texture_sets_value,
            "texture_splitter_sizes": (
                self.texture_splitter.sizes() if hasattr(self, "texture_splitter") else None
            ),
        }

    def _reload_user_presets(self, select_name=None):
        current = ""
        if self.user_preset_combo.count() > 0:
            current = self.user_preset_combo.currentText()
        if select_name is None:
            select_name = current
        self.user_preset_combo.blockSignals(True)
        self.user_preset_combo.clear()
        self.user_preset_combo.addItem(DEFAULT_USER_PRESET_NAME)
        selected_index = 0
        if select_name and select_name.lower() == DEFAULT_USER_PRESET_NAME.lower():
            selected_index = 0
        for preset in self._user_presets:
            name = preset.get("name")
            if not name:
                continue
            self.user_preset_combo.addItem(name)
            if select_name and name.lower() == select_name.lower():
                selected_index = self.user_preset_combo.count() - 1
        self.user_preset_combo.setCurrentIndex(selected_index)
        self.user_preset_combo.blockSignals(False)
        self._update_preset_buttons()

    def _update_preset_buttons(self):
        index = self.user_preset_combo.currentIndex()
        self.delete_preset_btn.setEnabled(index > 0)

    def _apply_user_preset_selection(self, index):
        if self._loading or index < 0:
            return
        if index == 0:
            options = self._default_preset_options or self._build_default_preset_options()
        else:
            preset = self._user_presets[index - 1]
            options = preset.get("options", {})
        self._loading = True
        self._apply_saved_state(options)
        self._on_textures_toggle(self.textures_cb.isChecked())
        self._loading = False
        self._update_preset_buttons()

    def _save_user_preset(self):
        current_name = ""
        if self.user_preset_combo.currentIndex() > 0:
            current_name = self.user_preset_combo.currentText()
        name, ok = QtWidgets.QInputDialog.getText(
            self,
            "Save Bridge Preset",
            "Preset name:",
            text=current_name,
        )
        if not ok:
            return
        name = name.strip()
        if not name:
            return
        if name.lower() == DEFAULT_USER_PRESET_NAME.lower():
            show_message(
                "GoB Bridge",
                f"'{DEFAULT_USER_PRESET_NAME}' is reserved and cannot be overwritten.",
                QtWidgets.QMessageBox.Warning,
            )
            return
        options = self._serialize_options(self.get_options())
        existing_index = None
        for i, preset in enumerate(self._user_presets):
            if preset.get("name", "").lower() == name.lower():
                existing_index = i
                break
        if existing_index is not None:
            self._user_presets[existing_index]["options"] = options
        else:
            self._user_presets.append({"name": name, "options": options})
        save_persistent_state(user_presets=self._user_presets)
        self._reload_user_presets(select_name=name)

    def _delete_user_preset(self):
        index = self.user_preset_combo.currentIndex()
        if index <= 0:
            return
        name = self.user_preset_combo.currentText()
        response = QtWidgets.QMessageBox.question(
            self,
            "Delete Bridge Preset",
            f"Delete preset '{name}'?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        if response != QtWidgets.QMessageBox.Yes:
            return
        del self._user_presets[index - 1]
        save_persistent_state(user_presets=self._user_presets)
        self._reload_user_presets()

    def _build_default_preset_options(self):
        mesh_option = None
        if self.mesh_combo.count() > 0:
            mesh_option = self.mesh_combo.itemData(0)
        preset_ref = None
        preset = self._find_default_export_preset()
        if preset:
            preset_ref = {
                "kind": preset.get("kind"),
                "name": preset.get("name"),
            }
        return {
            "mesh_option": mesh_option,
            "preset": preset_ref,
            "output_maps": None,
            "export_settings": dict(DEFAULT_EXPORT_SETTINGS),
            "texture_sets": None,
            "texture_splitter_sizes": self.texture_splitter.sizes(),
            "open_blender_project": False,
            "open_temp_blender_project": False,
        }

    def _find_default_export_preset(self):
        candidates = self._all_presets or self._presets or []
        if not candidates:
            return None
        def haystack(preset):
            return f"{preset.get('name', '')} {preset.get('label', '')}".lower()
        for preset in candidates:
            text = haystack(preset)
            if "blender" in text and ("principled" in text or "bsdf" in text):
                return preset
        for preset in candidates:
            text = haystack(preset)
            if "blender" in text:
                return preset
        for preset in candidates:
            text = haystack(preset)
            if "principled" in text or "bsdf" in text:
                return preset
        return candidates[0]

    def get_options(self):
        preset = None
        if self._presets and self.preset_combo.currentIndex() >= 0:
            preset = self._presets[self.preset_combo.currentIndex()]
        output_maps = []
        for i in range(self.map_list.count()):
            item = self.map_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            if item.checkState() == QtCore.Qt.Checked:
                output_maps.append(item.data(QtCore.Qt.UserRole))
        texture_sets = []
        for i in range(self.texset_list.count()):
            item = self.texset_list.item(i)
            if not (item.flags() & QtCore.Qt.ItemIsUserCheckable):
                continue
            if item.checkState() == QtCore.Qt.Checked:
                texture_sets.append(item.data(QtCore.Qt.UserRole))
        export_settings = {
            "file_format": self.format_combo.currentData(),
            "bit_depth": self.bitdepth_combo.currentData(),
            "size_log2": self.res_combo.currentData(),
            "padding_algorithm": self.padding_combo.currentData(),
            "dilation_distance": self.dilation_spin.value(),
            "dithering": self.dither_cb.isChecked(),
        }
        return {
            "export_mesh": self.mesh_cb.isChecked(),
            "export_textures": self.textures_cb.isChecked(),
            "open_blender_project": self.open_blender_cb.isChecked(),
            "open_temp_blender_project": self.open_temp_blender_cb.isChecked(),
            "mesh_option": self.mesh_combo.currentData(),
            "preset": preset,
            "output_maps": output_maps,
            "export_settings": export_settings,
            "texture_sets": texture_sets,
        }

    def persist_last_settings(self, options):
        state = self._serialize_options(options)
        save_persistent_state(
            last_settings=state,
            user_presets=self._user_presets,
            project_dir=self._project_dir,
        )


class QuickPanel(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(2)
        self.import_btn = QtWidgets.QToolButton()
        self.export_btn = QtWidgets.QToolButton()
        self.import_btn.setText("GoB Import")
        self.export_btn.setText("GoB Export")
        for btn in (self.import_btn, self.export_btn):
            btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
            btn.setAutoRaise(True)
            btn.setCursor(QtCore.Qt.PointingHandCursor)
        self.import_btn.setToolTip("Import from Blender")
        self.export_btn.setToolTip("Send to Blender")
        self.import_btn.clicked.connect(import_from_blender)
        self.export_btn.clicked.connect(send_to_blender)
        layout.addWidget(self.import_btn)
        layout.addWidget(self.export_btn)


def _resolve_export_shelf():
    shelf_enum = getattr(sp.ui, "Shelf", None)
    if shelf_enum:
        for name in ("Export", "export", "EXPORT"):
            if hasattr(shelf_enum, name):
                return getattr(shelf_enum, name)
    return None


def _add_quick_panel_ui():
    global _quick_panel_widget
    if _quick_panel_widget is not None:
        return None
    _quick_panel_widget = QuickPanel()
    if hasattr(sp.ui, "add_shelf_widget"):
        shelf = _resolve_export_shelf()
        if shelf is not None:
            try:
                return sp.ui.add_shelf_widget(shelf, _quick_panel_widget)
            except Exception:
                pass
        try:
            return sp.ui.add_shelf_widget("Export", _quick_panel_widget)
        except Exception:
            pass
    if hasattr(sp.ui, "add_dock_widget"):
        try:
            return sp.ui.add_dock_widget("GoB Bridge", _quick_panel_widget)
        except Exception:
            pass
    return None


def clear_auto_import_flag(manifest_path, manifest):
    if not manifest_path or not manifest or not manifest.get("auto_import"):
        return
    manifest["auto_import"] = False
    try:
        write_manifest(manifest_path, manifest)
    except Exception:
        return


def manifest_timestamp(manifest, manifest_path):
    if isinstance(manifest, dict):
        for key in ("auto_import_at", "timestamp"):
            value = manifest.get(key)
            if isinstance(value, (int, float)):
                return float(value)
    try:
        return float(Path(manifest_path).stat().st_mtime)
    except Exception:
        return 0.0


def manifest_targets_current_project(manifest, manifest_path):
    if not sp.project.is_open():
        return True
    sp_project_file = get_sp_project_file_path_or_temp()
    manifest_sp_file = manifest.get("sp_project_file") or manifest.get("sp_project_path")
    if sp_project_file and manifest_sp_file:
        if paths_match(sp_project_file, manifest_sp_file):
            return True
    try:
        current_dir = get_project_dir()
    except Exception:
        current_dir = None
    manifest_dir = project_dir_from_manifest_path(manifest_path)
    if current_dir and manifest_dir:
        return current_dir == manifest_dir
    current_name = get_project_name()
    manifest_name = sanitize_name(str(manifest.get("project") or ""))
    if current_name and manifest_name:
        return current_name.lower() == manifest_name.lower()
    return False


def parse_force_new_token(argv=None):
    args = argv if argv is not None else sys.argv
    for arg in args:
        if not isinstance(arg, str):
            continue
        for prefix in FORCE_NEW_TOKEN_ARG_PREFIXES:
            if arg.startswith(prefix):
                token = arg.split("=", 1)[-1].strip()
                if token:
                    return token
    return ""


def load_force_new_token():
    token = os.environ.get(FORCE_NEW_TOKEN_ENV)
    if token:
        return token.strip()
    return parse_force_new_token()


def manifest_force_new_token(manifest):
    if not isinstance(manifest, dict):
        return ""
    token = manifest.get("force_new_token")
    return str(token).strip() if token else ""


def force_new_token_matches(manifest):
    token = manifest_force_new_token(manifest)
    if not token:
        return False
    return bool(_force_new_token and token == _force_new_token)


def should_accept_force_new_manifest(manifest):
    token = manifest_force_new_token(manifest)
    if token:
        if _force_new_token and token == _force_new_token:
            return True
        return not sp.project.is_open()
    return not sp.project.is_open()


def import_from_blender(manifest_path=None, clear_auto_import=False):
    global _auto_import_in_progress
    bridge_roots = get_candidate_bridge_roots()
    project_dir = None
    manifest = None
    if manifest_path:
        manifest_path = Path(manifest_path)
        if manifest_path.exists():
            manifest = read_manifest(manifest_path)
            project_dir = project_dir_from_manifest_path(manifest_path)
    else:
        sp_project_file = get_sp_project_file_path_or_temp() if sp.project.is_open() else ""
        if sp_project_file:
            candidate = find_manifest_for_sp_project(
                bridge_roots,
                sp_project_file,
                source="blender",
            )
            if candidate:
                manifest_path = candidate
                manifest = read_manifest(manifest_path)
                project_dir = project_dir_from_manifest_path(manifest_path)
        if not manifest:
            project_dir = get_project_dir()
            manifest_path = find_project_manifest_path(project_dir)
            if manifest_path and manifest_path.exists():
                manifest = read_manifest(manifest_path)
        if not manifest or manifest.get("source") != "blender":
            latest = find_latest_manifest(bridge_roots, source="blender")
            if not latest:
                show_message("GoB Bridge", "No Blender export manifest found.", QtWidgets.QMessageBox.Warning)
                return
            manifest_path = latest
            manifest = read_manifest(manifest_path)
            project_dir = project_dir_from_manifest_path(manifest_path)
    if not manifest:
        show_message("GoB Bridge", "Failed to read Blender export manifest.", QtWidgets.QMessageBox.Warning)
        return
    linked_blender_file = str(manifest.get("blender_file") or "")
    project_name = sanitize_name(str(manifest.get("project") or ""))

    mesh_path = manifest.get("mesh_fbx")
    high_path = manifest.get("high_mesh_fbx")
    high_exported = manifest.get("high_mesh_exported")
    had_high_path = bool(high_path)
    if high_path:
        high_path = Path(high_path)
        if not high_path.is_absolute() and project_dir:
            high_path = project_dir / high_path
        high_path = str(high_path)
    if high_path and not Path(high_path).is_file():
        fallback = project_dir / BLENDER_HIGH_FILENAME if project_dir else None
        if fallback and fallback.exists():
            high_path = str(fallback)
        else:
            alt = find_mesh_in_roots(bridge_roots, project_name, BLENDER_HIGH_FILENAME)
            if alt:
                high_path = alt
            else:
                high_path = None
                if had_high_path:
                    clear_high_poly_when_ready()
    if not high_path and high_exported is False:
        clear_high_poly_when_ready()
    force_new_project = bool(manifest.get("force_new_project"))
    if mesh_path:
        mesh_path = Path(mesh_path)
        if not mesh_path.is_absolute() and project_dir:
            mesh_path = project_dir / mesh_path
        mesh_path = str(mesh_path)
    if not mesh_path or not Path(mesh_path).is_file():
        fallback = project_dir / BLENDER_EXPORT_FILENAME if project_dir else None
        if fallback and fallback.exists():
            mesh_path = str(fallback)
        else:
            alt = find_mesh_in_roots(bridge_roots, project_name, BLENDER_EXPORT_FILENAME)
            if alt:
                mesh_path = alt
    if not mesh_path or not Path(mesh_path).is_file():
        show_message("GoB Bridge", "Blender FBX file not found.", QtWidgets.QMessageBox.Warning)
        return

    if sp.project.is_open() and not force_new_project:
        if not manifest_targets_current_project(manifest, manifest_path):
            if clear_auto_import:
                return
            show_message(
                "GoB Bridge",
                "Blender export targets a different project. Close the current project and import again.",
                QtWidgets.QMessageBox.Warning,
            )
            return
        if not sp.project.is_in_edition_state():
            return
        def make_reload_settings(preserve):
            try:
                return sp.project.MeshReloadingSettings(
                    import_cameras=False,
                    preserve_strokes=preserve,
                )
            except Exception:
                return None

        settings = make_reload_settings(True)
        fallback_settings = make_reload_settings(False)
        attempted_fallback = False

        def _finish_reload(success):
            global _auto_import_in_progress
            _auto_import_in_progress = False
            if success:
                if high_path:
                    apply_high_poly_when_ready(high_path)
                if clear_auto_import:
                    clear_auto_import_flag(manifest_path, manifest)
                sp_project_file = (
                    get_sp_project_file_path_or_temp()
                    or str(manifest.get("sp_project_file") or "")
                    or str(manifest.get("sp_project_path") or "")
                )
                if linked_blender_file:
                    update_project_settings(
                        {"linked_blender_file": linked_blender_file},
                        project_dir=get_project_dir(),
                    )
                    if sp_project_file:
                        update_link_registry(
                            sp_project_file=sp_project_file,
                            blender_file=linked_blender_file,
                            update_blender_link=not force_new_project,
                        )
                if sp_project_file:
                    write_manifest_sp_project_file(manifest, project_dir, sp_project_file)
                write_active_sp_info()
                show_message("GoB Bridge", "Mesh reloaded from Blender.")
            else:
                show_message("GoB Bridge", "Mesh reload failed.", QtWidgets.QMessageBox.Warning)

        def _on_reload(status):
            if status == sp.project.ReloadMeshStatus.SUCCESS:
                _finish_reload(True)
                return
            nonlocal attempted_fallback
            if not attempted_fallback and fallback_settings:
                attempted_fallback = True
                try:
                    sp.project.reload_mesh(mesh_path, fallback_settings, _on_reload)
                    return
                except Exception:
                    pass
            _finish_reload(False)

        _auto_import_in_progress = True
        try:
            sp.project.reload_mesh(mesh_path, settings or fallback_settings, _on_reload)
        except Exception as exc:
            _auto_import_in_progress = False
            message = str(exc).lower()
            if "busy" in message:
                global _auto_import_busy_until
                _auto_import_busy_until = time.time() + 2.0
                return
            if (fallback_settings and settings and not attempted_fallback and
                    any(key in message for key in ("stroke", "preserv", "scale", "unit"))):
                attempted_fallback = True
                try:
                    _auto_import_in_progress = True
                    sp.project.reload_mesh(mesh_path, fallback_settings, _on_reload)
                    return
                except Exception:
                    _auto_import_in_progress = False
            show_message("GoB Bridge", f"Mesh reload failed: {exc}", QtWidgets.QMessageBox.Warning)
        return

    if force_new_project:
        if not should_accept_force_new_manifest(manifest):
            if not clear_auto_import:
                show_message(
                    "GoB Bridge",
                    "New instance project detected for a different Painter instance.",
                    QtWidgets.QMessageBox.Warning,
                )
            return
        if sp.project.is_open():
            try:
                sp.project.close()
            except Exception:
                pass
    project_settings = build_project_settings(manifest.get("sp_project_settings"))
    try:
        _auto_import_in_progress = True
        create_kwargs = {"mesh_file_path": mesh_path}
        if project_settings:
            create_kwargs["settings"] = project_settings
        try:
            sp.project.create(**create_kwargs)
        except TypeError:
            if "import_settings" in create_kwargs:
                create_kwargs.pop("import_settings", None)
                sp.project.create(**create_kwargs)
            else:
                raise
        if force_new_project and force_new_token_matches(manifest):
            global _force_new_token
            _force_new_token = ""
        if high_path:
            apply_high_poly_when_ready(high_path)
        if clear_auto_import:
            clear_auto_import_flag(manifest_path, manifest)
        sp_project_file = (
            get_sp_project_file_path_or_temp()
            or str(manifest.get("sp_project_file") or "")
            or str(manifest.get("sp_project_path") or "")
        )
        if linked_blender_file:
            update_project_settings(
                {"linked_blender_file": linked_blender_file},
                project_dir=get_project_dir(),
            )
            if sp_project_file:
                update_link_registry(
                    sp_project_file=sp_project_file,
                    blender_file=linked_blender_file,
                    update_blender_link=not force_new_project,
                )
        if force_new_project:
            update_project_settings(
                {"force_new_project": True},
                project_dir=project_dir,
            )
        if sp_project_file:
            write_manifest_sp_project_file(manifest, project_dir, sp_project_file)
        write_active_sp_info()
        _auto_import_in_progress = False
    except Exception:
        _auto_import_in_progress = False
        show_message("GoB Bridge", "Failed to create project from Blender mesh.", QtWidgets.QMessageBox.Warning)


def send_to_blender():
    if not sp.project.is_open():
        show_message("GoB Bridge", "Open a project before exporting.", QtWidgets.QMessageBox.Warning)
        return

    dialog = ExportDialog()
    if dialog.exec() != QtWidgets.QDialog.Accepted:
        return

    options = dialog.get_options()
    if not options["export_mesh"] and not options["export_textures"]:
        show_message("GoB Bridge", "Select mesh and/or textures to export.", QtWidgets.QMessageBox.Warning)
        return
    dialog.persist_last_settings(options)

    allow_temp_open = bool(options.get("open_temp_blender_project"))
    sp_project_file = get_sp_project_file_path_or_temp()
    project_dir = project_dir_for_send(sp_project_file)
    write_bridge_root_hint(project_dir.parent)
    ensure_dir(project_dir)

    existing_manifest = read_manifest(find_project_manifest_path(project_dir))
    existing_sp_project_file = ""
    if isinstance(existing_manifest, dict):
        existing_sp_project_file = str(existing_manifest.get("sp_project_file") or "")
    linked_blender_file = read_linked_blender_file(project_dir)
    linked_blender_exists = ""
    if linked_blender_file:
        try:
            if Path(linked_blender_file).is_file():
                if (not is_temp_blender_file(linked_blender_file)) or allow_temp_open:
                    linked_blender_exists = linked_blender_file
                else:
                    linked_blender_exists = ""
        except OSError:
            linked_blender_exists = ""
    force_new_project = bool(existing_manifest and existing_manifest.get("force_new_project"))
    if not force_new_project:
        project_settings = load_project_settings(project_dir)
        if isinstance(project_settings, dict):
            force_new_project = bool(project_settings.get("force_new_project"))
    primary_sp_project_file = ""
    if force_new_project and linked_blender_file:
        primary_sp_project_file = resolve_primary_sp_project_for_blender(
            linked_blender_file,
            sp_project_file,
        )

    manifest = {
        "version": 1,
        "source": "substance_painter",
        "project": get_project_name(),
        "timestamp": time.time(),
    }
    if linked_blender_file:
        manifest["blender_file"] = linked_blender_file
    if sp_project_file:
        manifest["sp_project_file"] = sp_project_file
    elif existing_sp_project_file:
        manifest["sp_project_file"] = existing_sp_project_file
    if primary_sp_project_file:
        manifest["link_sp_project_file"] = primary_sp_project_file
    if force_new_project:
        manifest["force_new_project"] = True
    if sp_project_file and linked_blender_file:
        update_link_registry(
            sp_project_file=sp_project_file,
            blender_file=linked_blender_file,
            update_blender_link=not primary_sp_project_file,
        )
        if primary_sp_project_file:
            update_link_registry(
                sp_project_file=primary_sp_project_file,
                blender_file=linked_blender_file,
            )
    exported_any = False
    mesh_exported = False
    texture_errors = []
    texture_warnings = []

    if options["export_mesh"]:
        mesh_path = project_dir / SP_EXPORT_FILENAME
        result = sp.export.export_mesh(str(mesh_path), options["mesh_option"])
        if result.status != sp.export.ExportStatus.Success:
            show_message("GoB Bridge", result.message, QtWidgets.QMessageBox.Warning)
        else:
            manifest["mesh_fbx"] = str(mesh_path)
            mesh_exported = True
            exported_any = True

    if options["export_textures"]:
        preset = options.get("preset") or pick_export_preset()
        if not preset:
            texture_errors.append("No export preset found.")
        else:
            normal_format = infer_normal_map_format_from_preset(preset)
            if not normal_format:
                normal_format = get_sp_normal_map_format()
            if normal_format:
                manifest["normal_map_format"] = normal_format
            textures_dir = project_dir / "textures"
            if textures_dir.exists():
                try:
                    shutil.rmtree(textures_dir)
                except OSError:
                    texture_errors.append("Failed to clear old texture exports.")
            ensure_dir(textures_dir)
            output_maps = options.get("output_maps") or []
            texture_sets = options.get("texture_sets") or []
            if not output_maps:
                texture_errors.append("No output maps selected.")
            elif not texture_sets:
                texture_errors.append("No texture sets selected.")
            else:
                stack_roots = collect_stack_roots(texture_sets)
                stacks = [stack for _, stack in stack_roots]
                missing_before = collect_missing_map_channels(
                    preset,
                    output_maps,
                    texture_sets,
                    stack_roots=stack_roots,
                )
                auto_enabled = auto_enable_missing_channels(
                    missing_before,
                    texture_sets,
                    stack_roots=stack_roots,
                )
                preset_for_export = preset
                sanitized_maps, changed, removed_channels = sanitize_map_definitions(
                    preset,
                    texture_sets,
                    stacks=stacks,
                )
                if sanitized_maps and changed:
                    base_name = preset.get("name") or "Custom Preset"
                    custom_name = f"{base_name} (GoB)"
                    preset_for_export = {
                        "kind": "custom",
                        "name": custom_name,
                        "definition": {
                            "name": custom_name,
                            "maps": sanitized_maps,
                        },
                    }
                basecolor_has_opacity = preset_basecolor_has_opacity(
                    preset_for_export,
                    stacks=stacks,
                )
                manifest["basecolor_has_opacity"] = basecolor_has_opacity
                export_settings = options.get("export_settings") or DEFAULT_EXPORT_SETTINGS
                export_params = build_export_parameters(export_settings)
                export_list = build_export_list_for_preset(
                    preset_for_export,
                    output_maps,
                    texture_sets,
                    stack_roots=stack_roots,
                )
                if not export_list:
                    texture_errors.append("No matching maps found for the current texture sets.")
                else:
                    missing_after = collect_missing_map_channels(
                        preset_for_export,
                        output_maps,
                        texture_sets,
                        stack_roots=stack_roots,
                    )
                    selected_roots = [root for root, _ in stack_roots]
                    selected_map_names = [name for name in output_maps if name]
                    exported_by_root = {}
                    for entry in export_list:
                        root = entry.get("rootPath")
                        maps = entry.get("filter", {}).get("outputMaps", []) or []
                        if root:
                            exported_by_root[root] = set(maps)
                    roots_without_exports = [
                        root for root in selected_roots if not exported_by_root.get(root)
                    ]
                    skipped_missing = {}
                    for map_name, roots in missing_after.items():
                        for root, names in roots.items():
                            if map_name not in exported_by_root.get(root, set()):
                                skipped_missing.setdefault(map_name, {})[root] = names
                    skipped_unknown = {}
                    for root in selected_roots:
                        if root in roots_without_exports:
                            continue
                        exported = exported_by_root.get(root, set())
                        for map_name in selected_map_names:
                            if map_name in exported:
                                continue
                            if root in skipped_missing.get(map_name, {}):
                                continue
                            skipped_unknown.setdefault(map_name, []).append(root)
                    if (auto_enabled or removed_channels or skipped_missing or
                            skipped_unknown or roots_without_exports):
                        preset_label = preset.get("label") or preset.get("name") or "Selected template"
                        warning_lines = [f"Template: {preset_label}"]
                        if auto_enabled:
                            enabled_parts = []
                            for root in sorted(auto_enabled):
                                names = sorted(auto_enabled[root])
                                labels = ", ".join(channel_display_name(name) for name in names)
                                enabled_parts.append(f"{root}: {labels}")
                            if enabled_parts:
                                warning_lines.append(
                                    "Auto-enabled channels:\n" + "\n".join(enabled_parts)
                                )
                        if removed_channels:
                            removed_lines = []
                            for map_name in sorted(removed_channels):
                                names = sorted(removed_channels[map_name])
                                labels = ", ".join(channel_display_name(name) for name in names)
                                removed_lines.append(
                                    f"{friendly_map_label(map_name)}: {labels}"
                                )
                            if removed_lines:
                                warning_lines.append(
                                    "Channels removed from the template (missing on some selected texture sets):\n"
                                    + "\n".join(removed_lines)
                                )
                        if roots_without_exports:
                            warning_lines.append(
                                "No output maps were generated for:\n"
                                + "\n".join(sorted(roots_without_exports))
                            )
                        if skipped_missing:
                            missing_lines = []
                            for map_name in sorted(skipped_missing):
                                roots = skipped_missing[map_name]
                                root_parts = []
                                for root in sorted(roots):
                                    names = sorted(set(roots[root]))
                                    labels = ", ".join(channel_display_name(name) for name in names)
                                    root_parts.append(f"{root}: {labels}")
                                missing_lines.append(
                                    f"{friendly_map_label(map_name)} -> " + "; ".join(root_parts)
                                )
                            if missing_lines:
                                warning_lines.append(
                                    "Skipped maps for some texture sets (missing channels):\n"
                                    + "\n".join(missing_lines)
                                )
                        if skipped_unknown:
                            unknown_lines = []
                            for map_name in sorted(skipped_unknown):
                                roots = "; ".join(sorted(skipped_unknown[map_name]))
                                unknown_lines.append(
                                    f"{friendly_map_label(map_name)} -> {roots}"
                                )
                            if unknown_lines:
                                warning_lines.append(
                                    "Skipped maps for some texture sets (template output not generated):\n"
                                    + "\n".join(unknown_lines)
                                )
                        warning_lines.append(
                            "Enable missing channels via Texture Set Settings > Channels (+) "
                            "for the affected texture sets."
                        )
                        texture_warnings.append("\n\n".join(warning_lines))
                    attempts = []
                    if preset_for_export["kind"] == "custom":
                        attempts.append(("custom", preset_for_export))
                    else:
                        attempts.append(("url", preset_for_export))
                        if preset_for_export["kind"] == "resource":
                            for fallback in collect_export_presets():
                                if (fallback["kind"] == "predefined" and
                                        fallback["name"].lower() == preset_for_export["name"].lower()):
                                    attempts.append(("predefined_url", fallback))
                                    break
                    tried = False
                    for label, preset_info in attempts:
                        if preset_info["kind"] == "custom":
                            export_config = build_custom_export_config(
                                textures_dir,
                                preset_info["definition"],
                                output_maps,
                                export_params,
                                export_list=export_list,
                            )
                        else:
                            export_config = build_export_config(
                                textures_dir,
                                preset_info["url"],
                                output_maps,
                                export_params,
                                export_list=export_list,
                            )
                        if not export_config:
                            texture_errors.append("No texture sets available to export.")
                            break
                        tried = True
                        append_log(project_dir, f"Export attempt ({label})", export_config)
                        try:
                            export_result = sp.export.export_project_textures(export_config)
                        except Exception as exc:
                            texture_errors.append(f"{label}: {exc}")
                            append_log(project_dir, f"Export failed ({label}): {exc}")
                            continue
                        if export_result.status != sp.export.ExportStatus.Success:
                            texture_errors.append(f"{label}: {export_result.message}")
                            append_log(project_dir, f"Export failed ({label}): {export_result.message}")
                            continue
                        textures = []
                        for files in export_result.textures.values():
                            for file in files:
                                if not file:
                                    continue
                                path = Path(file)
                                if not path.is_absolute():
                                    path = textures_dir / path
                                textures.append(str(path))
                        manifest["textures_dir"] = str(textures_dir)
                        manifest["textures"] = textures
                        exported_any = True
                        texture_errors = []
                        break
                    if not tried and not texture_errors:
                        texture_errors.append("No export preset available.")

    manifest["mesh_exported"] = mesh_exported
    if not exported_any:
        if options.get("open_blender_project") and linked_blender_exists:
            open_linked_blender_file(linked_blender_exists, project_dir=project_dir)
        show_message("GoB Bridge", "Nothing was exported.", QtWidgets.QMessageBox.Warning)
        return

    manifest_path = project_manifest_path(project_dir)
    ensure_dir(manifest_path.parent)
    write_manifest(manifest_path, manifest)
    if options.get("open_blender_project"):
        if linked_blender_exists:
            if not open_linked_blender_file(linked_blender_exists, project_dir=project_dir):
                show_message(
                    "GoB Bridge",
                    "Failed to open the linked Blender project.",
                    QtWidgets.QMessageBox.Warning,
                )
        else:
            show_message(
                "GoB Bridge",
                "Linked Blender project not found.",
                QtWidgets.QMessageBox.Warning,
            )
    if texture_errors:
        details = "\n".join(texture_errors)
        show_message("GoB Bridge", f"Texture export failed:\n{details}", QtWidgets.QMessageBox.Warning)
    elif texture_warnings:
        show_message("GoB Bridge", "Export complete.", QtWidgets.QMessageBox.Information)
    else:
        show_message("GoB Bridge", "Export complete. Use Blender to import.")


_ui_elements = []
_quick_panel_widget = None
_auto_import_timer = None
_auto_import_last_time = 0.0
_auto_import_last_path = None
_auto_import_in_progress = False
_auto_import_busy_until = 0.0
_auto_import_last_scan = 0.0
_auto_import_last_active_write = 0.0
AUTO_IMPORT_POLL_INTERVAL_MS = 3000
AUTO_IMPORT_SCAN_COOLDOWN = 5.0
AUTO_IMPORT_ACTIVE_WRITE_INTERVAL = 5.0


def start_plugin():
    global _force_new_token
    _force_new_token = load_force_new_token()
    action_import = QtGui.QAction("GoB Bridge: Import from Blender")
    action_import.triggered.connect(import_from_blender)
    sp.ui.add_action(sp.ui.ApplicationMenu.File, action_import)
    _ui_elements.append(action_import)

    action_send = QtGui.QAction("GoB Bridge: Send to Blender")
    action_send.triggered.connect(send_to_blender)
    sp.ui.add_action(sp.ui.ApplicationMenu.File, action_send)
    _ui_elements.append(action_send)

    try:
        quick_element = _add_quick_panel_ui()
        if quick_element is not None:
            _ui_elements.append(quick_element)
    except Exception:
        pass
    write_active_sp_info()
    try:
        start_update_check(auto_prompt=True)
    except Exception:
        pass
    def _auto_import_poll():
        global _auto_import_last_time
        global _auto_import_last_path
        global _auto_import_in_progress
        global _auto_import_busy_until
        global _auto_import_last_scan
        global _auto_import_last_active_write
        now = time.time()
        if now - _auto_import_last_active_write >= AUTO_IMPORT_ACTIVE_WRITE_INTERVAL:
            write_active_sp_info()
            _auto_import_last_active_write = now
        if _auto_import_in_progress:
            return
        if _auto_import_busy_until and now < _auto_import_busy_until:
            return
        manifest_path = None
        manifest = None
        active_blender = find_active_blender_info()
        if active_blender:
            project_dir = active_blender.get("project_dir")
            if project_dir:
                candidate_path = find_project_manifest_path(project_dir)
                if candidate_path and candidate_path.exists():
                    manifest = read_manifest(candidate_path)
                    if manifest and manifest.get("source") == "blender" and manifest.get("auto_import"):
                        manifest_path = candidate_path
                    else:
                        manifest = None
        if not manifest_path:
            if now - _auto_import_last_scan < AUTO_IMPORT_SCAN_COOLDOWN:
                return
            _auto_import_last_scan = now
            manifest_path = find_latest_manifest(get_candidate_bridge_roots(), source="blender")
            if not manifest_path:
                return
        if manifest is None:
            manifest = read_manifest(manifest_path)
        if not manifest or not manifest.get("auto_import"):
            return
        ts = manifest_timestamp(manifest, manifest_path)
        if manifest.get("force_new_project"):
            if not should_accept_force_new_manifest(manifest):
                return
        elif sp.project.is_open():
            if not manifest_targets_current_project(manifest, manifest_path):
                return
        if (str(manifest_path) == _auto_import_last_path and ts <= _auto_import_last_time):
            return
        import_from_blender(manifest_path=manifest_path, clear_auto_import=True)
        if not _auto_import_in_progress:
            _auto_import_last_time = ts
            _auto_import_last_path = str(manifest_path)

    global _auto_import_timer
    try:
        _auto_import_timer = QtCore.QTimer()
        _auto_import_timer.setInterval(AUTO_IMPORT_POLL_INTERVAL_MS)
        _auto_import_timer.timeout.connect(_auto_import_poll)
        _auto_import_timer.start()
    except Exception:
        _auto_import_timer = None


def close_plugin():
    for element in _ui_elements:
        sp.ui.delete_ui_element(element)
    _ui_elements.clear()
    global _quick_panel_widget
    _quick_panel_widget = None
    global _auto_import_timer
    if _auto_import_timer is not None:
        try:
            _auto_import_timer.stop()
        except Exception:
            pass
    _auto_import_timer = None
    global _auto_import_last_time
    global _auto_import_last_path
    _auto_import_last_time = 0.0
    _auto_import_last_path = None
    global _auto_import_in_progress
    global _auto_import_busy_until
    _auto_import_in_progress = False
    _auto_import_busy_until = 0.0
    global _force_new_token
    _force_new_token = ""
    write_active_sp_info()


if __name__ == "__main__":
    start_plugin()
