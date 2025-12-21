import json
import os
import re
import shutil
import time
import urllib.error
import urllib.request
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets, QtNetwork

import substance_painter as sp
import substance_painter.event as sp_event


BRIDGE_ENV_VAR = "GOB_SP_BRIDGE_DIR"
BRIDGE_ROOT_HINT_FILENAME = "bridge_root.json"
MANIFEST_FILENAME = "bridge.json"
BLENDER_EXPORT_FILENAME = "b2sp.fbx"
SP_EXPORT_FILENAME = "sp2b.fbx"
LOG_FILENAME = "sp_export_log.txt"
ACTIVE_SP_INFO_FILENAME = "active_sp.json"
HIGH_POLY_RETRY_DELAY_MS = 800
HIGH_POLY_RETRY_COUNT = 60
UPDATE_URL = (
    "https://raw.githubusercontent.com/CIoudGuy/Blender-to-Substance-Painter-and-back-Gob/"
    "refs/heads/main/version.json"
)
PLUGIN_VERSION = "0.1.5"

EXPORT_FORMATS = [
    ("png", "PNG"),
    ("tga", "TGA"),
    ("tiff", "TIFF"),
    ("exr", "EXR"),
]
EXPORT_BIT_DEPTHS = [
    ("8", "8-bit"),
    ("16", "16-bit"),
    ("32", "32-bit"),
]
EXPORT_RESOLUTIONS = [
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


def bridge_root_hint_path():
    return Path(default_bridge_dir()).expanduser() / BRIDGE_ROOT_HINT_FILENAME


def read_bridge_root_hint():
    path = bridge_root_hint_path()
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


def load_persistent_state():
    data = load_settings()
    return {
        "version": data.get("version", SETTINGS_VERSION),
        "last_settings": data.get("last_settings", {}),
        "user_presets": data.get("user_presets", []),
    }


def save_persistent_state(last_settings=None, user_presets=None):
    data = load_settings()
    if not isinstance(data, dict):
        data = {}
    if last_settings is not None:
        data["last_settings"] = last_settings
    if user_presets is not None:
        data["user_presets"] = user_presets
    data["version"] = SETTINGS_VERSION
    save_settings(data)


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
        name = sp.project.name()
    except Exception:
        name = ""
    if not name:
        name = "untitled"
    return sanitize_name(name)


def get_project_dir():
    return get_bridge_root() / get_project_name()


def active_sp_info_path():
    return get_bridge_root() / ACTIVE_SP_INFO_FILENAME


def write_active_sp_info():
    info = {
        "timestamp": time.time(),
        "project_open": False,
    }
    try:
        if sp.project.is_open():
            info["project_open"] = True
            info["project_name"] = get_project_name()
            info["project_dir"] = str(get_project_dir())
    except Exception:
        pass
    path = active_sp_info_path()
    ensure_dir(path.parent)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(info, handle, indent=2, ensure_ascii=True)
    except OSError:
        return


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
        return
    box = QtWidgets.QMessageBox()
    box.setIcon(QtWidgets.QMessageBox.Information)
    box.setWindowTitle("GoB Bridge Update")
    box.setText(
        f"Update available: {info['version']} (current {info['local_version']})"
    )
    notes = info.get("notes")
    if notes:
        box.setInformativeText(str(notes))
    open_button = None
    if info.get("download_url"):
        open_button = box.addButton("Open Download", QtWidgets.QMessageBox.AcceptRole)
    box.addButton("Later", QtWidgets.QMessageBox.RejectRole)
    box.exec()
    if open_button and box.clickedButton() == open_button:
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(info["download_url"]))


def show_update_result(result, show_no_update=False):
    status = result.get("status") if result else None
    if status == "update":
        show_update_dialog(result.get("info"))
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
_update_net_manager = None
_update_reply = None
_update_timeout_timer = None
_update_status_kind = "idle"
_update_status_text = "Update: not checked yet"
_last_update_info = None
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


def start_update_check(show_no_update=False):
    global _update_check_in_progress
    global _update_check_started_at
    global _update_check_show_no_update
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
        show_update_result(result, show_no_update=_update_check_show_no_update)
        _update_check_show_no_update = False

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


def append_log(project_dir, message, data=None):
    try:
        log_path = Path(project_dir) / LOG_FILENAME
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
        texsets = list(sp.textureset.all_texture_sets())
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
    for texset in sp.textureset.all_texture_sets():
        stacks = texset.all_stacks()
        if stacks:
            for stack in stacks:
                root = texset.name()
                if stack.name():
                    root = f"{root}/{stack.name()}"
                entry = {"rootPath": root}
                if output_maps:
                    entry["filter"] = {"outputMaps": output_maps}
                export_list.append(entry)
        else:
            entry = {"rootPath": texset.name()}
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
        return sp.textureset.Stack.from_name(texset.name(), stack_name)
    except Exception:
        return None


def get_output_map_definitions(preset_info):
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
                    stack = sp.textureset.get_active_stack()
                    if not stack:
                        stacks = []
                        for texset in sp.textureset.all_texture_sets():
                            stacks.extend(texset.all_stacks())
                        stack = stacks[0] if stacks else None
                    if not stack:
                        return []
                    return preset.list_output_maps(stack)
        except Exception:
            return []
    return []


def get_output_map_names(preset_info):
    return extract_output_map_names(get_output_map_definitions(preset_info))


DOC_MAP_TO_CHANNEL = {
    "basecolor": "BaseColor",
    "base_color": "BaseColor",
    "albedo": "BaseColor",
    "diffuse": "BaseColor",
    "roughness": "Roughness",
    "glossiness": "Glossiness",
    "metallic": "Metallic",
    "metalness": "Metallic",
    "normal": "Normal",
    "height": "Height",
    "displacement": "Height",
    "opacity": "Opacity",
    "emissive": "Emissive",
    "emission": "Emissive",
    "specular": "Specular",
    "specularlevel": "SpecularLevel",
}


def stack_has_doc_map(stack, doc_map_name):
    if not doc_map_name:
        return True
    lookup = DOC_MAP_TO_CHANNEL.get(doc_map_name.lower())
    if not lookup:
        return True
    channel_type = sp.textureset.ChannelType.__members__.get(lookup)
    if not channel_type:
        return True
    return stack.has_channel(channel_type)


def get_required_doc_maps(map_def):
    required = set()
    if not isinstance(map_def, dict):
        return required
    channels = map_def.get("channels") or []
    for channel in channels:
        if channel.get("srcMapType") != "documentMap":
            continue
        name = channel.get("srcMapName")
        if name:
            required.add(name.lower())
    return required


def build_export_list_for_preset(preset_info, selected_output_maps, selected_texture_sets=None):
    export_list = []
    map_defs = get_output_map_definitions(preset_info)
    if not map_defs:
        return export_list
    selected_sets = None
    if selected_texture_sets:
        selected_sets = {name.lower() for name in selected_texture_sets if name}
    for texset in sp.textureset.all_texture_sets():
        if selected_sets and texset.name().lower() not in selected_sets:
            continue
        stacks = texset.all_stacks()
        if not stacks:
            stack = get_stack_for_textureset(texset, "")
            stacks = [stack] if stack else []
        for stack in stacks:
            if not stack:
                continue
            root = texset.name()
            if stack.name():
                root = f"{root}/{stack.name()}"
            valid_maps = []
            for map_def in map_defs:
                if isinstance(map_def, dict):
                    map_name = map_def.get("fileName")
                    if not map_name or map_name not in selected_output_maps:
                        continue
                    required = get_required_doc_maps(map_def)
                    missing = [req for req in required if not stack_has_doc_map(stack, req)]
                    if missing:
                        continue
                    valid_maps.append(map_name)
                elif isinstance(map_def, str):
                    if map_def in selected_output_maps:
                        valid_maps.append(map_def)
                else:
                    map_name = getattr(map_def, "fileName", None) or getattr(map_def, "file_name", None)
                    if map_name and map_name in selected_output_maps:
                        valid_maps.append(map_name)
            if not valid_maps and selected_output_maps:
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
    valid_formats = {value for value, _label in EXPORT_FORMATS}
    valid_bit_depths = {value for value, _label in EXPORT_BIT_DEPTHS}
    valid_resolutions = {value for value, _label in EXPORT_RESOLUTIONS}
    valid_padding = {value for value, _label in PADDING_ALGORITHMS}
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
            texsets = sp.textureset.all_texture_sets()
            if texsets:
                stacks = texsets[0].all_stacks()
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


class ExportDialog(QtWidgets.QDialog):
    def __init__(self):
        super().__init__(QtWidgets.QApplication.activeWindow())
        self.setWindowTitle("GoB Bridge - Send to Blender")
        self._presets = []
        self._all_presets = []
        self._user_presets = []
        self._pending_map_selection = None
        self._pending_texture_sets = None
        self._loading = True

        state = load_persistent_state()
        self._last_state = state.get("last_settings", {})
        self._user_presets = [
            preset for preset in state.get("user_presets", [])
            if isinstance(preset, dict) and preset.get("name")
        ]
        self._pending_map_selection = self._last_state.get("output_maps")
        self._pending_texture_sets = self._last_state.get("texture_sets")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(12, 12, 12, 12)
        self.setMinimumWidth(920)

        header = QtWidgets.QLabel("GoB Bridge Export")
        header.setStyleSheet("font-weight: 600; font-size: 14px;")
        layout.addWidget(header)

        update_bar = QtWidgets.QHBoxLayout()
        self.update_status_label = QtWidgets.QLabel()
        self.update_status_label.setText(_update_status_text)
        self.update_status_label.setStyleSheet("font-weight: 600;")
        self.update_check_btn = QtWidgets.QPushButton("Check Updates")
        self.update_download_btn = QtWidgets.QPushButton("Download")
        self.update_download_btn.setEnabled(False)
        update_bar.addWidget(self.update_status_label)
        update_bar.addStretch()
        update_bar.addWidget(self.update_check_btn)
        update_bar.addWidget(self.update_download_btn)
        update_widget = QtWidgets.QWidget()
        update_widget.setLayout(update_bar)
        layout.addWidget(update_widget)

        preset_bar = QtWidgets.QHBoxLayout()
        preset_bar.addWidget(QtWidgets.QLabel("Bridge preset"))
        self.user_preset_combo = QtWidgets.QComboBox()
        preset_bar.addWidget(self.user_preset_combo, 1)
        self.save_preset_btn = QtWidgets.QPushButton("Save")
        self.delete_preset_btn = QtWidgets.QPushButton("Delete")
        preset_bar.addWidget(self.save_preset_btn)
        preset_bar.addWidget(self.delete_preset_btn)
        layout.addLayout(preset_bar)

        self.mesh_group = QtWidgets.QGroupBox("Mesh Export")
        mesh_layout = QtWidgets.QVBoxLayout(self.mesh_group)
        self.mesh_cb = QtWidgets.QCheckBox("Export mesh (FBX)")
        self.mesh_cb.setChecked(False)
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
        mesh_layout.addLayout(mesh_form)
        layout.addWidget(self.mesh_group)

        self.texture_group = QtWidgets.QGroupBox("Texture Export")
        texture_layout = QtWidgets.QVBoxLayout(self.texture_group)
        self.textures_cb = QtWidgets.QCheckBox("Export textures")
        self.textures_cb.setChecked(False)
        texture_layout.addWidget(self.textures_cb)

        texture_split = QtWidgets.QHBoxLayout()
        texture_layout.addLayout(texture_split)

        self.texture_sets_group = QtWidgets.QGroupBox("Texture Sets")
        texset_layout = QtWidgets.QVBoxLayout(self.texture_sets_group)
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
        self.texset_list = QtWidgets.QListWidget()
        self.texset_list.setMinimumWidth(240)
        self.texset_list.setAlternatingRowColors(True)
        self.texset_list.setUniformItemSizes(True)
        self.texset_list.itemChanged.connect(self._update_texture_set_count)
        texset_layout.addWidget(self.texset_list)
        texture_split.addWidget(self.texture_sets_group, 1)

        self.texture_params_group = QtWidgets.QGroupBox("General Export Parameters")
        params_form = QtWidgets.QFormLayout(self.texture_params_group)

        self.output_dir_edit = QtWidgets.QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setText(str(get_project_dir() / "textures"))
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

        texture_split.addWidget(self.texture_params_group, 2)

        self.map_group = QtWidgets.QGroupBox("List of Exports")
        maps_layout = QtWidgets.QVBoxLayout(self.map_group)
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
        self.map_list = QtWidgets.QListWidget()
        self.map_list.setMinimumHeight(240)
        self.map_list.setAlternatingRowColors(True)
        self.map_list.setUniformItemSizes(True)
        self.map_list.itemChanged.connect(self._update_map_count)
        maps_layout.addWidget(self.map_list)
        texture_layout.addWidget(self.map_group)

        layout.addWidget(self.texture_group)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.textures_cb.toggled.connect(self._on_textures_toggle)
        self.mesh_cb.toggled.connect(self.mesh_combo.setEnabled)
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
        self.update_check_btn.clicked.connect(lambda: start_update_check(show_no_update=True))
        self.update_download_btn.clicked.connect(self._open_update_download)
        add_update_listener(self._refresh_update_status)

        self._populate_texture_sets()
        self._all_presets = collect_export_presets()
        self._filter_presets("")
        self._reload_user_presets()
        self._apply_saved_state(self._last_state)
        self._on_textures_toggle(self.textures_cb.isChecked())
        self._loading = False
        self._refresh_update_status()

    def closeEvent(self, event):
        remove_update_listener(self._refresh_update_status)
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

    def _open_update_download(self):
        if _last_update_info and _last_update_info.get("download_url"):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl(_last_update_info["download_url"]))

    def _on_textures_toggle(self, enabled):
        active = enabled and self.preset_combo.count() > 0
        self.preset_combo.setEnabled(active)
        self.preset_search.setEnabled(enabled)
        self.texture_sets_group.setEnabled(enabled)
        self.texture_params_group.setEnabled(enabled)
        self.map_group.setEnabled(enabled)
        if self._loading:
            return
        self._refresh_map_list()

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
        texsets = sp.textureset.all_texture_sets()
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
            name = texset.name()
            item = QtWidgets.QListWidgetItem(name)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
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
        preset_ref = state.get("preset")
        if preset_ref:
            self._select_preset_by_ref(preset_ref)
        self._pending_map_selection = state.get("output_maps")
        self._pending_texture_sets = state.get("texture_sets")
        self._apply_pending_texture_sets()
        self._refresh_map_list()

    def _serialize_options(self, options):
        preset = options.get("preset")
        preset_ref = None
        if preset:
            preset_ref = {"kind": preset.get("kind"), "name": preset.get("name")}
        mesh_option = options.get("mesh_option")
        mesh_key = mesh_option_key(mesh_option) if mesh_option is not None else None
        return {
            "export_mesh": options.get("export_mesh", True),
            "export_textures": options.get("export_textures", True),
            "mesh_option": mesh_key,
            "preset": preset_ref,
            "output_maps": options.get("output_maps", []),
            "export_settings": options.get("export_settings", {}),
            "texture_sets": options.get("texture_sets", []),
        }

    def _reload_user_presets(self, select_name=None):
        self.user_preset_combo.blockSignals(True)
        self.user_preset_combo.clear()
        self.user_preset_combo.addItem("Custom")
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

    def _apply_user_preset_selection(self, index):
        if self._loading or index <= 0:
            return
        preset = self._user_presets[index - 1]
        options = preset.get("options", {})
        self._loading = True
        self._apply_saved_state(options)
        self._on_textures_toggle(self.textures_cb.isChecked())
        self._loading = False

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

    def get_options(self):
        preset = None
        if self.preset_combo.isEnabled() and self._presets and self.preset_combo.currentIndex() >= 0:
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
            "mesh_option": self.mesh_combo.currentData(),
            "preset": preset,
            "output_maps": output_maps,
            "export_settings": export_settings,
            "texture_sets": texture_sets,
        }

    def persist_last_settings(self, options):
        state = self._serialize_options(options)
        save_persistent_state(last_settings=state, user_presets=self._user_presets)


class QuickPanel(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(2)
        self.import_btn = QtWidgets.QToolButton()
        self.export_btn = QtWidgets.QToolButton()
        self.update_btn = QtWidgets.QToolButton()
        self.import_btn.setText("GoB Import")
        self.export_btn.setText("GoB Export")
        self.update_btn.setText("Check Update")
        for btn in (self.import_btn, self.export_btn, self.update_btn):
            btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
            btn.setAutoRaise(True)
            btn.setCursor(QtCore.Qt.PointingHandCursor)
        self.import_btn.setToolTip("Import from Blender")
        self.export_btn.setToolTip("Send to Blender")
        self.update_btn.setToolTip("Check for updates")
        self.import_btn.clicked.connect(import_from_blender)
        self.export_btn.clicked.connect(send_to_blender)
        self.update_btn.clicked.connect(lambda: start_update_check(show_no_update=True))
        layout.addWidget(self.import_btn)
        layout.addWidget(self.export_btn)
        layout.addWidget(self.update_btn)


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


def import_from_blender(manifest_path=None, clear_auto_import=False):
    global _auto_import_in_progress
    bridge_roots = get_candidate_bridge_roots()
    project_dir = None
    manifest = None
    if manifest_path:
        manifest_path = Path(manifest_path)
        if manifest_path.exists():
            manifest = read_manifest(manifest_path)
            project_dir = manifest_path.parent
    else:
        project_dir = get_project_dir()
        manifest_path = project_dir / MANIFEST_FILENAME
        if manifest_path.exists():
            manifest = read_manifest(manifest_path)
        if not manifest or manifest.get("source") != "blender":
            latest = find_latest_manifest(bridge_roots, source="blender")
            if not latest:
                show_message("GoB Bridge", "No Blender export manifest found.", QtWidgets.QMessageBox.Warning)
                return
            manifest_path = latest
            manifest = read_manifest(manifest_path)
            project_dir = manifest_path.parent
    if not manifest:
        show_message("GoB Bridge", "Failed to read Blender export manifest.", QtWidgets.QMessageBox.Warning)
        return

    mesh_path = manifest.get("mesh_fbx")
    high_path = manifest.get("high_mesh_fbx")
    force_new_project = bool(manifest.get("force_new_project"))
    if not mesh_path:
        fallback = project_dir / BLENDER_EXPORT_FILENAME
        mesh_path = str(fallback) if fallback.exists() else None
    if not mesh_path or not Path(mesh_path).is_file():
        show_message("GoB Bridge", "Blender FBX file not found.", QtWidgets.QMessageBox.Warning)
        return

    if sp.project.is_open() and not force_new_project:
        if not sp.project.is_in_edition_state():
            return
        settings = sp.project.MeshReloadingSettings(
            import_cameras=False,
            preserve_strokes=True,
        )

        def _on_reload(status):
            global _auto_import_in_progress
            _auto_import_in_progress = False
            if status == sp.project.ReloadMeshStatus.SUCCESS:
                if high_path:
                    apply_high_poly_when_ready(high_path)
                if clear_auto_import:
                    clear_auto_import_flag(manifest_path, manifest)
                show_message("GoB Bridge", "Mesh reloaded from Blender.")
            else:
                show_message("GoB Bridge", "Mesh reload failed.", QtWidgets.QMessageBox.Warning)

        _auto_import_in_progress = True
        try:
            sp.project.reload_mesh(mesh_path, settings, _on_reload)
        except Exception as exc:
            _auto_import_in_progress = False
            message = str(exc).lower()
            if "busy" in message:
                global _auto_import_busy_until
                _auto_import_busy_until = time.time() + 2.0
                return
            show_message("GoB Bridge", f"Mesh reload failed: {exc}", QtWidgets.QMessageBox.Warning)
        return
        return

    if sp.project.is_open() and force_new_project:
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
        if high_path:
            apply_high_poly_when_ready(high_path)
        if clear_auto_import:
            clear_auto_import_flag(manifest_path, manifest)
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

    project_dir = get_project_dir()
    ensure_dir(project_dir)

    manifest = {
        "version": 1,
        "source": "substance_painter",
        "project": get_project_name(),
        "timestamp": time.time(),
    }
    exported_any = False
    texture_errors = []

    if options["export_mesh"]:
        mesh_path = project_dir / SP_EXPORT_FILENAME
        result = sp.export.export_mesh(str(mesh_path), options["mesh_option"])
        if result.status != sp.export.ExportStatus.Success:
            show_message("GoB Bridge", result.message, QtWidgets.QMessageBox.Warning)
        else:
            manifest["mesh_fbx"] = str(mesh_path)
            exported_any = True

    if options["export_textures"]:
        preset = options.get("preset") or pick_export_preset()
        if not preset:
            texture_errors.append("No export preset found.")
        else:
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
                export_settings = options.get("export_settings") or DEFAULT_EXPORT_SETTINGS
                export_params = build_export_parameters(export_settings)
                export_list = build_export_list_for_preset(preset, output_maps, texture_sets)
                if not export_list:
                    texture_errors.append("No matching maps found for the current texture sets.")
                else:
                    attempts = []
                    if preset["kind"] == "custom":
                        attempts.append(("custom", preset))
                    else:
                        attempts.append(("url", preset))
                        if preset["kind"] == "resource":
                            for fallback in collect_export_presets():
                                if (fallback["kind"] == "predefined" and
                                        fallback["name"].lower() == preset["name"].lower()):
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
                        for _key, files in export_result.textures.items():
                            textures.extend(files)
                        manifest["textures_dir"] = str(textures_dir)
                        manifest["textures"] = textures
                        exported_any = True
                        texture_errors = []
                        break
                    if not tried and not texture_errors:
                        texture_errors.append("No export preset available.")

    if not exported_any:
        show_message("GoB Bridge", "Nothing was exported.", QtWidgets.QMessageBox.Warning)
        return

    write_manifest(project_dir / MANIFEST_FILENAME, manifest)
    if texture_errors:
        details = "\n".join(texture_errors)
        show_message("GoB Bridge", f"Texture export failed:\n{details}", QtWidgets.QMessageBox.Warning)
    else:
        show_message("GoB Bridge", "Export complete. Use Blender to import.")


_ui_elements = []
_quick_panel_widget = None
_auto_import_timer = None
_auto_import_last_time = 0.0
_auto_import_last_path = None
_auto_import_in_progress = False
_auto_import_busy_until = 0.0


def start_plugin():
    action_import = QtGui.QAction("GoB Bridge: Import from Blender")
    action_import.triggered.connect(import_from_blender)
    sp.ui.add_action(sp.ui.ApplicationMenu.File, action_import)
    _ui_elements.append(action_import)

    action_send = QtGui.QAction("GoB Bridge: Send to Blender")
    action_send.triggered.connect(send_to_blender)
    sp.ui.add_action(sp.ui.ApplicationMenu.File, action_send)
    _ui_elements.append(action_send)

    action_update = QtGui.QAction("GoB Bridge: Check for Updates")
    action_update.triggered.connect(lambda: start_update_check(show_no_update=True))
    sp.ui.add_action(sp.ui.ApplicationMenu.File, action_update)
    _ui_elements.append(action_update)

    try:
        quick_element = _add_quick_panel_ui()
        if quick_element is not None:
            _ui_elements.append(quick_element)
    except Exception:
        pass
    start_update_check()
    write_active_sp_info()
    def _auto_import_poll():
        global _auto_import_last_time
        global _auto_import_last_path
        global _auto_import_in_progress
        global _auto_import_busy_until
        write_active_sp_info()
        if _auto_import_in_progress:
            return
        if _auto_import_busy_until and time.time() < _auto_import_busy_until:
            return
        manifest_path = find_latest_manifest(get_candidate_bridge_roots(), source="blender")
        if not manifest_path:
            return
        manifest = read_manifest(manifest_path)
        if not manifest or not manifest.get("auto_import"):
            return
        ts = manifest_timestamp(manifest, manifest_path)
        if (str(manifest_path) == _auto_import_last_path and ts <= _auto_import_last_time):
            return
        import_from_blender(manifest_path=manifest_path, clear_auto_import=True)
        if not _auto_import_in_progress:
            _auto_import_last_time = ts
            _auto_import_last_path = str(manifest_path)

    global _auto_import_timer
    try:
        _auto_import_timer = QtCore.QTimer()
        _auto_import_timer.setInterval(1500)
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
    write_active_sp_info()


if __name__ == "__main__":
    start_plugin()
