bl_info = {
    "name": "GoB SP Bridge",
    "author": "Cloud Guy | cloud_was_taken on Discord",
    "version": (0, 1, 0),
    "blender": (4, 5, 0),
    "location": "View3D > Sidebar > GoB SP",
    "description": "Send FBX to Substance 3D Painter and import meshes/textures back",
    "category": "Import-Export",
}

import json
import os
import shutil
import subprocess
import time
from pathlib import Path

import bpy
from bpy.props import BoolProperty, FloatProperty, StringProperty
from bpy.types import AddonPreferences, Operator, Panel


BRIDGE_ENV_VAR = "GOB_SP_BRIDGE_DIR"
MANIFEST_FILENAME = "bridge.json"
BLENDER_EXPORT_FILENAME = "b2sp.fbx"
BLENDER_HIGH_FILENAME = "b2sp_hi.fbx"
ACTIVE_SP_INFO_FILENAME = "active_sp.json"
ACTIVE_SP_INFO_MAX_AGE = 120.0

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".tga", ".exr"}

MAP_KEYWORDS = [
    ("basecolor", "base_color"),
    ("base_color", "base_color"),
    ("albedo", "base_color"),
    ("diffuse", "base_color"),
    ("metallic", "metallic"),
    ("metalness", "metallic"),
    ("roughness", "roughness"),
    ("glossiness", "glossiness"),
    ("normal", "normal"),
    ("ambientocclusion", "ao"),
    ("occlusion", "ao"),
    ("opacity", "opacity"),
    ("alpha", "opacity"),
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


def collect_low_poly_objects(context, prefs):
    suffixes = parse_suffixes(getattr(prefs, "low_poly_suffixes", ""))
    if suffixes:
        candidates = [
            obj for obj in context.scene.objects
            if obj.type == "MESH" and is_name_with_suffix(obj.name, suffixes)
        ]
        if candidates:
            return candidates
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
    return get_bridge_root(prefs) / get_project_name(context)




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


def get_candidate_bridge_roots(prefs):
    roots = []
    env_path = os.environ.get(BRIDGE_ENV_VAR)
    if env_path:
        roots.append(Path(env_path))
    if prefs and prefs.bridge_dir:
        roots.append(Path(prefs.bridge_dir))
    docs = windows_documents_dir()
    if docs:
        roots.append(Path(docs) / "GoB_SP_Bridge")
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


def folder_size_bytes(path):
    if not path or not path.exists():
        return 0
    total = 0
    for dirpath, _dirnames, filenames in os.walk(path):
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


def format_bytes(value):
    size = float(value or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024.0 or unit == "TB":
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


def detect_map_type(stem_lower):
    for keyword, map_type in MAP_KEYWORDS:
        if keyword in stem_lower:
            return map_type, keyword
    return None, None




def guess_texture_set_name(stem, keyword):
    if not keyword:
        return stem
    lower = stem.lower()
    idx = lower.find(keyword)
    if idx <= 0:
        return stem
    base = stem[:idx].rstrip(" _-.")
    return base or stem


def gather_texture_paths(manifest):
    paths = []
    if not manifest:
        return paths
    if isinstance(manifest.get("textures"), list):
        paths.extend(manifest["textures"])
    textures_dir = manifest.get("textures_dir")
    if textures_dir:
        for ext in IMAGE_EXTS:
            for path in Path(textures_dir).glob(f"*{ext}"):
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
        stem = Path(path).stem
        map_type, keyword = detect_map_type(stem.lower())
        if not map_type:
            continue
        texset = guess_texture_set_name(stem, keyword)
        grouped.setdefault(texset, {})[map_type] = path
    return grouped


def load_image(path):
    try:
        return bpy.data.images.load(path, check_existing=True)
    except RuntimeError:
        return None


def build_material(mat, maps):
    mat["gob_bridge_material"] = True
    mat.use_nodes = True
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

    y = 300
    step = -220
    for map_type in ("base_color", "ao", "metallic", "roughness", "glossiness",
                     "normal", "height", "opacity", "emission"):
        if map_type not in maps:
            continue
        tex = nodes.new("ShaderNodeTexImage")
        tex.location = (-400, y)
        y += step
        image = load_image(maps[map_type])
        if not image:
            continue
        tex.image = image
        if map_type in {"normal", "roughness", "metallic", "ao", "height", "opacity", "glossiness"}:
            try:
                image.colorspace_settings.name = "Non-Color"
            except TypeError:
                pass
        if map_type == "base_color":
            base_node = tex
        elif map_type == "ao":
            ao_node = tex
        elif map_type == "metallic":
            metallic_node = tex
        elif map_type == "roughness":
            roughness_node = tex
        elif map_type == "glossiness":
            gloss_node = tex
        elif map_type == "normal":
            normal_node = tex
        elif map_type == "height":
            height_node = tex
        elif map_type == "opacity":
            opacity_node = tex
        elif map_type == "emission":
            emission_node = tex

    if base_node and ao_node:
        mix = nodes.new("ShaderNodeMixRGB")
        mix.blend_type = "MULTIPLY"
        mix.inputs["Fac"].default_value = 1.0
        mix.location = (-50, 200)
        links.new(base_node.outputs["Color"], mix.inputs["Color1"])
        links.new(ao_node.outputs["Color"], mix.inputs["Color2"])
        links.new(mix.outputs["Color"], principled.inputs["Base Color"])
    elif base_node:
        links.new(base_node.outputs["Color"], principled.inputs["Base Color"])

    if metallic_node:
        links.new(metallic_node.outputs["Color"], principled.inputs["Metallic"])

    if roughness_node:
        links.new(roughness_node.outputs["Color"], principled.inputs["Roughness"])
    elif gloss_node:
        invert = nodes.new("ShaderNodeInvert")
        invert.location = (-100, -260)
        links.new(gloss_node.outputs["Color"], invert.inputs["Color"])
        links.new(invert.outputs["Color"], principled.inputs["Roughness"])

    if normal_node:
        normal_map = nodes.new("ShaderNodeNormalMap")
        normal_map.location = (-50, -520)
        links.new(normal_node.outputs["Color"], normal_map.inputs["Color"])
        links.new(normal_map.outputs["Normal"], principled.inputs["Normal"])

    if height_node:
        disp = nodes.new("ShaderNodeDisplacement")
        disp.inputs["Scale"].default_value = 0.1
        disp.location = (200, -520)
        links.new(height_node.outputs["Color"], disp.inputs["Height"])
        links.new(disp.outputs["Displacement"], output.inputs["Displacement"])

    if emission_node:
        links.new(emission_node.outputs["Color"], principled.inputs["Emission"])

    if opacity_node:
        links.new(opacity_node.outputs["Color"], principled.inputs["Alpha"])
        mat.blend_method = "BLEND"
        mat.shadow_method = "HASHED"

    return mat


def get_or_build_material(name, maps):
    mat = bpy.data.materials.get(name)
    if mat and not mat.get("gob_bridge_material"):
        mat = None
    if not mat:
        mat = bpy.data.materials.new(name=name)
    return build_material(mat, maps)


def assign_material_to_object(obj, material, texset_name, all_groups):
    if obj.type != "MESH":
        return
    target_slot = None
    if obj.material_slots:
        for idx, slot in enumerate(obj.material_slots):
            if slot.material and slot.material.name.lower() == texset_name.lower():
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


def apply_textures_to_objects(objects, grouped):
    if not grouped:
        return
    materials = {}
    for texset, maps in grouped.items():
        mat_name = texset
        mat = get_or_build_material(mat_name, maps)
        materials[texset.lower()] = mat

    groups = list(materials.items())
    for obj in objects:
        if obj.type != "MESH":
            continue
        assigned = False
        for key, mat in materials.items():
            for slot in obj.material_slots:
                if slot.material and slot.material.name.lower() == key:
                    assign_material_to_object(obj, mat, slot.material.name, materials)
                    assigned = True
        if assigned:
            continue
        for key, mat in groups:
            if key in obj.name.lower():
                assign_material_to_object(obj, mat, mat.name, materials)
                assigned = True
                break
        if not assigned:
            assign_material_to_object(obj, groups[0][1], groups[0][1].name, materials)


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


def _mode_from_context(mode_name):
    if not mode_name:
        return "OBJECT"
    if mode_name.startswith("EDIT_"):
        return "EDIT"
    return mode_name


def auto_unwrap_objects(objects, only_missing=True):
    def build_override(obj):
        wm = bpy.context.window_manager
        for window in wm.windows:
            screen = window.screen
            for area in screen.areas:
                if area.type != "VIEW_3D":
                    continue
                for region in area.regions:
                    if region.type != "WINDOW":
                        continue
                    return {
                        "window": window,
                        "screen": screen,
                        "area": area,
                        "region": region,
                        "active_object": obj,
                        "object": obj,
                        "selected_objects": [obj],
                        "selected_editable_objects": [obj],
                    }
        return None

    selected = [obj for obj in bpy.context.selected_objects if object_is_valid(obj)]
    active = bpy.context.view_layer.objects.active
    prev_mode = bpy.context.mode
    target_mode = _mode_from_context(prev_mode)
    try:
        if target_mode != "OBJECT":
            bpy.ops.object.mode_set(mode="OBJECT")
    except RuntimeError:
        pass

    for obj in objects:
        if not object_is_valid(obj) or obj.type != "MESH":
            continue
        has_uvs = bool(obj.data.uv_layers)
        if only_missing and has_uvs:
            continue
        if not has_uvs:
            obj.data.uv_layers.new(name="UVMap")
        for sel in bpy.context.selected_objects:
            try:
                sel.select_set(False)
            except ReferenceError:
                continue
        try:
            obj.select_set(True)
        except ReferenceError:
            continue
        bpy.context.view_layer.objects.active = obj
        override = build_override(obj)
        try:
            angle_limit = 66.0
            if override:
                with bpy.context.temp_override(**override):
                    bpy.ops.object.mode_set(mode="EDIT")
                    bpy.ops.mesh.select_all(action="SELECT")
                    try:
                        bpy.ops.uv.smart_project(angle_limit=angle_limit, island_margin=0.02)
                    except RuntimeError:
                        bpy.ops.uv.unwrap(method="ANGLE_BASED", margin=0.02)
                    bpy.ops.object.mode_set(mode="OBJECT")
            else:
                bpy.ops.object.mode_set(mode="EDIT")
                bpy.ops.mesh.select_all(action="SELECT")
                try:
                    bpy.ops.uv.smart_project(angle_limit=angle_limit, island_margin=0.02)
                except RuntimeError:
                    bpy.ops.uv.unwrap(method="ANGLE_BASED", margin=0.02)
                bpy.ops.object.mode_set(mode="OBJECT")
        except RuntimeError:
            try:
                bpy.ops.object.mode_set(mode="OBJECT")
            except RuntimeError:
                pass

    for sel in bpy.context.selected_objects:
        try:
            sel.select_set(False)
        except ReferenceError:
            continue
    for obj in selected:
        try:
            obj.select_set(True)
        except ReferenceError:
            continue
    bpy.context.view_layer.objects.active = active
    try:
        if target_mode != "OBJECT":
            bpy.ops.object.mode_set(mode=target_mode)
    except RuntimeError:
        pass


def export_fbx_objects(filepath, objects, prefs=None, strip_uvs=False, unwrap_uvs=False):
    unwrap_uvs = bool(unwrap_uvs) and not strip_uvs
    export_objs = [obj for obj in objects if object_is_valid(obj) and obj.type == "MESH"]
    if not export_objs:
        return False
    temp_objects = []
    renamed_objects = []
    if strip_uvs or unwrap_uvs:
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
    if unwrap_uvs:
        auto_unwrap_objects(export_objs, only_missing=False)
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


def export_selected_fbx(filepath, prefs=None, strip_uvs=False, unwrap_uvs=False):
    return export_fbx_objects(
        filepath,
        bpy.context.selected_objects,
        prefs=prefs,
        strip_uvs=strip_uvs,
        unwrap_uvs=unwrap_uvs,
    )


def object_has_uvs(obj):
    if obj.type != "MESH":
        return False
    return bool(obj.data.uv_layers)


def collect_high_poly_objects(context, prefs, low_objects):
    candidates = collect_high_poly_candidates(context, prefs)
    if not low_objects:
        return candidates
    low_set = {obj.name for obj in low_objects}
    return [obj for obj in candidates if obj.name not in low_set]


def collect_high_poly_candidates(context, prefs):
    objects = []
    suffixes = parse_suffixes(getattr(prefs, "high_poly_suffixes", ""))
    if suffixes:
        for obj in context.scene.objects:
            if obj.type != "MESH":
                continue
            if is_name_with_suffix(obj.name, suffixes):
                objects.append(obj)
    if prefs and getattr(prefs, "high_poly_collection_name", None):
        name = prefs.high_poly_collection_name.strip()
        if name:
            collection = bpy.data.collections.get(name)
            if collection:
                objects.extend([obj for obj in collection.objects if obj.type == "MESH"])
    for obj in context.scene.objects:
        if obj.type == "MESH" and obj.get("gob_high_poly"):
            objects.append(obj)
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
        if env_path and Path(env_path).is_file():
            return env_path

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

    for base in filter(None, [Path(program_files) / "Adobe", Path(program_files_x86) / "Adobe"]):
        if not base.exists():
            continue
        for exe in base.rglob("*Painter*.exe"):
            name = exe.name.lower()
            if "painter" in name and "substance" in name:
                return str(exe)
    return None


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
    export_high_poly: BoolProperty(
        name="Export high poly if available",
        default=True,
    )
    export_low_poly: BoolProperty(
        name="Export low poly",
        default=True,
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
    high_poly_collection_name: StringProperty(
        name="High Poly Collection",
        description="Collection name to export as high poly",
        default="",
    )
    sp_auto_unwrap: BoolProperty(
        name="Auto-unwrap (Blender)",
        description="Generate UVs in Blender before export",
        default=False,
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
    ui_show_sp_settings: BoolProperty(
        name="Show Substance Settings",
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
        low_objects = collect_low_poly_objects(context, prefs)
        high_candidates = []
        if prefs and prefs.export_high_poly:
            high_candidates = collect_high_poly_candidates(context, prefs)
            if high_candidates and low_objects:
                high_names = {obj.name for obj in high_candidates}
                low_objects = [obj for obj in low_objects if obj.name not in high_names]
        if not low_objects and (not prefs or prefs.export_low_poly):
            self.report({"ERROR"}, "Select or name at least one low poly mesh")
            return {"CANCELLED"}

        auto_unwrap = bool(prefs and prefs.sp_auto_unwrap)

        if low_objects and any(not object_has_uvs(obj) for obj in low_objects):
            if not auto_unwrap:
                self.report({"ERROR"}, "Missing UVs: enable auto-unwrap or unwrap in Blender")
                return {"CANCELLED"}

        active_info = find_active_sp_project_info(prefs)
        project_dir = active_info["project_dir"] if active_info else get_project_dir(context, prefs)
        ensure_dir(project_dir)

        export_path = project_dir / BLENDER_EXPORT_FILENAME
        old_manifest = read_manifest(project_dir / MANIFEST_FILENAME)
        old_mesh = old_manifest.get("mesh_fbx") if old_manifest else None

        if not prefs or prefs.export_low_poly:
            if not low_objects:
                self.report({"ERROR"}, "Low poly export enabled but no meshes found")
                return {"CANCELLED"}
            strip_uvs = False
            unwrap_uvs = auto_unwrap
            export_fbx_objects(
                export_path,
                low_objects,
                prefs=prefs,
                strip_uvs=strip_uvs,
                unwrap_uvs=unwrap_uvs,
            )
        elif not old_mesh:
            self.report({"ERROR"}, "Low poly export disabled and no previous low mesh found")
            return {"CANCELLED"}

        high_export_path = None
        if prefs and prefs.export_high_poly:
            high_objects = high_candidates
            if high_objects:
                high_export_path = project_dir / BLENDER_HIGH_FILENAME
                export_fbx_objects(high_export_path, high_objects, prefs=prefs)

        manifest_path = project_dir / MANIFEST_FILENAME
        sp_running = is_sp_running()
        manifest = {
            "version": 1,
            "source": "blender",
            "project": get_project_name(context),
            "mesh_fbx": str(export_path) if (not prefs or prefs.export_low_poly) else old_mesh,
            "timestamp": time.time(),
        }
        manifest["auto_import"] = True
        manifest["auto_import_at"] = time.time()
        if high_export_path:
            manifest["high_mesh_fbx"] = str(high_export_path)
        write_manifest(manifest_path, manifest)

        if prefs and prefs.auto_launch_sp:
            sp_exe = find_sp_exe(prefs)
            if sp_exe and not sp_running:
                try:
                    subprocess.Popen([sp_exe])
                except OSError:
                    self.report({"WARNING"}, "Failed to launch Substance Painter")
            elif not sp_exe:
                self.report({"WARNING"}, "Substance Painter executable not found")

        self.report({"INFO"}, "Exported FBX for Substance Painter")
        return {"FINISHED"}


class GOB_OT_ImportFromSP(Operator):
    bl_idname = "gob_sp.import_from_substance_painter"
    bl_label = "Import from Substance Painter"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        prefs = get_prefs(context)
        project_dir = get_project_dir(context, prefs)
        manifest_path = project_dir / MANIFEST_FILENAME
        manifest = None
        if manifest_path.exists():
            manifest = read_manifest(manifest_path)
        if not manifest or manifest.get("source") != "substance_painter":
            roots = get_candidate_bridge_roots(prefs)
            latest = find_latest_manifest(roots, source="substance_painter")
            if not latest:
                self.report({"ERROR"}, "No Substance Painter bridge manifest found")
                return {"CANCELLED"}
            manifest_path = latest
            manifest = read_manifest(manifest_path)
        if not manifest:
            self.report({"ERROR"}, "Failed to read bridge manifest")
            return {"CANCELLED"}

        mesh_path = manifest.get("mesh_fbx")
        new_objects = []
        if mesh_path and Path(mesh_path).is_file():
            new_objects = import_fbx(mesh_path)

        texture_paths = gather_texture_paths(manifest)
        targets = new_objects
        if not targets:
            targets = [obj for obj in context.selected_objects if obj.type == "MESH"]
        if texture_paths and targets:
            grouped = group_textures(texture_paths)
            apply_textures_to_objects(targets, grouped)

        self.report({"INFO"}, "Imported assets from Substance Painter")
        return {"FINISHED"}


class GOB_OT_ClearCacheGlobal(Operator):
    bl_idname = "gob_sp.clear_cache_global"
    bl_label = "Clear Global Cache"

    def execute(self, context):
        prefs = get_prefs(context)
        root = get_bridge_root(prefs)
        if not root.exists():
            self.report({"INFO"}, "Global cache is already empty")
            return {"FINISHED"}
        try:
            shutil.rmtree(root)
        except OSError:
            self.report({"WARNING"}, "Failed to clear global cache")
            return {"CANCELLED"}
        ensure_dir(root)
        self.report({"INFO"}, "Global cache cleared")
        return {"FINISHED"}

    def invoke(self, context, _event):
        return context.window_manager.invoke_confirm(self, _event)


class GOB_OT_ClearCacheLocal(Operator):
    bl_idname = "gob_sp.clear_cache_local"
    bl_label = "Clear Project Cache"

    def execute(self, context):
        prefs = get_prefs(context)
        root = get_project_dir(context, prefs)
        if not root.exists():
            self.report({"INFO"}, "Project cache is already empty")
            return {"FINISHED"}
        try:
            shutil.rmtree(root)
        except OSError:
            self.report({"WARNING"}, "Failed to clear project cache")
            return {"CANCELLED"}
        self.report({"INFO"}, "Project cache cleared")
        return {"FINISHED"}

    def invoke(self, context, _event):
        return context.window_manager.invoke_confirm(self, _event)


class GOB_OT_OpenDiscord(Operator):
    bl_idname = "gob_sp.open_discord"
    bl_label = "Join Discord"

    def execute(self, _context):
        bpy.ops.wm.url_open(url=DISCORD_INVITE_URL)
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
        row = layout.row(align=True)
        row.operator(GOB_OT_SendToSP.bl_idname, icon="EXPORT")
        row.operator(GOB_OT_ImportFromSP.bl_idname, icon="IMPORT")
        if prefs:
            export_box = layout.box()
            row = export_box.row()
            icon = "TRIA_DOWN" if prefs.ui_show_export_settings else "TRIA_RIGHT"
            row.prop(prefs, "ui_show_export_settings", icon=icon, emboss=False, text="Export Options")
            if prefs.ui_show_export_settings:
                col = export_box.column(align=True)
                col.prop(prefs, "export_low_poly")
                col.prop(prefs, "export_high_poly")
                if prefs.export_high_poly:
                    col.prop(prefs, "low_poly_suffixes")
                    col.prop(prefs, "high_poly_suffixes")
                    col.prop(prefs, "high_poly_collection_name")
                    info = export_box.box()
                    info.label(text="Name meshes with your suffixes", icon="INFO")
                    info.label(text=f"Low: ends with {prefs.low_poly_suffixes or '_low'}")
                    info.label(text=f"High: ends with {prefs.high_poly_suffixes or '_high'}")
                elif prefs.export_low_poly:
                    info = export_box.box()
                    info.label(text="Low poly must end with suffix (low-only too)", icon="INFO")
                    info.label(text=f"Low: ends with {prefs.low_poly_suffixes or '_low'}")

            sp_box = layout.box()
            row = sp_box.row()
            icon = "TRIA_DOWN" if prefs.ui_show_sp_settings else "TRIA_RIGHT"
            row.prop(prefs, "ui_show_sp_settings", icon=icon, emboss=False, text="Substance Painter")
            if prefs.ui_show_sp_settings:
                col = sp_box.column(align=True)
                col.prop(prefs, "sp_auto_unwrap")
                col.label(text="Auto-unwrap runs in Blender before export")

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
            row = cache_box.row()
            icon = "TRIA_DOWN" if prefs.ui_show_cache else "TRIA_RIGHT"
            row.prop(prefs, "ui_show_cache", icon=icon, emboss=False, text="Cache")
            if prefs.ui_show_cache:
                global_size = bridge_cache_size_bytes(prefs)
                local_size = project_cache_size_bytes(context, prefs)
                cache_box.label(text=f"Global cache: {format_bytes(global_size)}")
                cache_box.label(text=f"Project cache: {format_bytes(local_size)}")
                row = cache_box.row(align=True)
                row.operator(GOB_OT_ClearCacheGlobal.bl_idname, icon="TRASH")
                row.operator(GOB_OT_ClearCacheLocal.bl_idname, icon="TRASH")

            links = layout.box()
            links.label(text="Community")
            links.operator(GOB_OT_OpenDiscord.bl_idname, icon="URL")


classes = (
    GOBSPPreferences,
    GOB_OT_SendToSP,
    GOB_OT_ImportFromSP,
    GOB_OT_ClearCacheGlobal,
    GOB_OT_ClearCacheLocal,
    GOB_OT_OpenDiscord,
    GOB_PT_Panel,
)


def register():
    for cls in classes:
        bpy.utils.register_class(cls)


def unregister():
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
