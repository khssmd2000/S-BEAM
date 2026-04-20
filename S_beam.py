# -*- coding: utf-8 -*-
__title__  = "TEMPLATE_S"
__author__ = "Assem Khalelova"

import os
import shutil
import datetime
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    FamilySymbol,
    FamilyInstance,
    View,
    ViewSheet,
    IndependentTag,
    OpenOptions,
    ModelPathUtils,
    Transaction,
    XYZ,
    Structure,
    Level,
    Line,
    StorageType,
    ElementId,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity
)

app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# TODO: fill in S beam family name, type name, template path and output folder
FAMILY_NAME   = "TODO_S_BEAM_FAMILY_NAME"
TYPE_NAME     = "TODO_S_BEAM_TYPE_NAME"
SINGLE_MODELS = r"TODO_S_BEAM_SINGLE_MODELS_FOLDER"
TEMPLATE_PATH = r"TODO_S_BEAM_TEMPLATE_PATH"
QTY_PARAM     = "MAG_Conteggio"
PREFIX_PARAM  = "Prefix_Mark"

SKIP_PARAM_NAMES = {
    "Family",
    "Family and Type",
    "Family Name",
    "Type",
    "Type Name",
    "Type Id",
    "Id",
    "Category",
    "Design Option",
    "Edited by",
    "Workset",
    "Phase Created",
    "Phase Demolished",
    "Reference Level",
    "Start Level Offset",
    "End Level Offset",
    "Cross-Section Rotation",
    "Start Extension",
    "End Extension",
    "yz Justification",
    "y Justification",
    "y Offset Value",
    "z Justification",
    "z Offset Value",
}

OST_FRAMING_ID = int(BuiltInCategory.OST_StructuralFraming)


# ── Failure handler ─────────────────────────────────────────────────────────────

class SilentFailureHandler(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        for msg in fa.GetFailureMessages():
            sev = msg.GetSeverity()
            if sev == FailureSeverity.Warning:
                fa.DeleteWarning(msg)
            elif sev == FailureSeverity.Error:
                try:
                    fa.DeleteWarning(msg)
                except:
                    try:
                        fa.ResolveFailure(msg)
                    except:
                        pass
        return FailureProcessingResult.Continue


# ── Read params from source ─────────────────────────────────────────────────────

def _read_param_value(p):
    if p.Definition is None:
        return None
    st  = p.StorageType
    val = None
    if st == StorageType.String:
        val = p.AsString()
        if val is None:
            val = ""
    elif st == StorageType.Integer:
        val = p.AsInteger()
    elif st == StorageType.Double:
        val = p.AsDouble()
    elif st == StorageType.ElementId:
        return None
    if val is None:
        return None
    return (str(st), val)


def read_all_instance_params(element):
    data = {}

    for p in element.Parameters:
        if p.IsReadOnly or p.Definition is None:
            continue
        name = p.Definition.Name
        if name in SKIP_PARAM_NAMES:
            continue
        result = _read_param_value(p)
        if result is not None:
            data[name] = result

    try:
        symbol = element.Symbol
        if symbol is not None:
            for p in symbol.Parameters:
                if p.Definition is None:
                    continue
                name = p.Definition.Name
                if name in SKIP_PARAM_NAMES:
                    continue
                if name in data:
                    continue
                result = _read_param_value(p)
                if result is not None:
                    data[name] = result
    except:
        pass

    return data


def read_cut_length(element):
    p = element.get_Parameter(BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH)
    if p is not None:
        return p.AsDouble()
    return None


def _read_beam_value(element, param_name):
    """Try instance params first, fall back to type (Symbol) params."""
    p = element.LookupParameter(param_name)
    if p is None and element.Symbol is not None:
        p = element.Symbol.LookupParameter(param_name)
    if p is None:
        return None
    st = p.StorageType
    if st == StorageType.Double:
        return p.AsDouble()
    elif st == StorageType.String:
        return p.AsString()
    elif st == StorageType.Integer:
        return p.AsInteger()
    return None


def read_volume_cls(element):
    return _read_beam_value(element, "MC CLS Sheet - SLG22")


def read_kg_totale(element):
    return _read_beam_value(element, "KG TOTALE Sheet - SLG22")


def read_proj_info(doc):
    """Read project info params in their native storage types."""
    info = doc.ProjectInformation
    data = {}
    for name in ("MAG_Nome_Commesa", "MAG_Numero_Commesa", "MAG_Nome_Cantiere"):
        p = info.LookupParameter(name)
        if p is None:
            continue
        st = p.StorageType
        val = None
        if st == StorageType.String:
            val = p.AsString()
        elif st == StorageType.Integer:
            val = p.AsInteger()
        elif st == StorageType.Double:
            val = p.AsDouble()
        if val is None or val == "":
            continue
        data[name] = (str(st), val)
    return data


# ── Collect from selection ──────────────────────────────────────────────────────

def collect_elements_by_mark_from_selection(doc, uidoc):
    counts      = {}
    reps        = {}
    param_data  = {}
    cut_lengths = {}
    volume_data = {}
    kg_data     = {}

    selection = uidoc.Selection.GetElementIds()
    print("  Selection count: {}".format(len(list(selection))))

    if not selection:
        print("ERROR: No elements selected. Please select beams first.")
        return counts, reps, param_data, cut_lengths, volume_data, kg_data

    for el_id in selection:
        el = doc.GetElement(el_id)
        if el is None:
            continue

        if not isinstance(el, FamilyInstance):
            continue

        try:
            cat_id = el.Category.Id.IntegerValue
            if cat_id != OST_FRAMING_ID:
                continue
        except:
            continue

        prefix_p = el.LookupParameter(PREFIX_PARAM)
        if prefix_p is None:
            try:
                fn = el.Symbol.Family.Name if el.Symbol else "?"
                print("  SKIP: '{}' has no {} param".format(fn, PREFIX_PARAM))
            except:
                print("  SKIP: element has no {} param".format(PREFIX_PARAM))
            continue

        prefix = prefix_p.AsString() if prefix_p.AsString() else ""

        mark_p = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        mark   = mark_p.AsString() if mark_p and mark_p.AsString() else ""

        key = prefix + mark
        counts[key] = counts.get(key, 0) + 1

        if key not in reps:
            reps[key] = el
            param_data[key] = read_all_instance_params(el)
            cl = read_cut_length(el)
            if cl is not None:
                cut_lengths[key] = cl
            vol = read_volume_cls(el)
            if vol is not None:
                volume_data[key] = vol
            kg = read_kg_totale(el)
            if kg is not None:
                kg_data[key] = kg

    return counts, reps, param_data, cut_lengths, volume_data, kg_data


# ── File handling ───────────────────────────────────────────────────────────────

def create_from_template(template_path, dest_path):
    dest_dir = os.path.dirname(dest_path)
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    shutil.copy2(template_path, dest_path)


def open_document(filepath):
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)
    open_opts  = OpenOptions()
    open_opts.DetachFromCentralOption = 0
    return app.OpenDocumentFile(model_path, open_opts)


# ── Core logic ──────────────────────────────────────────────────────────────────

def find_existing_beam(target_doc, family_name, type_name):
    all_beams = FilteredElementCollector(target_doc)\
        .OfCategory(BuiltInCategory.OST_StructuralFraming)\
        .OfClass(FamilyInstance)\
        .ToElements()

    print("  Template has {} structural framing elements.".format(len(all_beams)))

    for el in all_beams:
        try:
            fn = el.Symbol.Family.Name
            tn = el.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if fn == family_name and tn == type_name:
                return el
        except:
            continue

    if all_beams:
        try:
            fn = all_beams[0].Symbol.Family.Name
            tn = all_beams[0].Symbol.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            print("  Using fallback beam: '{}' : '{}'".format(fn, tn))
        except:
            print("  Using fallback beam.")
        return all_beams[0]

    return None


def get_beam_endpoints(target_el):
    loc   = target_el.Location
    curve = loc.Curve
    return curve.GetEndPoint(0), curve.GetEndPoint(1)


def set_beam_cut_length(target_el, desired_cut_length):
    loc   = target_el.Location
    curve = loc.Curve
    start = curve.GetEndPoint(0)
    end   = curve.GetEndPoint(1)
    direction = (end - start).Normalize()

    start_ext = 0.0
    end_ext   = 0.0
    se_p = target_el.LookupParameter("Start Extension")
    ee_p = target_el.LookupParameter("End Extension")
    if se_p:
        start_ext = se_p.AsDouble()
    if ee_p:
        end_ext = ee_p.AsDouble()

    needed_curve_length = desired_cut_length - start_ext - end_ext
    new_end  = start + direction.Multiply(needed_curve_length)
    new_line = Line.CreateBound(start, new_end)
    loc.Curve = new_line


def reposition_tags_for_beam(target_doc, target_el, old_start, old_end, new_start, new_end):
    """
    Reproject each tag head onto the new beam axis: preserve the perpendicular
    offset, scale the along-axis distance by (new_length / old_length).
    """
    old_len = old_end.DistanceTo(old_start)
    new_len = new_end.DistanceTo(new_start)
    if old_len < 1e-9 or new_len < 1e-9:
        return

    old_axis = (old_end - old_start).Normalize()
    new_axis = (new_end - new_start).Normalize()
    scale    = new_len / old_len

    target_int = target_el.Id.IntegerValue
    all_views  = FilteredElementCollector(target_doc).OfClass(View).ToElements()

    moved   = 0
    scanned = 0
    for view in all_views:
        try:
            if view.IsTemplate:
                continue
        except:
            continue

        try:
            tags = FilteredElementCollector(target_doc, view.Id)\
                .OfClass(IndependentTag)\
                .ToElements()
        except:
            continue

        for tag in tags:
            scanned += 1
            try:
                tagged_ids = list(tag.GetTaggedLocalElementIds())
            except:
                tagged_ids = []

            hits = False
            for tid in tagged_ids:
                try:
                    if tid.IntegerValue == target_int:
                        hits = True
                        break
                except:
                    continue
            if not hits:
                continue

            try:
                head   = tag.TagHeadPosition
                offset = head - old_start
                along  = offset.DotProduct(old_axis)
                perp   = offset - old_axis.Multiply(along)

                new_along = along * scale
                new_head  = new_start + new_axis.Multiply(new_along) + perp

                tag.TagHeadPosition = new_head
                moved += 1
            except Exception as e:
                print("    Could not move tag {}: {}".format(tag.Id, e))

    print("  Scanned {} tags, repositioned {}.".format(scanned, moved))


def apply_params(target_el, param_dict, target_doc):
    written = 0
    failed  = []

    for name, (st, val) in param_dict.items():
        tgt_p = target_el.LookupParameter(name)
        if tgt_p is None or tgt_p.IsReadOnly:
            continue
        try:
            if st == "String":
                tgt_p.Set(str(val))
            elif st == "Integer":
                tgt_p.Set(int(val))
            elif st == "Double":
                tgt_p.Set(float(val))
            else:
                continue
            written += 1
        except:
            failed.append((name, st, val))

    if failed:
        target_doc.Regenerate()
        for name, st, val in failed:
            tgt_p = target_el.LookupParameter(name)
            if tgt_p is None or tgt_p.IsReadOnly:
                continue
            try:
                if st == "String":
                    tgt_p.Set(str(val))
                elif st == "Integer":
                    tgt_p.Set(int(val))
                elif st == "Double":
                    tgt_p.Set(float(val))
                written += 1
            except:
                print("    Could not set: {}".format(name))

    print("  Transferred {} parameters.".format(written))


def apply_proj_info(single_doc, proj_data):
    """Write project info, casting to match the target parameter's storage type."""
    info    = single_doc.ProjectInformation
    written = 0
    for name, entry in proj_data.items():
        p = info.LookupParameter(name)
        if p is None or p.IsReadOnly:
            print("    SKIP project info '{}' (missing or read-only)".format(name))
            continue

        if isinstance(entry, tuple):
            src_st, val = entry
        else:
            src_st, val = "String", entry

        target_st = str(p.StorageType)
        try:
            if target_st == "String":
                p.Set(str(val))
            elif target_st == "Integer":
                p.Set(int(val))
            elif target_st == "Double":
                p.Set(float(val))
            else:
                print("    SKIP project info '{}' (unsupported type {})".format(name, target_st))
                continue
            written += 1
        except Exception as e:
            print("    Could not set project info '{}': {}".format(name, e))
    print("  Transferred {} project info params.".format(written))


def _set_sheet_param(sheet, param_name, value):
    p = sheet.LookupParameter(param_name)
    if p is None or p.IsReadOnly:
        return False
    try:
        st = str(p.StorageType)
        if st == "String":
            p.Set(str(value))
        elif st == "Double":
            p.Set(float(value))
        elif st == "Integer":
            p.Set(int(value))
        return True
    except:
        print("    Could not set {}".format(param_name))
        return False


def set_sheet_extra_params(single_doc, volume_cls, kg_totale):
    today_str = datetime.datetime.now().strftime("%d.%m.%y")
    sheets    = FilteredElementCollector(single_doc).OfClass(ViewSheet).ToElements()

    for sheet in sheets:
        p_date = sheet.LookupParameter("MAG_Data_Tavola")
        if p_date and not p_date.IsReadOnly:
            try:
                p_date.Set(today_str)
            except:
                print("    Could not set MAG_Data_Tavola")

        if volume_cls is not None:
            _set_sheet_param(sheet, "MAG_Volume_CLS", volume_cls)

        if kg_totale is not None:
            _set_sheet_param(sheet, "MAG_Peso", kg_totale)


def set_sheet_qty(single_doc, qty):
    sheets  = FilteredElementCollector(single_doc).OfClass(ViewSheet).ToElements()
    updated = False
    for sheet in sheets:
        p = sheet.LookupParameter(QTY_PARAM)
        if p is not None and not p.IsReadOnly:
            p.Set(qty)
            updated = True
    return updated


def process_mark(mark_key, qty, param_dict, cut_length, proj_data, volume_cls, kg_totale):
    filename = mark_key + ".rvt"
    filepath = os.path.join(SINGLE_MODELS, filename)

    if os.path.exists(filepath):
        print("  {} already exists — opening.".format(filename))
    else:
        create_from_template(TEMPLATE_PATH, filepath)
        print("  Created {} from template.".format(filename))

    single_doc = None
    t = None

    try:
        single_doc = open_document(filepath)

        target_el = find_existing_beam(single_doc, FAMILY_NAME, TYPE_NAME)
        if target_el is None:
            print("  ERROR: No beam found in template. Skipping {}.".format(filename))
            return

        print("  Found beam instance.")

        t = Transaction(single_doc, "Set Parameters & Qty")
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(SilentFailureHandler())
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)
        t.Start()

        if cut_length is not None:
            try:
                old_start, old_end = get_beam_endpoints(target_el)
                set_beam_cut_length(target_el, cut_length)
                single_doc.Regenerate()
                new_start, new_end = get_beam_endpoints(target_el)
                reposition_tags_for_beam(
                    single_doc, target_el, old_start, old_end, new_start, new_end
                )
                single_doc.Regenerate()
                print("  Set beam Cut Length.")
            except Exception as e:
                print("  WARNING: Could not set Cut Length: {}".format(e))

        apply_params(target_el, param_dict, single_doc)
        apply_proj_info(single_doc, proj_data)
        set_sheet_extra_params(single_doc, volume_cls, kg_totale)

        if set_sheet_qty(single_doc, qty):
            print("  Set {} = {}".format(QTY_PARAM, qty))
        else:
            print("  WARNING: '{}' not writable on sheets.".format(QTY_PARAM))

        t.Commit()
        single_doc.Save()
        print("  OK: {}".format(filename))

    except Exception as e:
        print("  ERROR in {}: {}".format(filename, e))
        try:
            if t is not None and t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except:
            pass

    finally:
        if single_doc is not None:
            single_doc.Close(False)


# ── Main ────────────────────────────────────────────────────────────────────────

def main():
    print("=== Transfer Qty & Create Shop Drawings (S beams) ===\n")

    if not os.path.exists(TEMPLATE_PATH):
        print("ERROR: Template not found:\n  {}".format(TEMPLATE_PATH))
        return

    proj_data = read_proj_info(doc)
    if proj_data:
        print("Project info from main model:")
        for k, v in sorted(proj_data.items()):
            print("  {} = {}".format(k, v))
    else:
        print("WARNING: No project info params found.")

    counts, reps, param_data, cut_lengths, volume_data, kg_data = \
        collect_elements_by_mark_from_selection(doc, uidoc)

    if not counts:
        print("No matching elements in selection.")
        return

    print("Found {} unique marks in selection:".format(len(counts)))
    for k, v in sorted(counts.items()):
        print("  {} = {} pcs".format(k, v))

    first_key = sorted(counts.keys())[0]
    pdata = param_data.get(first_key, {})
    print("\nSample params read from [{}]:".format(first_key))
    for name, (st, val) in sorted(pdata.items()):
        if name in (PREFIX_PARAM, "Mark", "Cut Length",
                    "Comments", "KG TOTALE Sheet - SLG22"):
            print("  {} = {} ({})".format(name, val, st))
    cl = cut_lengths.get(first_key)
    if cl is not None:
        print("  Cut Length (geometry) = {:.4f} ft".format(cl))
    vol = volume_data.get(first_key)
    if vol is not None:
        print("  MC CLS Sheet - SLG22 = {}".format(vol))
    kg = kg_data.get(first_key)
    if kg is not None:
        print("  KG TOTALE Sheet - SLG22 = {}".format(kg))

    print("\nProcessing...\n")
    success, skipped = 0, 0

    for mark_key, qty in sorted(counts.items()):
        print("[{}]".format(mark_key))
        pdata = param_data.get(mark_key, {})
        cl    = cut_lengths.get(mark_key)
        vol   = volume_data.get(mark_key)
        kg    = kg_data.get(mark_key)
        if not pdata:
            print("  SKIP: no parameter data.")
            skipped += 1
            continue
        try:
            process_mark(mark_key, qty, pdata, cl, proj_data, vol, kg)
            success += 1
        except Exception as e:
            print("  FATAL: {}".format(e))
            skipped += 1

    print("\n=== Done: {} processed, {} skipped ===".format(success, skipped))

main()