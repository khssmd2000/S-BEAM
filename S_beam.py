# -*- coding: utf-8 -*-
__title__ = "TEMPLATE_S_beam"
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
    Viewport,
    ViewType,
    IndependentTag,
    OpenOptions,
    ModelPathUtils,
    Transaction,
    XYZ,
    BoundingBoxXYZ,
    Structure,
    Level,
    Line,
    StorageType,
    ElementId,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity
)

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# --- Configuration ---
FAMILY_NAME = "MAG_S_B_50_SLG23_KLZ_B"
TYPE_NAME = "S120B"
SINGLE_MODELS = r"X:\01_PROJECTS\01_ITALY\01_MAGNETTI\MG-2608-SD-AKNO FIORENZUOLA\04_Designs\3_SD_Structure\2_Beams"
TEMPLATE_PATH = r"X:\01_PROJECTS\01_ITALY\01_MAGNETTI\MG-2608-SD-AKNO FIORENZUOLA\05_Model\Template\S_template.rvt"
QTY_PARAM = "MAG_Conteggio"
PREFIX_PARAM = "Prefix_Mark"

CROP_MARGIN_FT = 1.5
MAX_VIEWPORT_WIDTH_MM = 260.0
MAX_VIEWPORT_HEIGHT_MM = 180.0
VIEW_SCALES = [10, 20, 25, 50, 75, 100, 125, 150, 200, 250, 500]

SKIP_PARAM_NAMES = {
    "Family", "Family and Type", "Family Name", "Type Name", "Type Id", "Id",
    "Category", "Design Option", "Edited by", "Workset", "Phase Created",
    "Phase Demolished", "Reference Level", "Start Level Offset", "End Level Offset",
    "Cross-Section Rotation", "Start Extension", "End Extension", "yz Justification",
    "y Justification", "y Offset Value", "z Justification", "z Offset Value"
}

OST_FRAMING_ID = ElementId(BuiltInCategory.OST_StructuralFraming)

# --- Forceful Failure Handler for Dimension Errors ---
class SilentFailureHandler(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        failures = fa.GetFailureMessages()
        if not failures:
            return FailureProcessingResult.Continue
        
        for f in failures:
            severity = f.GetSeverity()
            if severity == FailureSeverity.Warning:
                fa.DeleteWarning(f)
            elif severity == FailureSeverity.Error:
                # This fixes the "Dimension reference invalid" block
                if f.HasDefaultResolution():
                    fa.ResolveFailure(f)
                else:
                    try:
                        fa.ResolveFailure(f)
                    except:
                        pass
        # ProceedWithCommit tells Revit to force the changes through even if elements were deleted
        return FailureProcessingResult.ProceedWithCommit

# --- Utility Functions ---
def unlock_view(view):
    if view.ViewTemplateId != ElementId.InvalidElementId:
        view.ViewTemplateId = ElementId.InvalidElementId

def _read_param_value(p):
    if p.Definition is None or not p.HasValue: return None
    st = p.StorageType
    if st == StorageType.String: return ("String", p.AsString() or "")
    if st == StorageType.Integer: return ("Integer", p.AsInteger())
    if st == StorageType.Double: return ("Double", p.AsDouble())
    return None

def read_all_instance_params(element):
    data = {}
    for p in element.Parameters:
        if p.IsReadOnly or p.Definition is None: continue
        if p.Definition.Name in SKIP_PARAM_NAMES: continue
        res = _read_param_value(p)
        if res: data[p.Definition.Name] = res
    symbol = element.Symbol
    if symbol:
        for p in symbol.Parameters:
            if p.Definition.Name in SKIP_PARAM_NAMES or p.Definition.Name in data: continue
            res = _read_param_value(p)
            if res: data[p.Definition.Name] = res
    return data

def read_proj_info(doc):
    info = doc.ProjectInformation
    data = {}
    for name in ("MAG_Nome_Commesa", "MAG_Numero_Commesa", "MAG_Nome_Cantiere"):
        p = info.LookupParameter(name)
        if p:
            res = _read_param_value(p)
            if res: data[name] = res
    return data

def collect_elements_by_mark_from_selection(doc, uidoc):
    counts, reps, param_data, cut_lengths, volume_data, kg_data = {}, {}, {}, {}, {}, {}
    selection = uidoc.Selection.GetElementIds()
    if not selection: return counts, reps, param_data, cut_lengths, volume_data, kg_data
    for el_id in selection:
        el = doc.GetElement(el_id)
        if not isinstance(el, FamilyInstance) or el.Category.Id != OST_FRAMING_ID: continue
        prefix_p = el.LookupParameter(PREFIX_PARAM)
        mark_p = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        prefix = prefix_p.AsString() if prefix_p and prefix_p.AsString() else ""
        mark = mark_p.AsString() if mark_p and mark_p.AsString() else ""
        key = prefix + mark
        counts[key] = counts.get(key, 0) + 1
        if key not in reps:
            reps[key] = el
            param_data[key] = read_all_instance_params(el)
            cl_p = el.get_Parameter(BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH)
            if cl_p: cut_lengths[key] = cl_p.AsDouble()
            vol_p = el.LookupParameter("MC CLS Sheet - SLG22")
            if vol_p: volume_data[key] = vol_p.AsDouble() if vol_p.StorageType == StorageType.Double else vol_p.AsString()
            kg_p = el.LookupParameter("KG TOTALE Sheet - SLG22")
            if kg_p: kg_data[key] = kg_p.AsDouble() if kg_p.StorageType == StorageType.Double else kg_p.AsString()
    return counts, reps, param_data, cut_lengths, volume_data, kg_data

# --- Geometry Functions ---
def set_beam_cut_length(target_el, desired_cut_length):
    loc = target_el.Location
    curve = loc.Curve
    start, end = curve.GetEndPoint(0), curve.GetEndPoint(1)
    direction = (end - start).Normalize()
    se_p = target_el.LookupParameter("Start Extension")
    ee_p = target_el.LookupParameter("End Extension")
    s_ext = se_p.AsDouble() if se_p else 0.0
    e_ext = ee_p.AsDouble() if ee_p else 0.0
    needed_len = desired_cut_length - s_ext - e_ext
    if needed_len < 0.01: needed_len = 0.01
    new_end = start + direction.Multiply(needed_len)
    loc.Curve = Line.CreateBound(start, new_end)

def reposition_tags_for_beam(target_doc, target_el, old_start, old_end, new_start, new_end):
    old_len, new_len = old_end.DistanceTo(old_start), new_end.DistanceTo(new_start)
    if old_len < 1e-7 or new_len < 1e-7: return
    old_axis, new_axis = (old_end - old_start).Normalize(), (new_end - new_start).Normalize()
    scale, target_val = new_len / old_len, target_el.Id.Value
    tags = FilteredElementCollector(target_doc).OfClass(IndependentTag).ToElements()
    for tag in tags:
        try:
            if any(tid.Value == target_val for tid in tag.GetTaggedLocalElementIds()):
                head = tag.TagHeadPosition
                offset = head - old_start
                along = offset.DotProduct(old_axis)
                perp = offset - old_axis.Multiply(along)
                tag.TagHeadPosition = new_start + new_axis.Multiply(along * scale) + perp
        except: continue

def expand_view_crops_for_beam(target_doc, b_start, b_end, margin=CROP_MARGIN_FT):
    for view in FilteredElementCollector(target_doc).OfClass(View).ToElements():
        if view.IsTemplate or not view.CropBoxActive: continue
        unlock_view(view)
        crop = view.CropBox
        t_inv = crop.Transform.Inverse
        p1, p2 = t_inv.OfPoint(b_start), t_inv.OfPoint(b_end)
        mid_x, mid_y = (p1.X + p2.X)/2.0, (p1.Y + p2.Y)/2.0
        dx, dy = abs(p2.X - p1.X), abs(p2.Y - p1.Y)
        new_bbox = BoundingBoxXYZ()
        new_bbox.Transform = crop.Transform
        if dx >= dy:
            new_bbox.Min = XYZ(mid_x - dx/2 - margin, crop.Min.Y, crop.Min.Z)
            new_bbox.Max = XYZ(mid_x + dx/2 + margin, crop.Max.Y, crop.Max.Z)
        else:
            new_bbox.Min = XYZ(crop.Min.X, mid_y - dy/2 - margin, crop.Min.Z)
            new_bbox.Max = XYZ(crop.Max.X, mid_y + dy/2 + margin, crop.Max.Z)
        try: view.CropBox = new_bbox
        except: pass

def auto_adjust_view_scales(target_doc, beam_length_ft):
    for vp in FilteredElementCollector(target_doc).OfClass(Viewport).ToElements():
        view = target_doc.GetElement(vp.ViewId)
        if not view or view.IsTemplate: continue
        unlock_view(view)
        if "ALTO" in view.Name.upper() or view.ViewType == ViewType.FloorPlan:
            req_scale = (beam_length_ft * 304.8) / (MAX_VIEWPORT_WIDTH_MM * 0.85)
            view.Scale = int(next((s for s in VIEW_SCALES if s >= req_scale), VIEW_SCALES[-1]))
        sheet = target_doc.GetElement(vp.SheetId)
        vp.SetBoxCenter(XYZ((sheet.Outline.Min.U + sheet.Outline.Max.U)/2.0, (sheet.Outline.Min.V + sheet.Outline.Max.V)/2.0, 0))

# --- Application Functions ---
def apply_params(target_el, param_dict):
    sym = target_el.Symbol
    for name, (st, val) in param_dict.items():
        p = target_el.LookupParameter(name)
        if (p is None or p.IsReadOnly) and sym: p = sym.LookupParameter(name)
        if p and not p.IsReadOnly:
            try:
                if st == "String": p.Set(str(val))
                elif st == "Integer": p.Set(int(val))
                elif st == "Double": p.Set(float(val))
            except: continue

def process_mark(mark_key, qty, param_dict, cut_length, proj_data, vol, kg):
    filepath = os.path.join(SINGLE_MODELS, mark_key + ".rvt")
    if not os.path.exists(filepath):
        if not os.path.exists(os.path.dirname(filepath)): os.makedirs(os.path.dirname(filepath))
        shutil.copy2(TEMPLATE_PATH, filepath)
    
    single_doc = app.OpenDocumentFile(ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath), OpenOptions())
    try:
        beams = FilteredElementCollector(single_doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(FamilyInstance).ToElements()
        target_el = next((b for b in beams if b.Symbol.Family.Name == FAMILY_NAME), beams[0] if beams else None)
        if target_el:
            with Transaction(single_doc, "Update Beam and Graphics") as t:
                opts = t.GetFailureHandlingOptions()
                opts.SetFailuresPreprocessor(SilentFailureHandler()) # Dimension error solver
                t.SetFailureHandlingOptions(opts)
                t.Start()
                
                # 1. Params FIRST (Sets extensions)
                apply_params(target_el, param_dict)
                single_doc.Regenerate()
                
                # 2. Geometry SECOND
                if cut_length:
                    old_s, old_e = target_el.Location.Curve.GetEndPoint(0), target_el.Location.Curve.GetEndPoint(1)
                    set_beam_cut_length(target_el, cut_length)
                    single_doc.Regenerate() # This triggers and resolves dimension errors
                    new_s, new_e = target_el.Location.Curve.GetEndPoint(0), target_el.Location.Curve.GetEndPoint(1)
                    reposition_tags_for_beam(single_doc, target_el, old_s, old_e, new_s, new_e)
                    expand_view_crops_for_beam(single_doc, new_s, new_e)
                    auto_adjust_view_scales(single_doc, cut_length)

                # 3. Project and Sheet Info
                info = single_doc.ProjectInformation
                for name, (st, val) in proj_data.items():
                    p = info.LookupParameter(name)
                    if p: p.Set(val)
                
                for sheet in FilteredElementCollector(single_doc).OfClass(ViewSheet).ToElements():
                    pq = sheet.LookupParameter(QTY_PARAM)
                    if pq: pq.Set(qty)
                    pv = sheet.LookupParameter("MAG_Volume_CLS")
                    if pv and vol: pv.Set(vol)
                    pk = sheet.LookupParameter("MAG_Peso")
                    if pk and kg: pk.Set(kg)
                    pd = sheet.LookupParameter("MAG_Data_Tavola")
                    if pd: pd.Set(datetime.datetime.now().strftime("%d.%m.%y"))

                t.Commit()
            single_doc.Save()
            print("OK: {}".format(mark_key))
    finally:
        single_doc.Close(False)

def main():
    print("=== Processing Beams ===\n")
    proj_data = read_proj_info(doc)
    counts, reps, param_data, cut_lengths, volume_data, kg_data = collect_elements_by_mark_from_selection(doc, uidoc)
    for key, qty in sorted(counts.items()):
        print("Processing: {} ({} pcs)".format(key, qty))
        process_mark(key, qty, param_data[key], cut_lengths.get(key), proj_data, volume_data.get(key), kg_data.get(key))

if __name__ == "__main__":
    main()