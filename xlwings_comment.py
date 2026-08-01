import os
from typing import List, Tuple, Dict, Any

import xlwings as xw


def _default_out_path(in_path: str) -> str:
    base, ext = os.path.splitext(in_path)
    return f"{base}_with_Note{ext}"


def _get_sheet_shapes_count(sheet: xw.Sheet) -> int:
    try:
        # COM shapes count (best indicator drawings are intact)
        return int(sheet.api.Shapes.Count)
    except Exception:
        # fallback
        try:
            return len(sheet.shapes)
        except Exception:
            return -1


def _delete_existing_comment(rng: xw.Range) -> None:
    """
    Always replace policy:
      - if legacy comment exists, delete it
      - then add a new one
    """
    try:
        # ClearComments removes legacy comments
        rng.api.ClearComments()
    except Exception:
        # If no comment exists, ignore
        pass


def _add_comment(rng: xw.Range, text: str, autosize: bool = True) -> None:
    rng.api.AddComment(text)
    try:
        if autosize:
            rng.api.Comment.Shape.TextFrame.AutoSize = True
    except Exception:
        pass


def insert_comments_batch(
    in_path: str,
    sheet_name: str,
    placements: List[Tuple[str, str]],
    out_path: str | None = None,
    make_visible: bool = False,
    autosize: bool = True,
) -> Dict[str, Any]:
    """
    Insert notes/comments via xlwings (COM) to preserve drawings/shapes.
    Always replaces existing notes.

    placements: [(cell_address, note_text), ...]
    """
    if not out_path:
        out_path = _default_out_path(in_path)

    app = None
    wb = None
    summary = {
        "in_path": in_path,
        "out_path": out_path,
        "sheet": sheet_name,
        "requested": len(placements),
        "written": 0,
        "errors": 0,
        "shapes_before": None,
        "shapes_after": None,
        "shapes_intact": None,
    }

    try:
        app = xw.App(visible=make_visible, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(in_path)
        sht = wb.sheets[sheet_name]

        shapes_before = _get_sheet_shapes_count(sht)
        summary["shapes_before"] = shapes_before

        for addr, note_text in placements:
            try:
                rng = sht.range(addr)
                _delete_existing_comment(rng)
                _add_comment(rng, note_text, autosize=autosize)
                summary["written"] += 1
            except Exception:
                summary["errors"] += 1

        wb.save(out_path)

        shapes_after = _get_sheet_shapes_count(sht)
        summary["shapes_after"] = shapes_after
        summary["shapes_intact"] = (shapes_before == shapes_after) if (shapes_before >= 0 and shapes_after >= 0) else None
        return summary

    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass
