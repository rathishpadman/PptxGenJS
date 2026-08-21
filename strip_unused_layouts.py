#!/usr/bin/env python3
"""strip_unused_layouts.py — drop slide layouts and masters no slide uses.

The HCLTech base template carries 111 layouts and ~13MB of embedded brand
imagery. A finished deck uses a handful. Removing the rest keeps the file
small and stops recipients browsing layouts that were never part of the
deliverable.

Run this on every deck before sharing it:

    python3 scripts/strip_unused_layouts.py deck.pptx            # in place
    python3 scripts/strip_unused_layouts.py deck.pptx --out b.pptx
    python3 scripts/strip_unused_layouts.py deck.pptx --dry-run

Layouts still referenced by a slide are always kept, as is at least one
master. Dropping a layout's relationship orphans its part, so python-pptx
leaves it out on save along with any images only that layout used.
"""
import argparse
import os
import shutil
import sys

from pptx import Presentation


def strip(path, out=None, dry_run=False, quiet=False):
    prs = Presentation(path)
    used = {s.slide_layout.part for s in prs.slides}

    before_layouts = sum(len(m.slide_layouts) for m in prs.slide_masters)
    before_masters = len(prs.slide_masters)
    removed_layouts = []

    for master in prs.slide_masters:
        id_lst = master._element.get_or_add_sldLayoutIdLst()
        for entry in list(id_lst):
            layout_part = master.part.related_part(entry.rId)
            if layout_part in used:
                continue
            removed_layouts.append(layout_part.partname)
            id_lst.remove(entry)
            master.part.drop_rel(entry.rId)

    # A master with no surviving layouts is dead weight -- but never remove
    # the last one, or the deck loses its theme.
    master_id_lst = prs._element.get_or_add_sldMasterIdLst()
    removed_masters = []
    for entry in list(master_id_lst):
        if len(master_id_lst) <= 1:
            break
        master_part = prs.part.related_part(entry.rId)
        if len(master_part.slide_layouts):
            continue
        removed_masters.append(master_part.partname)
        master_id_lst.remove(entry)
        prs.part.drop_rel(entry.rId)

    kept = before_layouts - len(removed_layouts)
    if not quiet:
        print('%s: layouts %d -> %d, masters %d -> %d'
              % (os.path.basename(path), before_layouts, kept,
                 before_masters, before_masters - len(removed_masters)))

    if dry_run:
        return kept, len(removed_layouts), len(removed_masters)

    target = out or path
    prs.save(target)
    if not quiet:
        print('  saved %s (%.1f KB)' % (target, os.path.getsize(target) / 1024))
    return kept, len(removed_layouts), len(removed_masters)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pptx', nargs='+', help='deck(s) to strip')
    ap.add_argument('--out', help='write here instead of in place (single input only)')
    ap.add_argument('--backup', action='store_true',
                    help='keep the original alongside as <name>.orig.pptx')
    ap.add_argument('--dry-run', action='store_true', help='report only')
    args = ap.parse_args()

    if args.out and len(args.pptx) > 1:
        ap.error('--out takes a single input file')

    for path in args.pptx:
        if args.backup and not args.dry_run:
            shutil.copyfile(path, os.path.splitext(path)[0] + '.orig.pptx')
        strip(path, out=args.out, dry_run=args.dry_run)
    return 0


if __name__ == '__main__':
    sys.exit(main())
